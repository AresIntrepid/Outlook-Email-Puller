# Outlook Inbox export — full history on first run, incremental after that.
# Silent version — no console output, no prompts.
# Works for any user - paths are relative to their own profile.
#
# Maintains TWO output files:
#   emails_history.txt  - permanent, cumulative, NEVER trimmed. Full archive, in case you
#                          ever want to look back further than the recent window below.
#   emails_recent.txt   - rolling window, trimmed every run to the last $recentWindowDays
#                          days. This is the small, fast file the daily/weekly summary
#                          automations should read from — it never grows unbounded.

$recentWindowDays = 10   # covers the 7-day weekly summary plus a few days' buffer for a missed run

$outDir     = Join-Path $env:USERPROFILE "weekly report"
if (-not (Test-Path $outDir)) {
    New-Item -ItemType Directory -Path $outDir | Out-Null
}
$historyFile = Join-Path $outDir "emails_history.txt"   # cumulative archive, appended to over time, never trimmed
$recentFile  = Join-Path $outDir "emails_recent.txt"    # rolling window for the summary automations to read
$stateFile   = Join-Path $outDir "last_sync_state.txt"  # remembers the newest ReceivedTime we've saved
$logFile     = Join-Path $outDir "pull_emails_log.txt"

function Write-Log($msg) {
    Add-Content -Path $logFile -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - $msg"
}

# --- Self-install: the very first time this script runs on a machine, register it with Windows
#     Task Scheduler so it keeps running automatically from then on. Safe to leave in place
#     permanently - it checks whether the task already exists and does nothing if so, so
#     double-clicking this script again later (e.g. to force a manual pull) never creates a
#     duplicate task. This runs before the Outlook connection attempt below, so the recurring
#     task gets set up even if Outlook isn't ready yet at this exact moment. ---
$taskName = "PullEmailsAutoSync"
try {
    $existingTask = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
    if (-not $existingTask) {
        # Figure out what's actually running us: a plain .ps1 under powershell.exe/pwsh.exe, or a
        # ps2exe-compiled standalone .exe. $PSCommandPath is unreliable once ps2exe wraps the script
        # (it can point at an extracted temp file instead of the real .exe), so ask the OS directly
        # what process is running via the current process's own module path.
        $hostExePath = [System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
        if ($hostExePath -match '\\(powershell|pwsh)\.exe$') {
            # Running as a plain .ps1 - point the task at powershell.exe + this script's path.
            $action = New-ScheduledTaskAction -Execute "powershell.exe" `
                -Argument "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$PSCommandPath`""
        } else {
            # Running as a compiled exe (e.g. ps2exe) - point the task straight at the exe itself.
            $action = New-ScheduledTaskAction -Execute $hostExePath
        }

        # "Repeat every 1 hour for 18 hours, once a day starting at 6am" - built via a throwaway
        # one-time trigger so we can lift its Repetition settings onto the real daily trigger.
        $repeatSeed   = New-ScheduledTaskTrigger -Once -At (Get-Date -Hour 6 -Minute 0 -Second 0) `
            -RepetitionInterval (New-TimeSpan -Hours 1) -RepetitionDuration (New-TimeSpan -Hours 18)
        $dailyTrigger = New-ScheduledTaskTrigger -Daily -At (Get-Date -Hour 6 -Minute 0 -Second 0)
        $dailyTrigger.Repetition = $repeatSeed.Repetition

        $settings  = New-ScheduledTaskSettingsSet -WakeToRun -StartWhenAvailable -DontStopOnIdleEnd `
            -ExecutionTimeLimit (New-TimeSpan -Minutes 30)
        $principal = New-ScheduledTaskPrincipal -UserId "$env:USERDOMAIN\$env:USERNAME" -LogonType Interactive

        Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $dailyTrigger `
            -Settings $settings -Principal $principal `
            -Description "Auto-pulls Outlook emails into emails_recent.txt / emails_history.txt for the Claude email summary project." | Out-Null

        Write-Log "First run on this machine - registered Windows Scheduled Task '$taskName' (hourly, 6am-midnight, wakes the computer, catches up on missed runs)."
    }
} catch {
    Write-Log "WARNING: Could not auto-register the '$taskName' scheduled task ($($_.Exception.Message)). You'll need to set up the recurring run manually in Task Scheduler."
}

try {
    $o = New-Object -ComObject Outlook.Application
} catch {
    Write-Log "ERROR: Could not connect to Outlook. Make sure Outlook is installed and set up on this machine."
    exit 1
}

$ns    = $o.GetNamespace("MAPI")
$inbox = $ns.GetDefaultFolder(6)
$items = $inbox.Items

# --- Decide mode: first run (no state file yet) = full history. Otherwise = incremental. ---
$firstRun = -not (Test-Path $stateFile)

if ($firstRun) {
    Write-Log "No state file found - doing a full history pull (this can take a while for large mailboxes)."
    $items.Sort("[ReceivedTime]", $false)  # ascending, so the archive reads oldest -> newest
    $toProcess = $items
} else {
    $lastSyncIso = (Get-Content -Path $stateFile -Raw).Trim()
    try {
        $lastSync = [datetime]::Parse($lastSyncIso, $null, [System.Globalization.DateTimeStyles]::RoundtripKind)
    } catch {
        Write-Log "WARNING: Could not parse last_sync_state.txt ('$lastSyncIso'). Falling back to a full history pull."
        $lastSync = $null
    }
    if ($null -eq $lastSync) {
        $items.Sort("[ReceivedTime]", $false)
        $toProcess = $items
        $firstRun = $true   # treat like a first run for logging/write-mode purposes below
    } else {
        # Outlook's Restrict() filter wants its own date format (locale-dependent), separate from
        # the ISO string we keep in the state file.
        $filterDate = $lastSync.ToString("MM/dd/yyyy HH:mm:ss")
        $filter = "[ReceivedTime] > '$filterDate'"
        $toProcess = $items.Restrict($filter)
        $toProcess.Sort("[ReceivedTime]", $false)
        Write-Log "Incremental pull - fetching items received after $filterDate"
    }
}

$cutoff = (Get-Date).AddDays(-$recentWindowDays)

$sb       = New-Object System.Text.StringBuilder   # goes to history (everything)
$sbRecent = New-Object System.Text.StringBuilder   # goes to recent (only items within the window)
$found = 0
$maxSeen = $null

foreach ($it in $toProcess) {
    try {
        $rt = $it.ReceivedTime
        if ($null -eq $maxSeen -or $rt -gt $maxSeen) { $maxSeen = $rt }
        $time = $rt.ToString("yyyy-MM-dd HH:mm")
        $from = $it.SenderName
        $subj = $it.Subject
        $body = $it.Body
        if ($body -and $body.Length -gt 1500) { $body = $body.Substring(0, 1500) }
        $body = ($body -replace "\r?\n", " ").Trim()

        $entry = New-Object System.Text.StringBuilder
        [void]$entry.AppendLine("---")
        [void]$entry.AppendLine("TIME: $time")
        [void]$entry.AppendLine("FROM: $from")
        [void]$entry.AppendLine("SUBJ: $subj")
        [void]$entry.AppendLine("BODY: $body")
        [void]$entry.AppendLine("")
        $entryText = $entry.ToString()

        [void]$sb.Append($entryText)
        if ($rt -ge $cutoff) {
            [void]$sbRecent.Append($entryText)
        }
        $found++
    } catch {
        Write-Log "Error processing an item, skipped it."
    }
}

# --- History file: full history pull overwrites/creates the archive; incremental pulls append to it. Never trimmed. ---
if ($firstRun) {
    [System.IO.File]::WriteAllText($historyFile, $sb.ToString(), [System.Text.Encoding]::UTF8)
    Write-Log "Full history pull done. Wrote $found items to $historyFile"
} else {
    if ($found -gt 0) {
        [System.IO.File]::AppendAllText($historyFile, $sb.ToString(), [System.Text.Encoding]::UTF8)
    }
    Write-Log "Incremental pull done. Appended $found new item(s) to $historyFile"
}

# --- Recent file: first run writes it fresh (already filtered to the window); incremental runs append new items. ---
if ($firstRun) {
    [System.IO.File]::WriteAllText($recentFile, $sbRecent.ToString(), [System.Text.Encoding]::UTF8)
} else {
    if ($sbRecent.Length -gt 0) {
        [System.IO.File]::AppendAllText($recentFile, $sbRecent.ToString(), [System.Text.Encoding]::UTF8)
    }
}

# --- Trim the recent file down to the window every run, since old entries age out even with no new mail. ---
# The recent file stays small forever, so this re-parse is cheap regardless of how big the history file gets.
if (Test-Path $recentFile) {
    $recentContent = Get-Content -Path $recentFile -Raw
    if ($recentContent) {
        $blocks = $recentContent -split "(?m)^---\r?\n"
        $kept = New-Object System.Text.StringBuilder
        $keptCount = 0
        $droppedCount = 0
        foreach ($block in $blocks) {
            if (-not $block.Trim()) { continue }
            $keepBlock = $true
            if ($block -match "TIME:\s*(.+)") {
                $tStr = $matches[1].Trim()
                try {
                    $t = [datetime]::ParseExact($tStr, "yyyy-MM-dd HH:mm", $null)
                    if ($t -lt $cutoff) { $keepBlock = $false }
                } catch {
                    # Unparseable timestamp - keep it rather than silently lose an email.
                }
            }
            if ($keepBlock) {
                [void]$kept.Append("---`r`n")
                [void]$kept.Append($block)
                $keptCount++
            } else {
                $droppedCount++
            }
        }
        if ($droppedCount -gt 0) {
            [System.IO.File]::WriteAllText($recentFile, $kept.ToString(), [System.Text.Encoding]::UTF8)
            Write-Log "Trimmed recent file: kept $keptCount item(s), dropped $droppedCount item(s) older than $recentWindowDays days."
        }
    }
}

# --- Update the sync marker so next run only fetches what's newer than what we just saved. ---
if ($null -ne $maxSeen) {
    Set-Content -Path $stateFile -Value $maxSeen.ToString("o") -NoNewline
} elseif ($firstRun) {
    # Empty mailbox on a first run - still record a starting point so future runs behave correctly.
    Set-Content -Path $stateFile -Value (Get-Date).ToString("o") -NoNewline
}
