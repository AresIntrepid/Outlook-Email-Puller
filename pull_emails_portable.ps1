# Outlook Inbox export - full history on first run, incremental after that.
# Silent version - no console output, no prompts (except a first-run progress window).
# Works for any user - paths are relative to their own profile.
#
# Maintains THREE output files:
#   emails_history.txt  - permanent, cumulative, NEVER trimmed. Full archive, in case you
#                          ever want to look back further than the windows below.
#   emails_recent.txt   - rolling window, trimmed every run to the last $recentWindowDays
#                          days. The weekly summary automation reads this one.
#   emails_today.txt    - rolling window, trimmed every run to the last $todayWindowHours
#                          hours. The daily summary automation reads this one instead of
#                          emails_recent.txt, so it only ever has to read roughly a day's
#                          worth of data instead of the full 10-day window.
#
# Plus bookkeeping files:
#   last_sync_state.txt        - newest ReceivedTime we've saved (drives the incremental filter)
#   seen_ids.txt                - rolling list of recently-saved Outlook EntryIDs (and, for items
#                                  with no EntryID, content-hash keys). Main duplicate guard for
#                                  normal runs; see notes further down. Only holds the most recent
#                                  $seenIdCap ids, so a full mailbox rescan (see "Full mailbox scan
#                                  with an existing archive present" below) also cross-checks
#                                  against content already written to emails_history.txt, since
#                                  that rolling list alone isn't enough to cover one of those. This
#                                  cap is a rolling count across ALL ids, not just no-EntryID ones,
#                                  so in theory a content-hash key could still be evicted before the
#                                  same no-EntryID item reappears within the $overlapHours re-scan
#                                  window on an extremely high-volume mailbox; $seenIdCap is sized
#                                  generously to make that impractical in normal use.
#   unknown_time_failure_count.txt - counts CONSECUTIVE runs in which at least one item's
#                                  ReceivedTime could not even be read (see the "unknown time
#                                  failure" notes near the bottom). Reset to 0 the moment a run
#                                  completes without that problem.
#   pull_emails.lock            - simple PID+timestamp lock file preventing two copies of this
#                                  script (a scheduled run overlapping a manual run, or two
#                                  scheduled runs overlapping if one runs long) from writing to
#                                  the same files at the same time. See Acquire-Lock below.
#   Also note: emails_history.txt only ever stores the first 1500 characters of a body, and the
#   content-hash key used to detect duplicates during a full rescan is necessarily computed from
#   that same truncated text (there's no way to recover the untruncated original from what's on
#   disk). Two genuinely different emails would need to share the same to-the-minute time, sender
#   name, subject, AND first 1500 characters of body to collide there - narrow, but not provably
#   impossible for templated/automated senders. The same-session NOEID key (used for items Outlook
#   never gave an EntryID) doesn't have that constraint and hashes the full body instead.
#
# --- Reliability notes on this revision ---
# 1. Scheduled task now launches under whichever PowerShell host (Windows PowerShell OR
#    PowerShell 7 / pwsh) actually ran this registration, instead of always hardcoding
#    powershell.exe. Running under a different host than the one you tested with is a classic
#    source of "works interactively, breaks on schedule" bugs (different module set/defaults).
# 2. Long pulls (especially the very first full-history pull on a large mailbox) now checkpoint
#    their progress to disk every $checkpointInterval items instead of only at the very end. If
#    the process is killed mid-run (scheduled task's own time limit, reboot, logoff, Outlook
#    crash), everything up to the last checkpoint survives instead of the whole run being lost.
# 3. A single item whose ReceivedTime can't even be read used to unconditionally null out the
#    sync marker every run, forever, turning "incremental sync" into "full mailbox rescan on
#    every run" for as long as that one item existed. It now behaves like the already-existing
#    dated-failure give-up logic: the marker is held (forcing a retry) for up to
#    $maxUnknownTimeFailureRuns CONSECUTIVE runs, then the script gives up on that one item and
#    lets the marker advance based on everything else that succeeded.
# 4. A simple lock file now prevents two instances of this script from running concurrently and
#    corrupting each other's writes to seen_ids.txt / last_sync_state.txt / the window files.
$recentWindowDays          = 10    # covers the 7-day weekly summary plus a few days' buffer for a missed run
$todayWindowHours          = 24    # covers the daily summary's "since last run" window
$seenIdCap                 = 20000 # how many recent ids (EntryIDs + no-EntryID content hashes) to remember
$overlapHours              = 2     # how far back before the marker to re-scan; see the filter notes below
$maxItemErrorLogs          = 10    # per-item failures logged individually before switching to a summary
$checkpointInterval        = 250   # items processed between incremental disk flushes during a run
$maxUnknownTimeFailureRuns = 3     # consecutive runs an unreadable-timestamp item is retried before giving up on it
$lockStaleMinutes          = 60    # a lock file older than this AND whose owning PID is gone/dead is reclaimed
$dailyCacheRetentionDays   = 14    # daily_cache_YYYY-MM-DD.txt files older than this are deleted - see Trim-DailyCacheFiles notes below

$outDir = Join-Path $env:USERPROFILE "weekly report"
if (-not (Test-Path $outDir)) {
    New-Item -ItemType Directory -Path $outDir | Out-Null
}
$historyFile          = Join-Path $outDir "emails_history.txt"              # cumulative archive, appended to over time, never trimmed
$recentFile           = Join-Path $outDir "emails_recent.txt"               # rolling 10-day window - weekly automation reads this
$todayFile            = Join-Path $outDir "emails_today.txt"                # rolling 24-hour window - daily automation reads this
$stateFile            = Join-Path $outDir "last_sync_state.txt"             # remembers the newest ReceivedTime we've saved
$seenFile             = Join-Path $outDir "seen_ids.txt"                    # remembers recently-saved EntryIDs (duplicate guard)
$unknownFailCountFile = Join-Path $outDir "unknown_time_failure_count.txt"  # consecutive-run counter, see notes above
$lockFile             = Join-Path $outDir "pull_emails.lock"                # concurrency guard, see notes above
$logFile              = Join-Path $outDir "pull_emails_log.txt"

$utf8 = [System.Text.Encoding]::UTF8
$inv  = [System.Globalization.CultureInfo]::InvariantCulture
$sha  = [System.Security.Cryptography.SHA256]::Create()

function Get-ContentKey($t, $f, $s, $b) {
    $raw   = "$t`n$f`n$s`n$b"
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($raw)
    return [System.BitConverter]::ToString($sha.ComputeHash($bytes)).Replace("-", "")
}
function Get-BlockKey($block) {
    $t = $null; $f = ""; $s = ""; $b = ""
    if ($block -match "(?m)^TIME:\s*(.+)$") { $t = $matches[1].Trim() }
    if ($block -match "(?m)^FROM:\s*(.+)$") { $f = $matches[1].Trim() }
    if ($block -match "(?m)^SUBJ:\s*(.+)$") { $s = $matches[1].Trim() }
    if ($block -match "(?m)^BODY:\s*(.+)$") { $b = $matches[1].Trim() }
    if ($null -eq $t) { return $null }
    return (Get-ContentKey $t $f $s $b)
}
function Write-Log($msg) {
    Add-Content -Path $logFile -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - $msg" -Encoding UTF8
}
function Trim-LogFile($path, $maxLines) {
    if (-not (Test-Path $path)) { return }
    try {
        $lines = [System.IO.File]::ReadAllLines($path, $utf8)
        if ($lines.Count -gt $maxLines) {
            $keep = $lines[($lines.Count - $maxLines)..($lines.Count - 1)]
            [System.IO.File]::WriteAllLines($path, $keep, $utf8)
        }
    } catch {
    }
}
function Read-Utf8($path) {
    if (-not (Test-Path $path)) { return "" }
    try { return [System.IO.File]::ReadAllText($path, $utf8) } catch { return "" }
}
function Trim-WindowFile($filePath, $cutoff, $windowLabel) {
    if (-not (Test-Path $filePath)) { return }
    $content = Read-Utf8 $filePath
    if (-not $content) { return }
    $blocks = $content -split "(?m)^---\r?\n"
    $kept = New-Object System.Text.StringBuilder
    $keptCount = 0
    $droppedCount = 0
    foreach ($block in $blocks) {
        if (-not $block.Trim()) { continue }
        $keepBlock = $true
        if ($block -match "(?m)^TIME:\s*(.+)$") {
            $tStr = $matches[1].Trim()
            try {
                $t = [datetime]::ParseExact($tStr, "yyyy-MM-dd HH:mm", $inv)
                if ($t -lt $cutoff) { $keepBlock = $false }
            } catch {
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
        [System.IO.File]::WriteAllText($filePath, $kept.ToString(), $utf8)
        Write-Log "Trimmed $windowLabel file: kept $keptCount item(s), dropped $droppedCount item(s) outside the window."
    }
}
function Trim-DailyCacheFiles($dir, $retentionDays) {
    # The daily/weekly Claude Desktop automations write daily_cache_YYYY-MM-DD.txt files here as a
    # byproduct of the caching scheme (see the daily/weekly prompt instructions) so repeated runs
    # don't re-derive the same calendar day's categorization from raw email text. Nothing ever
    # reads a cache file older than the weekly lookback window, so left alone this would grow by
    # one small file per day forever with zero benefit past that point. This is intentionally done
    # HERE (a script that runs hourly on its own schedule) rather than as an instruction inside the
    # chat prompts, because file deletion by date is a purely mechanical decision that doesn't need
    # an LLM's judgment, and because it must not depend on the person remembering to trigger a
    # daily/weekly summary chat on any particular cadence - this way cleanup happens on schedule
    # regardless of whether those automations are ever opened.
    if (-not (Test-Path $dir)) { return }
    $cutoff = (Get-Date).Date.AddDays(-$retentionDays)
    $deleted = 0
    try {
        $candidates = Get-ChildItem -Path $dir -Filter "daily_cache_*.txt" -File -ErrorAction SilentlyContinue
    } catch {
        Write-Log "WARNING: could not list daily_cache_*.txt files in $dir for retention cleanup ($($_.Exception.Message))."
        return
    }
    foreach ($f in $candidates) {
        if ($f.Name -match '^daily_cache_(\d{4}-\d{2}-\d{2})\.txt$') {
            $fileDate = $null
            try { $fileDate = [datetime]::ParseExact($matches[1], "yyyy-MM-dd", $inv) } catch { $fileDate = $null }
            if ($null -ne $fileDate -and $fileDate -lt $cutoff) {
                try {
                    Remove-Item -Path $f.FullName -Force -ErrorAction Stop
                    $deleted++
                } catch {
                    Write-Log "WARNING: could not delete stale daily cache file '$($f.Name)' ($($_.Exception.Message))."
                }
            }
        }
        # Filenames that don't match the expected pattern are left alone rather than guessed at -
        # better to leak one unrecognized file than to delete something unrelated by mistake.
    }
    if ($deleted -gt 0) {
        Write-Log "Retention cleanup: deleted $deleted daily cache file(s) older than $retentionDays day(s)."
    }
}

# --- Concurrency guard -------------------------------------------------------------------------
# Prevents two instances of this script (a manual run overlapping the hourly scheduled run, or
# two scheduled runs overlapping if one runs unusually long) from writing to seen_ids.txt /
# last_sync_state.txt / the window files at the same time, which could otherwise corrupt or lose
# data (e.g. one process reading a window file mid-write by the other). A lock older than
# $lockStaleMinutes whose owning process is no longer running is treated as abandoned and reclaimed
# rather than permanently wedging the script.
#
# The actual claim is done with FileMode.CreateNew, which atomically fails if the file already
# exists - a plain "if not Test-Path then Set-Content" is NOT atomic (two instances can both pass
# the check before either writes, and both then believe they hold the lock); CreateNew is a single
# OS-level operation that only one caller can ever win.
function Try-ClaimLockFile($path, $content) {
    try {
        $fs = [System.IO.File]::Open($path, [System.IO.FileMode]::CreateNew, [System.IO.FileAccess]::Write)
        try {
            $bytes = [System.Text.Encoding]::UTF8.GetBytes($content)
            $fs.Write($bytes, 0, $bytes.Length)
        } finally { $fs.Close() }
        return $true
    } catch [System.IO.IOException] {
        return $false
    }
}
function Acquire-Lock($path, $staleMinutes) {
    $content = "$PID|$((Get-Date).ToUniversalTime().ToString('o'))"
    if (Try-ClaimLockFile $path $content) { return $true }
    # Someone already holds the file (or a stale leftover is sitting there) - decide whether to
    # reclaim it. This read-and-decide step doesn't need to be atomic: whatever we decide here,
    # the actual reclaim below still goes through the same atomic CreateNew, so a second instance
    # racing this same decision can still only win once.
    $isStale = $true
    try {
        $raw   = (Read-Utf8 $path).Trim()
        $parts = $raw -split '\|'
        if ($parts.Count -ge 2) {
            $lockPid  = 0
            $parsedPid = [int]::TryParse($parts[0], [ref]$lockPid)
            $lockTime = $null
            try { $lockTime = [datetime]::Parse($parts[1], $inv, [System.Globalization.DateTimeStyles]::RoundtripKind) } catch { $lockTime = $null }
            $ownerAlive = $false
            if ($parsedPid -and $lockPid -gt 0) {
                try { Get-Process -Id $lockPid -ErrorAction Stop | Out-Null; $ownerAlive = $true } catch { $ownerAlive = $false }
            }
            if ($ownerAlive -and $null -ne $lockTime -and ((Get-Date) - $lockTime) -lt (New-TimeSpan -Minutes $staleMinutes)) {
                $isStale = $false
            }
        }
    } catch {
        # Unreadable/corrupt lock file - treat as stale and reclaim it.
    }
    if (-not $isStale) { return $false }
    try { Remove-Item -Path $path -Force -ErrorAction SilentlyContinue } catch { }
    # Re-attempt through the same atomic claim rather than an unconditional overwrite - if another
    # instance reclaims first, this correctly loses instead of both instances proceeding.
    return (Try-ClaimLockFile $path $content)
}
function Release-Lock($path) {
    try { if (Test-Path $path) { Remove-Item -Path $path -Force -ErrorAction SilentlyContinue } } catch {
    }
}

if (-not (Acquire-Lock $lockFile $lockStaleMinutes)) {
    Write-Log "Another instance appears to already be running (lock file present and its owning process is alive within the last $lockStaleMinutes minute(s)). Exiting without doing any work."
    exit 0
}

try {

$taskName = "PullEmailsAutoSync"
try {
    $existingTask = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
    if (-not $existingTask) {
        $hostExePath = [System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
        if ($hostExePath -match '\\(powershell|pwsh)\.exe$') {
            # Launch the scheduled run with the SAME PowerShell host that registered the task
            # (Windows PowerShell or PowerShell 7), rather than hardcoding powershell.exe. If this
            # script is set up under pwsh but the task always launched powershell.exe, the
            # scheduled run would silently execute under a different engine/module set than the
            # one that was tested.
            $action = New-ScheduledTaskAction -Execute $hostExePath `
                -Argument "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$PSCommandPath`""
        } else {
            $action = New-ScheduledTaskAction -Execute $hostExePath
        }
        $repeatSeed   = New-ScheduledTaskTrigger -Once -At (Get-Date -Hour 6 -Minute 0 -Second 0) `
            -RepetitionInterval (New-TimeSpan -Hours 1) -RepetitionDuration (New-TimeSpan -Hours 18)
        $dailyTrigger = New-ScheduledTaskTrigger -Daily -At (Get-Date -Hour 6 -Minute 0 -Second 0)
        $dailyTrigger.Repetition = $repeatSeed.Repetition
        $settings  = New-ScheduledTaskSettingsSet -WakeToRun -StartWhenAvailable -DontStopOnIdleEnd `
            -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
            -ExecutionTimeLimit (New-TimeSpan -Minutes 30) -MultipleInstances IgnoreNew
        $principal = New-ScheduledTaskPrincipal -UserId "$env:USERDOMAIN\$env:USERNAME" -LogonType Interactive
        Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $dailyTrigger `
            -Settings $settings -Principal $principal `
            -Description "Auto-pulls Outlook emails into emails_today.txt / emails_recent.txt / emails_history.txt for the Claude email summary project." | Out-Null
        Write-Log "First run on this machine - registered Windows Scheduled Task '$taskName' (hourly, 6am-midnight, wakes the computer, runs on battery, catches up on missed runs, ignores overlapping starts)."
    } else {
        $s = $existingTask.Settings
        $needsUpdate = $false
        if ($s.DisallowStartIfOnBatteries -or $s.StopIfGoingOnBatteries) {
            $s.DisallowStartIfOnBatteries = $false
            $s.StopIfGoingOnBatteries     = $false
            $needsUpdate = $true
        }
        if ($s.MultipleInstances -ne 'IgnoreNew') {
            $s.MultipleInstances = 'IgnoreNew'
            $needsUpdate = $true
        }
        if ($needsUpdate) {
            Set-ScheduledTask -TaskName $taskName -Settings $s | Out-Null
            Write-Log "Repaired existing '$taskName' task: cleared the battery conditions and/or set MultipleInstances=IgnoreNew so overlapping runs no longer collide."
        }
        if ($existingTask.State -eq 'Disabled') {
            Write-Log "WARNING: scheduled task '$taskName' is DISABLED, so it will never run on its own. Re-enable it in Task Scheduler, or run: Enable-ScheduledTask -TaskName '$taskName'"
        }
        try {
            $regExe = ($existingTask.Actions | Select-Object -First 1).Execute
            if ($regExe) {
                $regExeResolved = [Environment]::ExpandEnvironmentVariables($regExe).Trim('"')
                if ([System.IO.Path]::IsPathRooted($regExeResolved) -and -not (Test-Path $regExeResolved)) {
                    Write-Log "WARNING: scheduled task '$taskName' points at '$regExeResolved', which no longer exists. Every scheduled run is failing. Delete the task and run this script again from its permanent location to re-register it."
                }
            }
        } catch {
        }
    }
} catch {
    Write-Log "WARNING: Could not auto-register or repair the '$taskName' scheduled task ($($_.Exception.Message)). You'll need to set up the recurring run manually in Task Scheduler."
}

try {
    $o = New-Object -ComObject Outlook.Application
} catch {
    Write-Log "ERROR: Could not connect to Outlook. Make sure Outlook is installed and set up on this machine."
    exit 1
}
try {
    $ns    = $o.GetNamespace("MAPI")
    $inbox = $ns.GetDefaultFolder(6)
    $items = $inbox.Items
} catch {
    Write-Log "ERROR: Connected to Outlook but could not open the Inbox ($($_.Exception.Message)). Outlook may still be starting up, or the mail profile may need attention."
    exit 1
}

$needsHistorySeed = (-not (Test-Path $recentFile)) -or (-not (Test-Path $todayFile))
$existingHistoryForSeed = ""
$historySeedFailed = $false
if ($needsHistorySeed) {
    $existingHistoryForSeed = Read-Utf8 $historyFile
    $historyLenForSeed = 0
    if (Test-Path $historyFile) { try { $historyLenForSeed = (Get-Item $historyFile).Length } catch { $historyLenForSeed = 0 } }
    if (-not $existingHistoryForSeed -and $historyLenForSeed -gt 0) {
        # Read-Utf8 swallows read errors and returns "" - without this check a transient lock
        # on emails_history.txt at exactly this moment would silently seed emails_recent.txt /
        # emails_today.txt as EMPTY even though the real archive has content. Skip seeding this
        # run instead; needsHistorySeed will still be true next run so it retries.
        $historySeedFailed = $true
        Write-Log "WARNING: $historyFile has content but could not be read to seed the rolling window file(s) (locked, or a transient share violation). Leaving them unseeded this run rather than writing them empty; will retry next run."
    }
}

$seenOrder = New-Object 'System.Collections.Generic.List[string]'
$seenIds   = New-Object 'System.Collections.Generic.HashSet[string]'
if (Test-Path $seenFile) {
    try {
        foreach ($line in [System.IO.File]::ReadAllLines($seenFile, $utf8)) {
            $id = $line.Trim()
            if ($id -and $seenIds.Add($id)) { [void]$seenOrder.Add($id) }
        }
    } catch {
        Write-Log "WARNING: Could not read seen_ids.txt - duplicate suppression is degraded for this run."
    }
}

$firstRun = -not (Test-Path $stateFile)
$lastSync = $null
if ($firstRun) {
    Write-Log "No state file found - doing a full history pull (this can take a while for large mailboxes)."
    $items.Sort("[ReceivedTime]", $false)
    $toProcess = $items
} else {
    $lastSyncIso = (Read-Utf8 $stateFile).Trim([char]0xFEFF, [char]0x200B).Trim()
    try {
        $lastSync = [datetime]::Parse($lastSyncIso, $inv, [System.Globalization.DateTimeStyles]::RoundtripKind)
    } catch {
        Write-Log "WARNING: Could not parse last_sync_state.txt ('$lastSyncIso'). Falling back to a full mailbox scan."
        $lastSync = $null
    }
    $lastSyncLocal = $null
    if ($null -ne $lastSync) {
        if ($lastSync.Kind -eq [System.DateTimeKind]::Utc) {
            $lastSyncLocal = $lastSync.ToLocalTime()
        } else {
            $lastSyncLocal = [datetime]::SpecifyKind($lastSync, [System.DateTimeKind]::Local)
        }
    }
    if ($null -eq $lastSync) {
        $items.Sort("[ReceivedTime]", $false)
        $toProcess = $items
        $firstRun = $true
    } else {
        $ci = [System.Globalization.CultureInfo]::CurrentCulture
        $fmt = $ci.DateTimeFormat.ShortDatePattern + " " + $ci.DateTimeFormat.ShortTimePattern
        $filterDate = $lastSyncLocal.AddHours(-$overlapHours).ToString($fmt, $ci)
        $filter = "[ReceivedTime] > '$filterDate'"
        $toProcess = $items.Restrict($filter)
        Write-Log "Incremental pull - fetching items received after $filterDate (marker $($lastSync.ToString('o')), $overlapHours-hour overlap, minute-precision filter)."
        $restrictCount = -1
        try { $restrictCount = $toProcess.Count } catch { $restrictCount = -1 }
        if ($restrictCount -le 0) {
            try {
                $items.Sort("[ReceivedTime]", $true)
                $newestInInbox = $null
                foreach ($probe in $items) { $newestInInbox = $probe.ReceivedTime; break }
                if ($null -ne $newestInInbox -and $newestInInbox -gt $lastSyncLocal) {
                    Write-Log "ERROR: filter '$filter' matched 0 items, but the Inbox holds mail newer than the marker (newest $($newestInInbox.ToString('yyyy-MM-dd HH:mm'))). The filter is not working on this machine. Falling back to a full Inbox scan for this run."
                    $toProcess = $items
                }
            } catch {
                Write-Log "WARNING: could not verify the empty filter result ($($_.Exception.Message))."
            }
        }
        $toProcess.Sort("[ReceivedTime]", $false)
    }
}

$historyHasContent = (Test-Path $historyFile) -and ((Get-Item $historyFile).Length -gt 0)
$rebuildHistory    = $firstRun -and (-not $historyHasContent)
$archiveKeys = New-Object 'System.Collections.Generic.HashSet[string]'
if ($firstRun -and $historyHasContent) {
    Write-Log "Full mailbox scan with an existing archive present - appending only unseen items rather than overwriting $historyFile ($($seenIds.Count) known id(s))."
    $existingHistoryForDedup = if ($needsHistorySeed) { $existingHistoryForSeed } else { Read-Utf8 $historyFile }
    foreach ($block in ($existingHistoryForDedup -split "(?m)^---\r?\n")) {
        if (-not $block.Trim()) { continue }
        $k = Get-BlockKey $block
        if ($k) { [void]$archiveKeys.Add($k) }
    }
    if ($archiveKeys.Count -eq 0) {
        Write-Log "ERROR: $historyFile has content but indexing it for the rescan produced 0 key(s) - it could not be read (locked, or a transient share violation). Aborting rather than risking duplicating the whole archive."
        exit 1
    }
    Write-Log "Indexed $($archiveKeys.Count) archived item(s) so the rescan cannot re-append them."
}

$cutoff      = (Get-Date).AddDays(-$recentWindowDays)
$todayCutoff = (Get-Date).AddHours(-$todayWindowHours)
$sb       = New-Object System.Text.StringBuilder
$sbRecent = New-Object System.Text.StringBuilder
$sbToday  = New-Object System.Text.StringBuilder
$found         = 0
$skipped       = 0
$maxSeen       = $null
$itemErrors    = 0
$oldestFailure = $null
$hadUnknownTimeFailure = $false
$processedCount = 0

# --- Checkpoint state ---------------------------------------------------------------------------
# The three file "initialized" flags track whether the one-time setup action for that file
# (an overwrite for a from-scratch rebuild, or a seed-from-history for a missing window file) has
# already happened THIS run. After that one-time action, every checkpoint (including the final one
# after the loop) simply appends whatever has accumulated in the corresponding StringBuilder since
# the previous checkpoint, then clears it. This lets a long pull flush partial progress to disk
# periodically - so a mid-run crash/timeout/logoff loses at most $checkpointInterval items' worth
# of work instead of the entire run - while producing byte-identical on-disk results to the
# original single-flush-at-the-end approach.
$script:historyInitialized = $false
$script:recentInitialized  = $false
$script:todayInitialized   = $false

function Flush-Checkpoint {
    # History file
    if (-not $script:historyInitialized) {
        if ($rebuildHistory) {
            [System.IO.File]::WriteAllText($historyFile, $sb.ToString(), $utf8)
        } elseif ($sb.Length -gt 0) {
            [System.IO.File]::AppendAllText($historyFile, $sb.ToString(), $utf8)
        }
        $script:historyInitialized = $true
    } elseif ($sb.Length -gt 0) {
        [System.IO.File]::AppendAllText($historyFile, $sb.ToString(), $utf8)
    }

    # Recent (10-day) file
    # NOTE: the seed-from-history step below re-reads emails_history.txt FRESH every time it's
    # attempted, rather than reusing the one-time $existingHistoryForSeed snapshot taken before
    # the loop started. This matters once checkpointing is in the picture: if the seed read fails
    # on an early checkpoint, we must NOT mark recentInitialized=true anyway, because every later
    # branch below only knows how to append (via AppendAllText, which silently CREATES the file
    # if it's missing) - if that branch ran first, the file would end up existing but permanently
    # missing its historical backfill, and since it now exists, $needsHistorySeed would be false
    # on every future run, so it could never be seeded at all. Deferring (not marking initialized)
    # keeps retrying at each later checkpoint and the final flush until the read actually succeeds
    # (or the history file is confirmed genuinely empty), and loses nothing in the meantime: any
    # items buffered for this file while deferred are already durably written to emails_history.txt
    # by the History-file section above, and get picked back up automatically once the seed finally
    # succeeds, because that fresh read naturally includes everything appended so far.
    if (-not $script:recentInitialized) {
        if ($rebuildHistory) {
            [System.IO.File]::WriteAllText($recentFile, $sbRecent.ToString(), $utf8)
            $script:recentInitialized = $true
        } elseif (-not (Test-Path $recentFile)) {
            $seedNow = Read-Utf8 $historyFile
            $histLenNow = 0
            if (Test-Path $historyFile) { try { $histLenNow = (Get-Item $historyFile).Length } catch { $histLenNow = 0 } }
            if ($seedNow -or $histLenNow -eq 0) {
                # Successful read (or the archive is legitimately empty, e.g. a brand-new mailbox) -
                # $seedNow already reflects everything written to $historyFile so far THIS run
                # (including this checkpoint's own new items, appended just above), so do NOT also
                # append $sbRecent here - that would duplicate them.
                [System.IO.File]::WriteAllText($recentFile, $seedNow, $utf8)
                $script:recentInitialized = $true
            } else {
                Write-Log "WARNING: still could not read $historyFile to seed $recentFile as of this checkpoint (locked, or a transient share violation) - deferring again."
            }
        } else {
            if ($sbRecent.Length -gt 0) { [System.IO.File]::AppendAllText($recentFile, $sbRecent.ToString(), $utf8) }
            $script:recentInitialized = $true
        }
    } elseif ($sbRecent.Length -gt 0) {
        [System.IO.File]::AppendAllText($recentFile, $sbRecent.ToString(), $utf8)
    }

    # Today (24-hour) file - mirrors the recent-file logic above, same reasoning
    if (-not $script:todayInitialized) {
        if ($rebuildHistory) {
            [System.IO.File]::WriteAllText($todayFile, $sbToday.ToString(), $utf8)
            $script:todayInitialized = $true
        } elseif (-not (Test-Path $todayFile)) {
            $seedNow = Read-Utf8 $historyFile
            $histLenNow = 0
            if (Test-Path $historyFile) { try { $histLenNow = (Get-Item $historyFile).Length } catch { $histLenNow = 0 } }
            if ($seedNow -or $histLenNow -eq 0) {
                [System.IO.File]::WriteAllText($todayFile, $seedNow, $utf8)
                $script:todayInitialized = $true
            } else {
                Write-Log "WARNING: still could not read $historyFile to seed $todayFile as of this checkpoint (locked, or a transient share violation) - deferring again."
            }
        } else {
            if ($sbToday.Length -gt 0) { [System.IO.File]::AppendAllText($todayFile, $sbToday.ToString(), $utf8) }
            $script:todayInitialized = $true
        }
    } elseif ($sbToday.Length -gt 0) {
        [System.IO.File]::AppendAllText($todayFile, $sbToday.ToString(), $utf8)
    }

    # Duplicate guard - persist what we've saved/seen so far
    try {
        $toWrite = $seenOrder
        if ($toWrite.Count -gt $seenIdCap) {
            $toWrite = $toWrite.GetRange($toWrite.Count - $seenIdCap, $seenIdCap)
        }
        [System.IO.File]::WriteAllLines($seenFile, $toWrite, $utf8)
    } catch {
        Write-Log "WARNING: checkpoint could not write seen_ids.txt ($($_.Exception.Message)). Duplicate suppression may be degraded if this run is interrupted."
    }

    # Sync marker - only safe to advance mid-run if nothing undated has failed yet THIS run;
    # the end-of-run give-up/hold logic further down makes the final, authoritative decision and
    # will overwrite this regardless, so an interim write here is a pure safety net, not the
    # source of truth.
    if (-not $hadUnknownTimeFailure -and $null -ne $maxSeen) {
        try {
            Set-Content -Path $stateFile -Value $maxSeen.ToUniversalTime().ToString("o") -Encoding UTF8 -NoNewline
        } catch {
            Write-Log "WARNING: checkpoint could not write last_sync_state.txt ($($_.Exception.Message))."
        }
    }

    [void]$sb.Clear()
    [void]$sbRecent.Clear()
    [void]$sbToday.Clear()
}

$totalToProcess = 0
if ($firstRun) {
    try { $totalToProcess = $toProcess.Count } catch { $totalToProcess = 0 }
}
$progressForm  = $null
$progressBar   = $null
$progressLabel = $null
if ($firstRun -and $totalToProcess -gt 0) {
    try {
        Add-Type -AssemblyName System.Windows.Forms
        Add-Type -AssemblyName System.Drawing
        $progressForm = New-Object System.Windows.Forms.Form
        $progressForm.Text            = "Setting up email sync"
        $progressForm.Width           = 420
        $progressForm.Height          = 140
        $progressForm.StartPosition   = "CenterScreen"
        $progressForm.FormBorderStyle = "FixedDialog"
        $progressForm.ControlBox      = $false
        $progressForm.MaximizeBox     = $false
        $progressForm.MinimizeBox     = $false
        $progressForm.TopMost         = $true
        $progressLabel = New-Object System.Windows.Forms.Label
        $progressLabel.Text     = "Pulling your full inbox history for the first time (0 / $totalToProcess)..."
        $progressLabel.AutoSize = $false
        $progressLabel.Width    = 380
        $progressLabel.Height   = 40
        $progressLabel.Top      = 15
        $progressLabel.Left     = 15
        $progressForm.Controls.Add($progressLabel)
        $progressBar = New-Object System.Windows.Forms.ProgressBar
        $progressBar.Minimum = 0
        $progressBar.Maximum = $totalToProcess
        $progressBar.Value   = 0
        $progressBar.Width   = 380
        $progressBar.Height  = 25
        $progressBar.Top     = 60
        $progressBar.Left    = 15
        $progressForm.Controls.Add($progressBar)
        $progressForm.Show()
        $progressForm.Refresh()
        [System.Windows.Forms.Application]::DoEvents()
    } catch {
        Write-Log "WARNING: Could not show the first-run progress window ($($_.Exception.Message)). Continuing without it."
        $progressForm = $null
    }
}

foreach ($it in $toProcess) {
    $rt = $null
    $processedCount++
    try {
        $eid = $null
        try { $eid = $it.EntryID } catch { $eid = $null }
        if ($eid -and $seenIds.Contains($eid)) {
            $skipped++
            try {
                $rtSeen = $it.ReceivedTime
                if ($null -eq $maxSeen -or $rtSeen -gt $maxSeen) { $maxSeen = $rtSeen }
            } catch { }
            continue
        }
        $rt = $it.ReceivedTime
        $time = $rt.ToString("yyyy-MM-dd HH:mm", $inv)
        $from = $it.SenderName
        $subj = $it.Subject
        $bodyFull = $it.Body
        $body = $bodyFull
        if ($body -and $body.Length -gt 1500) { $body = $body.Substring(0, 1500) }
        $from = ($from -replace "[\r\n]+", " ").Trim()
        $subj = ($subj -replace "[\r\n]+", " ").Trim()
        $body = ($body -replace "[\r\n]+", " ").Trim()
        # Full (untruncated) body, flattened the same way, used ONLY for the NOEID key below.
        # The archive-index comparison further down must keep using the truncated $body,
        # since that's all emails_history.txt ever stores for an entry - but the NOEID key
        # only ever compares against other hashes computed the same way in seen_ids.txt, so
        # it can safely use the fuller text to cut down on same-minute/same-subject items
        # colliding just because their first 1500 characters happen to match.
        $bodyForNoEidKey = if ($bodyFull) { ($bodyFull -replace "[\r\n]+", " ").Trim() } else { "" }
        if ($archiveKeys.Count -gt 0) {
            if ($archiveKeys.Contains((Get-ContentKey $time $from $subj $body))) {
                $skipped++
                if ($null -eq $maxSeen -or $rt -gt $maxSeen) { $maxSeen = $rt }
                continue
            }
        }
        if (-not $eid) {
            $eid = "NOEID|" + (Get-ContentKey $time $from $subj $bodyForNoEidKey)
            if ($seenIds.Contains($eid)) {
                $skipped++
                if ($null -eq $maxSeen -or $rt -gt $maxSeen) { $maxSeen = $rt }
                continue
            }
        }
        $entry = New-Object System.Text.StringBuilder
        [void]$entry.AppendLine("---")
        [void]$entry.AppendLine("TIME: $time")
        [void]$entry.AppendLine("FROM: $from")
        [void]$entry.AppendLine("SUBJ: $subj")
        [void]$entry.AppendLine("BODY: $body")
        [void]$entry.AppendLine("")
        $entryText = $entry.ToString()
        [void]$sb.Append($entryText)
        if ($rt -ge $cutoff)      { [void]$sbRecent.Append($entryText) }
        if ($rt -ge $todayCutoff) { [void]$sbToday.Append($entryText) }
        if ($seenIds.Add($eid)) { [void]$seenOrder.Add($eid) }
        if ($null -eq $maxSeen -or $rt -gt $maxSeen) { $maxSeen = $rt }
        $found++
    } catch {
        $itemErrors++
        if ($null -ne $rt -and ($null -eq $oldestFailure -or $rt -lt $oldestFailure)) {
            $oldestFailure = $rt
        } elseif ($null -eq $rt) {
            $hadUnknownTimeFailure = $true
        }
        if ($itemErrors -le $maxItemErrorLogs) {
            $whenStr = "unknown time"
            if ($null -ne $rt) { $whenStr = $rt.ToString("yyyy-MM-dd HH:mm", $inv) }
            Write-Log "WARNING: could not save the item received $whenStr ($($_.Exception.Message))."
        }
    }

    if ($progressForm -and ($processedCount % 10 -eq 0 -or $processedCount -eq $totalToProcess)) {
        $progressBar.Value  = [Math]::Min($processedCount, $totalToProcess)
        $progressLabel.Text = "Pulling your full inbox history for the first time ($processedCount / $totalToProcess checked, $found saved)..."
        $progressForm.Refresh()
        [System.Windows.Forms.Application]::DoEvents()
    }
    if ($processedCount % $checkpointInterval -eq 0) {
        Flush-Checkpoint
    }
}

if ($itemErrors -gt $maxItemErrorLogs) {
    Write-Log "WARNING: $itemErrors item(s) failed this run; only the first $maxItemErrorLogs were logged individually."
}
if ($progressForm) {
    $progressLabel.Text = "Done - saved $found email(s)."
    $progressBar.Value = $totalToProcess
    if ($totalToProcess -gt 0) { $progressBar.Value = $totalToProcess - 1 }
    $progressBar.Value = $totalToProcess
    $holdUntil = (Get-Date).AddMilliseconds(1500)
    while ((Get-Date) -lt $holdUntil) {
        [System.Windows.Forms.Application]::DoEvents()
        Start-Sleep -Milliseconds 50
    }
    $progressForm.Close()
    $progressForm.Dispose()
}

# Final flush - persists whatever accumulated since the last periodic checkpoint (or everything,
# if the run never reached $checkpointInterval items). Whether each file's contents came from one
# flush or several, the on-disk result is identical to the original single-flush-at-the-end design.
Flush-Checkpoint

if ($rebuildHistory) {
    Write-Log "Full history pull done. Wrote $found item(s) to $historyFile"
} else {
    Write-Log "Incremental pull done. Appended $found new item(s) to $historyFile ($skipped already-seen item(s) skipped, $itemErrors error(s))."
}

Trim-WindowFile -filePath $recentFile -cutoff $cutoff -windowLabel "recent (10-day)"
Trim-WindowFile -filePath $todayFile  -cutoff $todayCutoff -windowLabel "today (24-hour)"
Trim-DailyCacheFiles -dir $outDir -retentionDays $dailyCacheRetentionDays

# --- Whether it's safe to give up on the oldest failure and advance past it depends on whether
#     anything else this run actually worked. If some mail saved, or was recognized as already
#     saved, Outlook and the mailbox are fine and the failure belongs to one specific email -
#     exactly the case this give-up rule exists for. If NOTHING at all could be read this run,
#     the cause is systemic (Outlook mid-sync, a corrupt profile, an OST rebuild), not one bad
#     email - and without this check, a sustained outage makes the marker recede by roughly the
#     overlap window on every single run (each run's oldest failure sits at the edge of a window
#     defined by the previous run's already-receded marker, a feedback loop), drifting further
#     into the past for as long as the outage lasts, unbounded. Verified by simulating 20 days of
#     a continuous outage: without this check the marker drifted back 40 days; with it, the
#     marker correctly freezes in place instead of running away. ---
if ($null -ne $oldestFailure -and ($null -eq $maxSeen -or $maxSeen -ge $oldestFailure)) {
    $stallFloor = (Get-Date).AddDays(-$recentWindowDays)
    $anythingWorked = ($found -gt 0 -or $skipped -gt 0)
    if ($oldestFailure -lt $stallFloor -and $anythingWorked) {
        Write-Log "ERROR: an item received $($oldestFailure.ToString('yyyy-MM-dd HH:mm', $inv)) has failed to save for over $recentWindowDays days while other mail in the same run saved fine. Giving up on it and advancing the marker; that email will not appear in the archive."
        if ($null -eq $maxSeen -or $maxSeen -lt $oldestFailure) { $maxSeen = $oldestFailure }
    } elseif ($oldestFailure -lt $stallFloor) {
        Write-Log "ERROR: nothing in this run could be read at all, and the oldest failure is now over $recentWindowDays days old. NOT giving up on it: when every single item fails, the cause is Outlook or the mailbox rather than one bad email, and advancing the marker would write off good mail a run at a time. The marker stays put and this will repeat until the underlying problem is fixed."
    } else {
        $maxSeen = $oldestFailure.AddSeconds(-1)
        Write-Log "WARNING: holding the sync marker at $($maxSeen.ToString('yyyy-MM-dd HH:mm:ss', $inv)) so the $itemErrors failed item(s) are retried next run."
    }
}

# --- Unreadable-timestamp items used to unconditionally null $maxSeen on every occurrence, which
#     (since last_sync_state.txt then never gets written) permanently turns every future run into
#     a full mailbox rescan for as long as that one item exists. This mirrors the dated-failure
#     give-up pattern above: hold the marker (forcing a safe retry, since already-saved items are
#     re-skipped via the duplicate guards) for up to $maxUnknownTimeFailureRuns CONSECUTIVE runs,
#     then give up on that one item specifically and let the marker advance based on everything
#     that WAS read successfully this run. The counter resets to 0 the moment a run completes
#     clean, so a one-off transient COM hiccup behaves exactly as before (marker held, retried next
#     run) and only a truly stuck/broken item trips the give-up path.
if ($hadUnknownTimeFailure) {
    $prevUnknownCount = 0
    if (Test-Path $unknownFailCountFile) {
        try {
            $parsedOk = 0
            [void][int]::TryParse((Read-Utf8 $unknownFailCountFile).Trim(), [ref]$parsedOk)
            $prevUnknownCount = $parsedOk
        } catch {
            $prevUnknownCount = 0
        }
    }
    $thisUnknownCount = $prevUnknownCount + 1
    if ($thisUnknownCount -ge $maxUnknownTimeFailureRuns) {
        Write-Log "ERROR: an item with an unreadable received time has now failed $thisUnknownCount consecutive run(s) (limit $maxUnknownTimeFailureRuns). Giving up on waiting for it and advancing the marker based on everything that WAS read successfully this run; full window coverage can no longer be strictly guaranteed for whatever this item is, but it will not block sync going forward."
        try { Set-Content -Path $unknownFailCountFile -Value "0" -Encoding UTF8 -NoNewline } catch { }
        # Deliberately leave $maxSeen as already computed from successfully-read items - do NOT null it.
    } else {
        Write-Log "WARNING: at least one failed item's received time could not be read (occurrence $thisUnknownCount of $maxUnknownTimeFailureRuns before this is given up on). Leaving the sync marker unchanged; the whole window is retried next run (already-saved items are skipped again via the duplicate guards)."
        try { Set-Content -Path $unknownFailCountFile -Value "$thisUnknownCount" -Encoding UTF8 -NoNewline } catch { }
        $maxSeen = $null
    }
} else {
    if (Test-Path $unknownFailCountFile) {
        try { Remove-Item -Path $unknownFailCountFile -Force -ErrorAction SilentlyContinue } catch { }
    }
}

if ($null -ne $maxSeen) {
    Set-Content -Path $stateFile -Value $maxSeen.ToUniversalTime().ToString("o") -Encoding UTF8 -NoNewline
} elseif ($firstRun -and $totalToProcess -eq 0) {
    Set-Content -Path $stateFile -Value (Get-Date).ToUniversalTime().ToString("o") -Encoding UTF8 -NoNewline
}

Trim-LogFile -path $logFile -maxLines 2000

} finally {
    Release-Lock $lockFile
}
