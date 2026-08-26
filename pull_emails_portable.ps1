# Outlook Inbox export - full history on first run, incremental after that.
# Silent version - no console output, no prompts.
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
# Plus two bookkeeping files:
#   last_sync_state.txt - newest ReceivedTime we've saved (drives the incremental filter)
#   seen_ids.txt        - rolling list of recently-saved Outlook EntryIDs. This is what
#                          actually guarantees no duplicates; see the notes further down.

$recentWindowDays = 10   # covers the 7-day weekly summary plus a few days' buffer for a missed run
$todayWindowHours = 24   # covers the daily summary's "since last run" window
$seenIdCap        = 5000 # how many recent EntryIDs to remember for duplicate suppression
$overlapHours     = 2    # how far back before the marker to re-scan; see the filter notes below
$maxItemErrorLogs = 10   # per-item failures logged individually before switching to a summary

$outDir = Join-Path $env:USERPROFILE "weekly report"
if (-not (Test-Path $outDir)) {
    New-Item -ItemType Directory -Path $outDir | Out-Null
}

$historyFile = Join-Path $outDir "emails_history.txt"   # cumulative archive, appended to over time, never trimmed
$recentFile  = Join-Path $outDir "emails_recent.txt"    # rolling 10-day window - weekly automation reads this
$todayFile   = Join-Path $outDir "emails_today.txt"     # rolling 24-hour window - daily automation reads this
$stateFile   = Join-Path $outDir "last_sync_state.txt"  # remembers the newest ReceivedTime we've saved
$seenFile    = Join-Path $outDir "seen_ids.txt"         # remembers recently-saved EntryIDs (duplicate guard)
$logFile     = Join-Path $outDir "pull_emails_log.txt"

$utf8 = [System.Text.Encoding]::UTF8
$inv  = [System.Globalization.CultureInfo]::InvariantCulture

function Write-Log($msg) {
    # -Encoding UTF8 matters: without it Windows PowerShell writes the active ANSI codepage while
    # Trim-LogFile below reads and rewrites the file as UTF-8, so any non-ASCII character was
    # permanently mangled on the next trim. .NET exception messages are localised, so on a
    # non-English Windows that corrupted exactly the diagnostic lines worth reading.
    Add-Content -Path $logFile -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - $msg" -Encoding UTF8
}

# --- Keep the log from growing without bound. Runs hourly forever, so an untrimmed log is a
#     slow leak. Cheap to check, only rewrites when it's actually over the cap. ---
function Trim-LogFile($path, $maxLines) {
    if (-not (Test-Path $path)) { return }
    try {
        $lines = [System.IO.File]::ReadAllLines($path, $utf8)
        if ($lines.Count -gt $maxLines) {
            $keep = $lines[($lines.Count - $maxLines)..($lines.Count - 1)]
            [System.IO.File]::WriteAllLines($path, $keep, $utf8)
        }
    } catch {
        # Never let log maintenance break the actual pull.
    }
}

# --- Read a text file as UTF-8 explicitly. The originals were written with UTF8 encoding, so
#     read them back the same way rather than relying on Get-Content's default codepage
#     detection - that only works as long as the byte-order mark survives, and a file that ever
#     loses its BOM would start silently mangling non-ASCII sender names and subjects. ---
function Read-Utf8($path) {
    if (-not (Test-Path $path)) { return "" }
    try { return [System.IO.File]::ReadAllText($path, $utf8) } catch { return "" }
}

# --- Shared trim logic for both rolling-window files: parses each "---" delimited entry, checks
#     its TIME field against the given cutoff, and rewrites the file keeping only what's still
#     within the window. Written once and called once per window file below, rather than copied,
#     so a future fix to this logic only needs to happen in one place instead of drifting
#     between two near-identical copies. ---
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
        # Anchored to the start of a line. Unanchored, a "TIME:" appearing inside a subject or
        # body would be picked up as the record's timestamp and the block would be kept or
        # dropped on the wrong date.
        if ($block -match "(?m)^TIME:\s*(.+)$") {
            $tStr = $matches[1].Trim()
            try {
                # InvariantCulture, not $null. $null means CurrentCulture, and on a machine whose
                # locale uses different date separators this ParseExact throws for every single
                # block - which lands in the catch below and keeps everything, silently turning
                # the "rolling" windows into files that never roll.
                $t = [datetime]::ParseExact($tStr, "yyyy-MM-dd HH:mm", $inv)
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
        [System.IO.File]::WriteAllText($filePath, $kept.ToString(), $utf8)
        Write-Log "Trimmed $windowLabel file: kept $keptCount item(s), dropped $droppedCount item(s) outside the window."
    }
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
            # NOTE: the task stores this absolute path. If the exe lives somewhere transient like
            # Downloads, moving or cleaning it up silently breaks the task. Prefer keeping it in
            # a stable location such as %LOCALAPPDATA%\PullEmails\ before the first run.
            $action = New-ScheduledTaskAction -Execute $hostExePath
        }

        # "Repeat every 1 hour for 18 hours, once a day starting at 6am" - built via a throwaway
        # one-time trigger so we can lift its Repetition settings onto the real daily trigger.
        $repeatSeed   = New-ScheduledTaskTrigger -Once -At (Get-Date -Hour 6 -Minute 0 -Second 0) `
            -RepetitionInterval (New-TimeSpan -Hours 1) -RepetitionDuration (New-TimeSpan -Hours 18)
        $dailyTrigger = New-ScheduledTaskTrigger -Daily -At (Get-Date -Hour 6 -Minute 0 -Second 0)
        $dailyTrigger.Repetition = $repeatSeed.Repetition

        # -AllowStartIfOnBatteries / -DontStopIfGoingOnBatteries are REQUIRED here. Without them
        # New-ScheduledTaskSettingsSet defaults both battery conditions to "on", which means the
        # task refuses to start whenever the laptop is unplugged, and kills a run already in
        # progress the moment you pull the power.
        $settings  = New-ScheduledTaskSettingsSet -WakeToRun -StartWhenAvailable -DontStopOnIdleEnd `
            -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
            -ExecutionTimeLimit (New-TimeSpan -Minutes 30)

        # LogonType Interactive is deliberate and must stay. Outlook COM automation
        # (New-Object -ComObject Outlook.Application) needs a real interactive desktop session -
        # switching this to S4U or "run whether user is logged on or not" makes the pull fail.
        $principal = New-ScheduledTaskPrincipal -UserId "$env:USERDOMAIN\$env:USERNAME" -LogonType Interactive

        Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $dailyTrigger `
            -Settings $settings -Principal $principal `
            -Description "Auto-pulls Outlook emails into emails_today.txt / emails_recent.txt / emails_history.txt for the Claude email summary project." | Out-Null

        Write-Log "First run on this machine - registered Windows Scheduled Task '$taskName' (hourly, 6am-midnight, wakes the computer, runs on battery, catches up on missed runs)."
    } else {
        # --- Repair pass for machines that were set up by an older version of this script, where
        #     the task got registered with the default battery conditions. The block above is
        #     skipped once the task exists, so without this those machines would keep skipping
        #     every run made while unplugged. Only writes when something actually needs changing. ---
        $s = $existingTask.Settings
        if ($s.DisallowStartIfOnBatteries -or $s.StopIfGoingOnBatteries) {
            $s.DisallowStartIfOnBatteries = $false
            $s.StopIfGoingOnBatteries     = $false
            Set-ScheduledTask -TaskName $taskName -Settings $s | Out-Null
            Write-Log "Repaired existing '$taskName' task: cleared the battery conditions so it runs while unplugged."
        }

        # --- The checks below do NOT auto-repair, they only report. Get-ScheduledTask returning
        #     an object was previously treated as "the install is healthy", which hid the two
        #     most common ways this silently stops running forever. Deliberately not fixed
        #     automatically: re-enabling a task someone disabled on purpose, or rewriting the
        #     action path, are both decisions the user should make rather than a background
        #     job making them silently. ---
        if ($existingTask.State -eq 'Disabled') {
            Write-Log "WARNING: scheduled task '$taskName' is DISABLED, so it will never run on its own. Re-enable it in Task Scheduler, or run: Enable-ScheduledTask -TaskName '$taskName'"
        }

        try {
            $regExe = ($existingTask.Actions | Select-Object -First 1).Execute
            if ($regExe) {
                $regExeResolved = [Environment]::ExpandEnvironmentVariables($regExe).Trim('"')
                # A bare command name (e.g. "powershell.exe") resolves via PATH, so only flag
                # paths that are rooted and genuinely missing.
                if ([System.IO.Path]::IsPathRooted($regExeResolved) -and -not (Test-Path $regExeResolved)) {
                    Write-Log "WARNING: scheduled task '$taskName' points at '$regExeResolved', which no longer exists. Every scheduled run is failing. Delete the task and run this script again from its permanent location to re-register it."
                }
            }
        } catch {
            # Reporting only - never let this break the pull.
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

# GetDefaultFolder throws routinely in the real world - Outlook still building its profile, a
# corrupt OST, "cannot open your default e-mail folders". Unguarded, the script died here with an
# unhandled error and wrote NOTHING to the log, so the only trace was a scheduled task quietly
# recording a non-zero exit code that nobody looks at.
try {
    $ns    = $o.GetNamespace("MAPI")
    $inbox = $ns.GetDefaultFolder(6)
    $items = $inbox.Items
} catch {
    Write-Log "ERROR: Connected to Outlook but could not open the Inbox ($($_.Exception.Message)). Outlook may still be starting up, or the mail profile may need attention."
    exit 1
}

# --- Snapshot whatever's currently in the history archive, BEFORE this run adds anything new to
#     it. Used below to seed any rolling-window file that doesn't exist yet - either because this
#     is truly the first-ever run, or because it's a window file introduced in a script update on
#     a machine that's already been running an older version. Taking the snapshot now (before
#     history gets updated) means seeding a new window file can never double-count this run's own
#     new items with what gets appended to it further down.
#
#     Guarded, because this value is consumed at most once ever per window file but the read used
#     to happen on every single hourly run. emails_history.txt is the one file here that grows
#     without bound, so an unconditional read meant pulling the entire archive off disk and into
#     a string 24 times a day, forever, and discarding it. At ~600 emails/week that reaches
#     hundreds of MB per read within a few years. Read it only when something actually needs it. ---
$needsHistorySeed = (-not (Test-Path $recentFile)) -or (-not (Test-Path $todayFile))
$existingHistoryForSeed = ""
if ($needsHistorySeed) { $existingHistoryForSeed = Read-Utf8 $historyFile }

# --- Duplicate guard. Several separate things can make the same email come back a second time:
#
#       1. The scan window deliberately overlaps the marker by $overlapHours, so a normal run
#          re-fetches everything already saved in that window. That overlap is what protects
#          against clock changes and filter imprecision, and this guard is what makes it free.
#       2. A run that dies partway - logoff, shutdown, the 30-minute execution limit - has
#          already appended to emails_history.txt but never reached the state-file write at the
#          bottom. The next run starts from the old marker and refetches those same items.
#       3. A full mailbox scan on a machine that already has an archive, which happens whenever
#          last_sync_state.txt cannot be read. Every item comes back; only unseen ones are kept.
#
#     Reordering the writes doesn't fix this: writing the state marker first would trade
#     duplicates for permanently LOST email, which is far worse. So instead we remember the
#     EntryIDs we've already written and skip them on sight. That covers both causes, and it
#     stays correct no matter how a run is interrupted. ---
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

# --- Decide mode: first run (no state file yet) = full history. Otherwise = incremental. ---
$firstRun = -not (Test-Path $stateFile)
$lastSync = $null   # hoisted so the duplicate guard and logging below can both see it

if ($firstRun) {
    Write-Log "No state file found - doing a full history pull (this can take a while for large mailboxes)."
    $items.Sort("[ReceivedTime]", $false)  # ascending, so the archive reads oldest -> newest
    $toProcess = $items
} else {
    # Trim() alone is not enough: a UTF-8 byte-order mark surfaces as U+FEFF, which .NET does not
    # classify as whitespace, and a single stray character here is enough to fail the parse and
    # send the whole run down the full-scan path.
    $lastSyncIso = (Read-Utf8 $stateFile).Trim([char]0xFEFF, [char]0x200B).Trim()
    try {
        $lastSync = [datetime]::Parse($lastSyncIso, $inv, [System.Globalization.DateTimeStyles]::RoundtripKind)
    } catch {
        Write-Log "WARNING: Could not parse last_sync_state.txt ('$lastSyncIso'). Falling back to a full mailbox scan."
        $lastSync = $null
    }

    # The marker is now written in UTC so that changing timezone (travel) or a DST shift cannot
    # move it relative to incoming mail. Outlook's ReceivedTime is local wall-clock, so convert
    # back to local before comparing against it.
    #
    # Markers written by older versions have no offset and parse as Unspecified. DateTime's
    # ToLocalTime() treats Unspecified as UTC, which would shift those by the whole timezone
    # offset, so tag them Local explicitly instead of converting them.
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
        $firstRun = $true   # treat like a first run for logging/write-mode purposes below
    } else {
        # Outlook's Restrict() wants the date formatted in the machine's own locale, NOT a
        # hardcoded US pattern - "MM/dd/yyyy HH:mm:ss" throws on a machine set to, say, en-GB or
        # de-DE. Build the string from the current culture instead.
        #
        # CRITICAL: ShortTimePattern (h:mm), never LongTimePattern (h:mm:ss). Outlook's Jet filter
        # on [ReceivedTime] only supports MINUTE precision, and it does not reject a filter string
        # containing seconds - it silently matches NOTHING. Verified directly against a live
        # mailbox holding 5 qualifying items:
        #     "[ReceivedTime] > '8/26/2026 7:57:57 AM'"  -> 0    (seconds, silently empty)
        #     "[ReceivedTime] > '08/26/2026 07:57:57'"   -> 0    (seconds, silently empty)
        #     "[ReceivedTime] > '8/26/2026 7:57 AM'"     -> 5    (correct)
        # A filter with seconds makes every incremental pull return zero items forever, which also
        # freezes the sync marker (it only advances when items are found), so the failure is
        # completely silent: the log just keeps saying "Appended 0 new item(s)".
        #
        # We also deliberately back the filter off rather than using the exact marker. Restrict is
        # only a coarse prefilter here; asking for too much is free, because the EntryID guard
        # discards anything already saved. Being too NARROW silently drops mail forever.
        #
        # The overlap is hours, not minutes, specifically to absorb the clock moving backwards:
        # an NTP correction on a drifted laptop, a DST transition caught by a catch-up run, or a
        # short timezone change. Any of those can put the marker ahead of incoming mail, and
        # every minute of that gap is mail no future run would ever fetch. Re-scanning a couple
        # of hours of already-saved mail costs one property read per item and nothing else.
        $ci = [System.Globalization.CultureInfo]::CurrentCulture
        $fmt = $ci.DateTimeFormat.ShortDatePattern + " " + $ci.DateTimeFormat.ShortTimePattern
        $filterDate = $lastSyncLocal.AddHours(-$overlapHours).ToString($fmt, $ci)
        $filter = "[ReceivedTime] > '$filterDate'"
        $toProcess = $items.Restrict($filter)
        Write-Log "Incremental pull - fetching items received after $filterDate (marker $($lastSync.ToString('o')), $overlapHours-hour overlap, minute-precision filter)."

        # --- Safety net for the entire class of bug that has bitten this script twice now: a
        #     filter that Outlook accepts and then silently matches nothing. Zero results are
        #     indistinguishable from a genuinely quiet inbox, which is exactly why the seconds
        #     bug survived undetected for so long.
        #
        #     So when the filter returns nothing, ask the Inbox directly whether it holds anything
        #     newer than the marker. That is one property read on one item. If it does, the filter
        #     is lying and we fall back to scanning the whole Inbox for this run - slower, but the
        #     EntryID guard makes it correct, and it self-heals instead of quietly losing mail.
        #     This also covers the locale hazards the culture-formatted filter still carries
        #     (non-Gregorian calendars, bidi marks, ICU's narrow no-break space before AM/PM). ---
        $restrictCount = -1
        try { $restrictCount = $toProcess.Count } catch { $restrictCount = -1 }

        if ($restrictCount -le 0) {
            try {
                $items.Sort("[ReceivedTime]", $true)   # descending, so the first item is newest
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

# --- Whether this run is allowed to OVERWRITE the permanent archive, as opposed to appending to
#     it. This is deliberately much narrower than $firstRun.
#
#     $firstRun means "scan the whole mailbox", and it gets set not only on a genuine first run
#     but any time last_sync_state.txt cannot be read - a transient file lock from OneDrive, an
#     antivirus scan, a 0-byte file left by a run killed mid-write. Wiring an unrecoverable
#     WriteAllText on emails_history.txt to that condition meant a momentary hiccup on a
#     bookkeeping file could replace years of archive with "whatever is in the Inbox right now",
#     silently, under a log line reading "Full history pull done."
#
#     So: overwrite ONLY when there is demonstrably nothing to lose. In every other case the
#     archive is append-only, and the EntryID guard keeps a full rescan from duplicating what is
#     already in there. That also removes the old need to discard seen_ids.txt on a full run,
#     which is why that block is gone - the guard now stays armed, which is what stops a forced
#     re-pull from appending a second copy of the last 10 days into the rolling window files. ---
$historyHasContent = (Test-Path $historyFile) -and ((Get-Item $historyFile).Length -gt 0)
$rebuildHistory    = $firstRun -and (-not $historyHasContent)

if ($firstRun -and $historyHasContent) {
    Write-Log "Full mailbox scan with an existing archive present - appending only unseen items rather than overwriting $historyFile ($($seenIds.Count) known id(s))."
}

$cutoff      = (Get-Date).AddDays(-$recentWindowDays)
$todayCutoff = (Get-Date).AddHours(-$todayWindowHours)

$sb       = New-Object System.Text.StringBuilder   # goes to history (everything)
$sbRecent = New-Object System.Text.StringBuilder   # goes to recent (only items within the 10-day window)
$sbToday  = New-Object System.Text.StringBuilder   # goes to today (only items within the 24-hour window)
$found         = 0
$skipped       = 0
$maxSeen       = $null
$itemErrors    = 0
$oldestFailure = $null   # ReceivedTime of the earliest item this run failed to save; see below

# --- Progress window: only for the first-ever run (the full history pull), since that's the only
#     run slow enough to need one. Every run after this is a fast incremental pull (a handful of
#     items, done in a second or two) and stays completely silent, including on the hourly
#     scheduled runs - this window is a one-time thing the user sees exactly once, ever. ---
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
        # If Windows Forms isn't available for some reason, just skip the UI silently -
        # the pull itself must never fail because a progress bar couldn't be shown.
        Write-Log "WARNING: Could not show the first-run progress window ($($_.Exception.Message)). Continuing without it."
        $progressForm = $null
    }
}

foreach ($it in $toProcess) {
    $rt = $null   # hoisted so the catch below can tell how far back a failure reaches
    try {
        # Duplicate guard, checked before anything is written.
        $eid = $null
        try { $eid = $it.EntryID } catch { $eid = $null }
        if ($eid -and $seenIds.Contains($eid)) {
            $skipped++
            # This item is already saved, so the marker may safely move past it. Without this the
            # marker sat still whenever everything in range had already been seen, which left the
            # scan window permanently wider than it needed to be.
            try {
                $rtSeen = $it.ReceivedTime
                if ($null -eq $maxSeen -or $rtSeen -gt $maxSeen) { $maxSeen = $rtSeen }
            } catch { }
            continue
        }

        $rt = $it.ReceivedTime

        # InvariantCulture here too - Trim-WindowFile above parses this exact string back out with
        # $inv, so writing it with the machine's CurrentCulture (e.g. a locale where ":" isn't the
        # time separator) would make that parse fail on every single block, which is precisely the
        # "silently stops trimming" failure this format was chosen to prevent.
        $time = $rt.ToString("yyyy-MM-dd HH:mm", $inv)
        $from = $it.SenderName
        $subj = $it.Subject
        $body = $it.Body
        if ($body -and $body.Length -gt 1500) { $body = $body.Substring(0, 1500) }

        # Every field is flattened to a single line, not just the body. The record format is
        # line-oriented and "---" on its own line separates records, so a newline inside a sender
        # display name or subject splits one record into two. The orphan half has no TIME: line,
        # which means the rolling-window trim keeps it forever, and every downstream reader sees a
        # phantom email. Display names on unresolved external mail come from the From: header and
        # are not under our control.
        $from = ($from -replace "[\r\n]+", " ").Trim()
        $subj = ($subj -replace "[\r\n]+", " ").Trim()
        $body = ($body -replace "[\r\n]+", " ").Trim()

        # Some items expose no usable EntryID. Previously those fell through with $eid still null,
        # so nothing was ever recorded in seen_ids and the SAME email was re-appended on every
        # single run until newer mail pushed the window past it - dozens of copies over a quiet
        # weekend, each logged as a perfectly ordinary "Appended 1 new item(s)". Fall back to a
        # content key so the duplicate guard still has something to hold onto.
        if (-not $eid) {
            $eid = "NOEID|$time|$from|$subj"
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

        # --- The marker advances LAST, and only for an item that is now genuinely captured.
        #     It used to advance immediately after reading ReceivedTime, before SenderName,
        #     Subject and Body - any of which throw in normal operation on rights-protected mail,
        #     delivery reports, or an Outlook that is busy syncing. When one did, the item was
        #     skipped but the marker had already moved past it, so the next run's window started
        #     after an email that was never saved and no future run would ever fetch it again.
        #     Silent, permanent loss out of a file documented as a complete archive. ---
        if ($null -eq $maxSeen -or $rt -gt $maxSeen) { $maxSeen = $rt }
        $found++

        # Update every 10 items rather than every single one - keeps the UI responsive without
        # slowing the pull down with constant repaints on a large mailbox.
        if ($progressForm -and ($found % 10 -eq 0 -or $found -eq $totalToProcess)) {
            $progressBar.Value    = [Math]::Min($found, $totalToProcess)
            $progressLabel.Text   = "Pulling your full inbox history for the first time ($found / $totalToProcess)..."
            $progressForm.Refresh()
            [System.Windows.Forms.Application]::DoEvents()
        }
    } catch {
        $itemErrors++

        # Remember how far back the damage reaches. The marker is clamped to sit before this
        # below, so a failed item is retried on the next run instead of being stepped over.
        if ($null -ne $rt -and ($null -eq $oldestFailure -or $rt -lt $oldestFailure)) {
            $oldestFailure = $rt
        }

        # The old message was "Error processing an item, skipped it." with no identifying detail
        # at all, so there was no way to tell one unreadable delivery receipt from forty real
        # emails. Log enough to identify it - but cap the individual lines, because a mass failure
        # would otherwise write thousands of entries and Trim-LogFile would then scroll the
        # evidence out of the log within the same run.
        if ($itemErrors -le $maxItemErrorLogs) {
            $whenStr = "unknown time"
            if ($null -ne $rt) { $whenStr = $rt.ToString("yyyy-MM-dd HH:mm", $inv) }
            Write-Log "WARNING: could not save the item received $whenStr ($($_.Exception.Message))."
        }
    }
}

if ($itemErrors -gt $maxItemErrorLogs) {
    Write-Log "WARNING: $itemErrors item(s) failed this run; only the first $maxItemErrorLogs were logged individually."
}

if ($progressForm) {
    # Hold the finished state on screen for a moment so it's actually readable - on a small
    # mailbox the whole pull can finish in well under a second, which would otherwise make the
    # window flash and vanish before anyone can see it. This adds nothing noticeable on a large
    # mailbox, where the pull itself already takes minutes.
    $progressLabel.Text = "Done - saved $found email(s)."

    # The native Windows progress bar animates a value change over a short fill transition rather
    # than snapping instantly. Nudge it down then back up to Maximum to force an immediate
    # redraw instead of an animated one, so it visibly reads 100% right away.
    $progressBar.Value = $totalToProcess
    if ($totalToProcess -gt 0) { $progressBar.Value = $totalToProcess - 1 }
    $progressBar.Value = $totalToProcess

    # Keep pumping Windows messages for the whole hold - a plain Start-Sleep here would freeze
    # the message loop entirely, leaving the bar's fill animation visibly stuck mid-transition
    # for the whole hold instead of showing it complete.
    $holdUntil = (Get-Date).AddMilliseconds(1500)
    while ((Get-Date) -lt $holdUntil) {
        [System.Windows.Forms.Application]::DoEvents()
        Start-Sleep -Milliseconds 50
    }

    $progressForm.Close()
    $progressForm.Dispose()
}

# --- History file. Never trimmed, and only ever overwritten when there was demonstrably nothing
#     in it to lose (see $rebuildHistory above). Every other path appends. ---
if ($rebuildHistory) {
    [System.IO.File]::WriteAllText($historyFile, $sb.ToString(), $utf8)
    Write-Log "Full history pull done. Wrote $found item(s) to $historyFile"
} else {
    if ($found -gt 0) {
        [System.IO.File]::AppendAllText($historyFile, $sb.ToString(), $utf8)
    }
    Write-Log "Incremental pull done. Appended $found new item(s) to $historyFile ($skipped already-seen item(s) skipped, $itemErrors error(s))."
}

# --- Recent + today files: if a window file doesn't exist yet - whether this is truly the
#     first-ever run, or a window file introduced in an update running on a machine that already
#     has history - seed it from the pre-run history snapshot captured above first, then append
#     this run's own new items on top (same either way, first run or not). The trim step further
#     below then cuts whatever's in the file down to its actual window regardless of how it got
#     there, so a history-seeded file ends up correctly sized just like a freshly-built one. ---
if ($rebuildHistory) {
    # Building everything from nothing, so write the window files outright rather than seeding
    # and appending. $sbRecent and $sbToday already hold exactly the in-window items from this
    # full scan, and the archive they would have been seeded from is this same run's output.
    [System.IO.File]::WriteAllText($recentFile, $sbRecent.ToString(), $utf8)
    [System.IO.File]::WriteAllText($todayFile,  $sbToday.ToString(),  $utf8)
} else {
    if (-not (Test-Path $recentFile)) {
        [System.IO.File]::WriteAllText($recentFile, $existingHistoryForSeed, $utf8)
    }
    if ($sbRecent.Length -gt 0) {
        [System.IO.File]::AppendAllText($recentFile, $sbRecent.ToString(), $utf8)
    }

    if (-not (Test-Path $todayFile)) {
        [System.IO.File]::WriteAllText($todayFile, $existingHistoryForSeed, $utf8)
    }
    if ($sbToday.Length -gt 0) {
        [System.IO.File]::AppendAllText($todayFile, $sbToday.ToString(), $utf8)
    }
}

# --- Trim both rolling-window files down to their windows every run, since old entries age out
#     even on a run with no new mail at all. Both files stay small forever, so this re-parse is
#     cheap regardless of how big the history file gets. Same shared function, two windows. ---
Trim-WindowFile -filePath $recentFile -cutoff $cutoff -windowLabel "recent (10-day)"
Trim-WindowFile -filePath $todayFile  -cutoff $todayCutoff -windowLabel "today (24-hour)"

# --- Persist the duplicate guard BEFORE the state marker. Order matters on a crash: if this
#     write lands and the marker write doesn't, the next run refetches the same items and the
#     EntryID list quietly discards them - no duplicates, no loss. Capped to the most recent
#     $seenIdCap ids so it can't grow without bound; that's far more overlap than the 5-minute
#     Restrict window can ever produce. ---
try {
    $toWrite = $seenOrder
    if ($toWrite.Count -gt $seenIdCap) {
        $toWrite = $toWrite.GetRange($toWrite.Count - $seenIdCap, $seenIdCap)
    }
    [System.IO.File]::WriteAllLines($seenFile, $toWrite, $utf8)
} catch {
    Write-Log "WARNING: Could not write seen_ids.txt ($($_.Exception.Message)). Duplicate suppression may be degraded next run."
}

# --- Hold the marker behind anything this run failed to save, so the next run gets another go at
#     it instead of stepping over it permanently. A genuinely unreadable item therefore stalls the
#     marker and keeps logging a warning every hour, which is noisy on purpose: the alternative is
#     losing the email in silence. The stall is bounded so the rescan window can never outgrow the
#     $seenIdCap ids that keep the rescan from producing duplicates - past that point the item is
#     given up on, loudly, rather than degrading the whole tool. ---
if ($null -ne $oldestFailure -and $null -ne $maxSeen -and $maxSeen -ge $oldestFailure) {
    $stallFloor = (Get-Date).AddDays(-$recentWindowDays)
    if ($oldestFailure -lt $stallFloor) {
        Write-Log "ERROR: an item received $($oldestFailure.ToString('yyyy-MM-dd HH:mm', $inv)) has failed to save for over $recentWindowDays days. Giving up on it and advancing the marker; that email will not appear in the archive."
    } else {
        $maxSeen = $oldestFailure.AddSeconds(-1)
        Write-Log "WARNING: holding the sync marker at $($maxSeen.ToString('yyyy-MM-dd HH:mm:ss', $inv)) so the $itemErrors failed item(s) are retried next run."
    }
}

# --- Update the sync marker so next run only fetches what's newer than what we just saved.
#     Stored in UTC: the marker is compared against Outlook's local ReceivedTime, so keeping it as
#     local wall-clock meant a timezone change or DST shift moved the marker relative to incoming
#     mail and silently skipped everything in the gap. ---
if ($null -ne $maxSeen) {
    Set-Content -Path $stateFile -Value $maxSeen.ToUniversalTime().ToString("o") -Encoding UTF8 -NoNewline
} elseif ($firstRun) {
    # Empty mailbox on a first run - still record a starting point so future runs behave correctly.
    Set-Content -Path $stateFile -Value (Get-Date).ToUniversalTime().ToString("o") -Encoding UTF8 -NoNewline
}

Trim-LogFile -path $logFile -maxLines 2000
