# Outlook Email Puller

A small Windows tool that exports your Outlook inbox to plain text files, useful for feeding into AI tools (like Claude with a connected folder) that can't read your inbox directly.

## What it does

* Connects to your local Outlook desktop app (via COM automation)
* On first run, pulls your **entire inbox history** and writes it to `emails_history.txt` — a permanent archive that's never trimmed
* On every run after that, pulls only what's arrived since the last run and appends it to `emails_history.txt`
* Also maintains two rolling-window files, kept in sync with the archive but automatically trimmed every run so they stay fast to read no matter how large `emails_history.txt` grows:
  * `emails_recent.txt` — last 10 days
  * `emails_today.txt` — last 24 hours (this is the one a daily summary should read, so it only ever has to process roughly a day's worth of data instead of a full 10-day window)
* **Sets itself up to run automatically.** The very first time it runs on a machine, it registers a Windows Scheduled Task (`PullEmailsAutoSync`) that repeats hourly from 6am–midnight, wakes the computer if it's asleep, catches up if a run gets missed, and refuses to run a second copy of itself if a previous run is still in progress
* **Shows progress on the one run that's slow.** The very first run (the full history pull) can take a few minutes on a large mailbox, so it shows a small progress window for that run only. Every run after that — including all future hourly background runs — is fast and completely silent, no window at all
* **Checkpoints progress during long pulls.** A very large first-time history pull writes its progress to disk periodically instead of only at the end, so a scheduled task hitting its time limit, a reboot, or a logoff mid-pull loses at most a few hundred emails' worth of progress rather than the whole run

Output format per email (same across all three files):

```
---
TIME: 2026-08-05 09:48
FROM: Sender Name
SUBJ: Subject line
BODY: First 1500 characters of the email body, flattened to one line

```

## Requirements

* Windows
* Microsoft Outlook desktop app installed and set up with your mail profile (this does not work with Outlook web/OWA — it needs the desktop COM interface)

## Usage

### Option A: Run the compiled app (easiest)

1. Download `PullEmails.exe` from the [Releases](https://github.com/AresIntrepid/Outlook-Email-Puller/releases) page
2. Double-click it
3. Approve any Windows SmartScreen prompt ("More info" → "Run anyway") — this is expected for an unsigned exe
4. A small progress window shows the first pull happening — this can take a few minutes for a large inbox. It closes itself automatically when done; if you want more detail on what it did, check `pull_emails_log.txt` (see below)
5. Output lands in `C:\Users\<you>\weekly report\` — you're done. It's now running itself automatically in the background from here on, silently, with no further windows

### Option B: Run the PowerShell script directly

```
powershell -ExecutionPolicy Bypass -File "pull_emails_portable.ps1"
```

Same behavior as the exe, including the self-install step and the first-run progress window.

## Automating it

You don't need to do anything — this is now automatic. On first run (either the exe or the script), it registers a Windows Scheduled Task called `PullEmailsAutoSync` for you, configured to:

* Run hourly, 6am–midnight, every day
* Wake the computer if it's asleep when a run is due
* Catch up on a missed run as soon as the machine is available again
* Ignore a new trigger if a previous run of the task is still going, rather than starting a second one alongside it
* Run under your own logged-in account (required — Outlook's COM automation doesn't work reliably running as SYSTEM or "whether logged on or not")
* Run on battery as well as plugged in
* **Only run while you are logged in.** Outlook's COM automation requires an interactive desktop session, so a locked screen is fine but a signed-out or shut-down machine is not. Missed runs are caught up automatically at your next login, and nothing is lost, since the pull resumes from a stored marker.

If for some reason the self-install fails (check `pull_emails_log.txt` for a warning), you can set it up manually instead: Task Scheduler → Create Basic Task → Trigger: Daily, repeating hourly → Action: Start a program → `powershell` with arguments `-ExecutionPolicy Bypass -File "C:\path\to\pull_emails_portable.ps1"`.

## Where the folder comes from

The script auto-creates a `weekly report` folder inside `%USERPROFILE%` (`C:\Users\<you>\weekly report`) if it doesn't already exist — no manual setup needed. It contains:

* `emails_history.txt` — the permanent, complete archive
* `emails_recent.txt` — the rolling 10-day window
* `emails_today.txt` — the rolling 24-hour window
* `last_sync_state.txt` — internal bookkeeping (the UTC timestamp of the newest email saved so far)
* `seen_ids.txt` — internal bookkeeping (recently-saved Outlook item IDs, used to suppress duplicates)
* `unknown_time_failure_count.txt` — internal bookkeeping (only appears if an item's received time couldn't be read; tracks how many consecutive runs that's happened so the script knows when to give up on it rather than stall forever — see v2.5 below)
* `pull_emails.lock` — internal bookkeeping (exists only while a run is actively in progress; prevents two copies of the script from running at once)
* `pull_emails_log.txt` — a running log of what each run did
* `daily_cache_YYYY-MM-DD.txt` (one per past day) — **not created by this script.** These are written by the Claude Desktop/Cowork daily and weekly automations described below, as part of their own caching scheme, so re-running a summary doesn't re-derive the same day's categorization from scratch every time. This script only cleans up old ones (see the caching section below) — it never reads or writes their content itself.

## Using it with Claude / Cowork

Connect the `weekly report` folder as a project folder in Claude Desktop (Cowork mode). A sensible default split:

* A **daily** summary automation should read `emails_today.txt` — it's already limited to the last 24 hours, so no further date filtering is needed
* A **weekly** summary automation should read `emails_recent.txt` — filter within it to whatever week window you want (e.g. Saturday through Friday)
* Anything else — a question naming a specific person or company, a request for full history, or anything the recent files don't answer — should fall back to `emails_history.txt`, searched by filtering for the name/domain rather than reading the whole file into context (see the project instructions below)

**Important:** connected folders aren't watched live — Claude can end up working from a cached view of a file instead of what's actually on disk right now. Build your instructions to explicitly say something like *"always re-read this file fresh — never rely on a cached or previously-read version"* rather than assuming it'll notice the file changed on its own.

### Named-person/company lookups stay fast as the archive grows

`emails_history.txt` only ever grows, so a naive "read the whole file" approach to a question like *"full picture on Delta"* gets slower forever as your archive ages. The project instructions below tell Claude to search/filter `emails_history.txt` for the name or domain first, and only load the matching entries into context — instead of reading the entire file. This keeps person/company lookups fast regardless of how many years of mail you've accumulated.

### Daily/weekly summary caching

Re-running a weekly summary the day after you last ran it means re-deriving categorization for 6 days you already summarized, just to get 1 new day's worth of information. The daily and weekly prompts below solve this with a small caching scheme, entirely within the chat automations (no extra script or API involved):

* Once a calendar day has **fully elapsed**, whichever automation runs next (daily or weekly — either one can do it) writes a `daily_cache_YYYY-MM-DD.txt` file summarizing that day's categorized emails, if one doesn't already exist.
* Future runs read that cached file instead of re-deriving the day's categorization from raw email text.
* **Today is never cached** — a cache written for a day that isn't over yet would be permanently wrong once more mail arrives, so the live "today" summary is always computed fresh from `emails_today.txt`/`emails_recent.txt`, never from a cache file.
* Because either automation can fill in a missing day, this works correctly regardless of the order or frequency you actually run things in — you don't need daily to have run every day in sequence before weekly benefits from it.
* **Cleanup is handled by the script, not the chat prompts.** Since nothing ever reads a cache file older than the weekly lookback window, this script deletes `daily_cache_*.txt` files older than 14 days on every run (see v2.6 below) — a purely mechanical, date-based deletion that doesn't need an LLM's judgment, and that happens on the script's own hourly schedule regardless of whether you ever open a summary chat.

### Project instructions

Applies to every chat in the project, including ad-hoc one-off questions.

```
Emails are saved in three files in this folder: emails_today.txt (last 24 hours, the fastest and smallest), emails_recent.txt (last 10 days), and emails_history.txt (the complete, unbounded archive going back to when tracking started).

Use emails_today.txt for anything about today or "since I last checked." Use emails_recent.txt for anything else recent — this week, the last several days — and as the default when I don't specify a period.

Use emails_history.txt whenever I ask about a specific person or company by name — including phrasings like "full picture on X," "everything from X," "history with X," or "how many times has X emailed me" — or whenever I explicitly ask you to look back further than 10 days, or the answer isn't found in the other two files. For these named-entity queries specifically: do NOT read the entire file into context. Instead, search/filter emails_history.txt for the name, email address, or domain I gave you (case-insensitive, and try reasonable variants — e.g. "Delta" should also match "delta.com" or "Delta Air Lines") and only load the matching entries into context for your answer. This keeps the response fast regardless of how large the archive has grown. Only fall back to a full unfiltered read of the file if a targeted search genuinely doesn't find anything plausible and you need to double-check nothing was missed due to a naming mismatch.

For any other request that requires emails_history.txt but isn't about a specific named sender (e.g. "how many emails total have I gotten"), search or filter as appropriate rather than always loading the whole file, but use your judgment — some questions genuinely require scanning everything.

All three files are updated outside this conversation by a scheduled script — always re-read them fresh each time; never rely on a cached or previously-read version. If necessary, update all three files using the exe/script provided in this folder. Never ask about connecting an email account.
```

### Daily automation

Runs once a day. Reads `emails_today.txt`, which is already pre-filtered to the last 24 hours by the script, so no date filtering happens in the prompt itself. Also silently backfills yesterday's cache file if one doesn't exist yet, so the weekly automation can reuse it later.

```
Summarize the emails in the given period. Read emails_today.txt in this folder directly — never ask about connecting an email account. This file is updated outside this conversation by a scheduled script, so always re-read it fresh; never rely on a cached or previously-read version. It's a rolling window of the last 24 hours, appended in chronological order (oldest first, newest last) — include every entry in the file, since it's already limited to the right window.

Title it "Inbox summary ([date range])". Group emails into these five sections, in this order, and include every section even if empty (say "None" rather than omitting it): Needs action — awaiting a reply or decision from me. Resolved — was actionable but already handled (note how/when). Overdue — deadline has passed with no action taken. FYI — informational, no action needed. Noise — automated alerts, verification codes, spam. Mention only as a count, don't list individually. 1–2 sentence bullets per item. Output the summary directly as plain text in your reply — do not create or attach a document, artifact, or file for it.

--- Cache backfill (do this silently, after the summary above; don't mention it in your reply) ---
Check whether a file named daily_cache_<YESTERDAY'S DATE, format YYYY-MM-DD>.txt already exists in this folder (yesterday = the calendar day immediately before today, fully elapsed). If it already exists, do nothing further. If it does NOT exist, generate one now: read emails_recent.txt in this folder, filter to only entries whose TIME falls on yesterday's calendar date, classify them into the same five categories used above, and write a file named daily_cache_<YYYY-MM-DD>.txt with exactly this format:

DATE: <YYYY-MM-DD>
NEEDS_ACTION:
- <1-2 sentence bullet>
- <1-2 sentence bullet>
RESOLVED:
- <bullet>
OVERDUE:
- <bullet>
FYI:
- <bullet>
NOISE_COUNT: <integer>

Use "None" as the only line under a section if it's empty (still include the header). Never generate or overwrite a cache file for TODAY's date — today isn't over yet, so a cache for it would be incomplete and wrong if reused later. Only ever cache a date that has fully elapsed.
```

### Weekly automation

Runs once a week (Friday). Reads `emails_recent.txt` (the 10-day window) and filters within the prompt to the Saturday–Friday range, since the file itself isn't pre-filtered that tightly. Reuses cached days where they already exist, and fills in whichever are still missing.

```
Summarize the emails from the past week. This covers last Saturday through today (Friday).

For each date in that range:
- If the date is TODAY (not yet fully elapsed): always compute it fresh. Read emails_recent.txt in this folder directly — always re-read it fresh, never rely on a cached or previously-read version — filter to entries whose TIME falls on today's date, and classify them into the five categories below. Never read or write a cache file for today.
- If the date is in the past (fully elapsed): check whether daily_cache_<YYYY-MM-DD>.txt already exists in this folder for that date.
  - If it exists, read it and use its NEEDS_ACTION / RESOLVED / OVERDUE / FYI bullets and NOISE_COUNT directly — do not recompute or re-derive them from raw email text.
  - If it does NOT exist, read emails_recent.txt, filter to entries whose TIME falls on that date, classify them into the five categories, AND write a daily_cache_<YYYY-MM-DD>.txt file in the exact format below so future runs (daily or weekly) can reuse it. Never write a cache file for a date that hasn't fully elapsed yet.

Cache file format (must match exactly, for compatibility with the daily automation):
DATE: <YYYY-MM-DD>
NEEDS_ACTION:
- <bullet>
RESOLVED:
- <bullet>
OVERDUE:
- <bullet>
FYI:
- <bullet>
NOISE_COUNT: <integer>
Use "None" as the only line under an empty section (still include the header).

After gathering all seven days (a mix of cached reads and freshly computed days, however many of each), combine them into one summary. Title it "Weekly inbox summary ([Sat–Fri date range])". Group into the same five sections, in this order, each included even if empty (say "None" rather than omitting it): Needs action, Resolved, Overdue, FYI (1-2 sentence bullets per item, carried over as-is from cached days), and Noise (sum the NOISE_COUNT values across all seven days and report only the total count, never list individually).

If emails_recent.txt doesn't cover a date that needs fresh computation (older than its 10-day window), fall back to emails_history.txt filtered to that date instead.

Output the summary directly as plain text in your reply — do not create or attach a document, artifact, or file for it.
```

## Privacy note

This tool only reads your own inbox and writes to a local file on your own machine — it doesn't send anything over the network. Each person who runs it only gets access to their own mail; there's no way to pull someone else's inbox with this. (The daily/weekly caching scheme above runs entirely inside your Claude Desktop/Cowork conversation and local files too — no separate API or account is involved.)

## Building the exe yourself

```
Install-Module -Name ps2exe -Scope CurrentUser -Force
Invoke-ps2exe -inputFile "pull_emails_portable.ps1" -outputFile "PullEmails.exe" -noConsole -STA -title "Pull Outlook Emails" -icon "mail_icon.ico"
```

Notes on the flags: `-noConsole` (not `-noConsole:$false`) because this runs silently once an hour in the background via the scheduled task, and a console window popping up unannounced every hour would be unwelcome. `-STA` is needed because the first-run progress window uses Windows Forms, which requires a single-threaded apartment to behave reliably alongside Outlook's own COM automation.

## Checking it's working

A pull that finds nothing looks identical to a pull that is broken, so it is worth knowing what a healthy log line looks like. The quick check:

```
Get-Content "$env:USERPROFILE\weekly report\pull_emails_log.txt" -Tail 3
```

A healthy run reads something like:

```
Incremental pull done. Appended 0 new item(s) to ... (5 already-seen item(s) skipped, 0 error(s)).
```

Zero appended with a **non-zero skipped count** is the good signal. The scan window deliberately overlaps the last sync by two hours, so a working run always re-examines a handful of already-saved emails and discards them. Zero appended **and** zero skipped, repeatedly, while mail is arriving, means the filter is returning nothing and something is wrong.

The other field worth watching is the sync marker:

```
Get-Content "$env:USERPROFILE\weekly report\last_sync_state.txt"
```

It should track close to your newest email. A marker frozen at an old value while mail keeps arriving is the clearest sign of a stalled pull.

Fuller check, including the scheduled task itself:

```
$dir = "$env:USERPROFILE\weekly report"
Get-ScheduledTask -TaskName PullEmailsAutoSync | Select-Object TaskName, State
Get-ScheduledTaskInfo -TaskName PullEmailsAutoSync | Select-Object LastRunTime, LastTaskResult, NextRunTime, NumberOfMissedRuns
Get-Content "$dir\last_sync_state.txt"
Get-Content "$dir\pull_emails_log.txt" -Tail 6
```

`State` should be `Ready`, and `LastTaskResult` should be `0`. A `LastTaskResult` of `267011` just means the task has not run yet.

If a run seems to have not done anything and you see a `pull_emails.lock` file sitting in the folder, that normally just means a run is currently in progress — it deletes itself when the run finishes. It should never persist for more than about an hour; if it does, check `pull_emails_log.txt`, since a lock older than 60 minutes with no owning process still running is automatically reclaimed on the next run rather than blocking things forever.

## Version history

**v2.6**
* Added automatic cleanup of old `daily_cache_YYYY-MM-DD.txt` files (see the caching section above) — these are written by the Claude Desktop daily/weekly automations, not this script, but since nothing ever reads one older than the weekly lookback window, the script now deletes any older than 14 days on every run so the folder doesn't accumulate one small file per day forever with no purpose. Deliberately implemented here rather than as a chat-prompt instruction, since it's a purely mechanical date-based decision that shouldn't depend on an LLM's judgment or on how often you happen to open a summary chat.
* Updated the project/daily/weekly prompt instructions: named-person/company lookups against `emails_history.txt` now search/filter for the name or domain instead of reading the entire file into context, so they stay fast regardless of archive size; and the daily/weekly automations now maintain the `daily_cache_*.txt` scheme described above so repeated weekly runs don't re-derive the same day's categorization from scratch.

**v2.5**
* **Fixed a genuine race condition in the scheduled-task lock.** The lock was previously claimed with a "check if the file exists, then create it" sequence, which is not atomic — two instances starting close together could both see no lock present and both proceed, defeating the lock entirely. Reproduced against a live PID with two concurrent claim attempts; the second one incorrectly won. Fixed by claiming the lock with `FileMode.CreateNew`, an atomic OS-level operation that only one caller can ever succeed at; the stale-lock reclaim path was re-verified to still work correctly through the same atomic claim.
* **Fixed mid-run checkpointing being able to permanently disable seeding of `emails_recent.txt`/`emails_today.txt`.** If seeding one of those files from the archive failed on an early checkpoint (e.g. a transient file lock), the file was still marked as "handled" for the rest of the run, so every later checkpoint silently created it via append with only partial content — and because it now existed, no future run would ever retry seeding it either. The seed step now re-reads the archive fresh at every checkpoint attempt and only marks the file as done once that read genuinely succeeds (or the archive is confirmed to be legitimately empty), so a transient failure is retried instead of permanently skipped, with no data loss and no duplicated content in either outcome.

**v2.4**
* Scheduled task now launches under whichever PowerShell host (Windows PowerShell or PowerShell 7 / `pwsh`) actually registered it, instead of always hardcoding `powershell.exe` — running under a different host than the one you tested with is a classic source of "works interactively, breaks on schedule" bugs.
* Long pulls (especially the very first full-history pull on a large mailbox) now checkpoint their progress to disk periodically instead of only at the very end, so a scheduled task hitting its time limit, a reboot, or a logoff mid-pull loses at most a few hundred emails' worth of progress instead of the whole run.
* An item whose received time can't even be read used to unconditionally reset the sync marker every single run, forever, turning "incremental sync" into "full mailbox rescan on every run" for as long as that one item existed. It now behaves like the existing dated-failure give-up logic: the marker is held (forcing a retry) for a few consecutive runs, then the script gives up on that one item specifically and lets the marker advance based on everything else that succeeded.
* Added a lock file preventing two copies of the script from running concurrently and corrupting each other's writes to the bookkeeping files (later hardened to be fully atomic in v2.5, see above).
* Scheduled task is now registered with `MultipleInstances = IgnoreNew`, so Task Scheduler itself won't fire an overlapping run if a previous one is still going.

**v2.3**
* **Fixed silent, permanent mail loss when an item fails to read.** The sync marker advanced as soon as an email's timestamp was read, before its sender, subject and body. Any of those throw in normal operation, on rights-protected mail, delivery reports, or an Outlook busy syncing. When one did, the email was skipped but the marker had already moved past it, so no future run would ever fetch it again. The marker now advances only after an email is genuinely written, and is held behind any failure so it is retried on the next run rather than stepped over. Failures now log the item's timestamp and the actual error instead of an anonymous "Error processing an item".
* **Fixed the permanent archive being destroyed by a transient file error.** A `last_sync_state.txt` that could not be read triggered a full mailbox scan, and a full scan overwrote `emails_history.txt` outright, replacing the entire archive with a snapshot of the current Inbox. A momentary file lock from cloud sync or an antivirus scan was enough to trigger it, and the log line read "Full history pull done." The archive is now append-only unless it is provably empty.
* Fixed a forced re-pull appending a second copy of the last 10 days into `emails_recent.txt` and `emails_today.txt`, which would have made the daily and weekly summaries double-count.
* Fixed emails whose Outlook EntryID could not be read being re-saved on every run indefinitely, since they were never recorded in the duplicate guard. They now fall back to a content-based key.
* **Added a safety net for silently-empty filters.** When the incremental filter returns nothing, the script now checks whether the Inbox actually holds mail newer than the marker. If it does, it logs an ERROR and scans the full Inbox for that run instead of trusting the empty result. Zero results are otherwise indistinguishable from a quiet inbox, which is exactly how the v2.2 bug went unnoticed.
* The sync marker is now stored in UTC and the scan window overlaps by 2 hours instead of 5 minutes, so a timezone change, a DST transition, or a clock correction cannot open a gap that mail falls through.
* Stopped reading the entire archive into memory on every hourly run for a value used at most once per install. At roughly 600 emails a week that read reaches hundreds of MB within a few years.
* Sender and subject are now stripped of line breaks, so a newline in either cannot split one record into two and leave an untrimmable fragment in the rolling windows.
* Opening the Inbox is now error-handled and logged, rather than dying silently with nothing written to the log.
* The scheduled-task check now reports a disabled task or a missing executable path instead of treating any existing task as healthy.

**v2.2**
* **Fixed silent mail loss in every incremental pull.** Outlook's `Restrict()` filter on `[ReceivedTime]` only supports minute precision, and when handed a filter string containing seconds it does not error, it silently matches nothing. Every incremental run returned zero items, and because the sync marker only advances when items are found, it froze too. Verified against a live mailbox holding 5 qualifying items: `'8/26/2026 7:57:57 AM'` returned 0, `'8/26/2026 7:57 AM'` returned 5.
* Fixed the scheduled task refusing to run on battery, and added an in-place repair so existing installs get the fix rather than only fresh setups.
* Fixed timestamps being written with the machine's culture but parsed back with `InvariantCulture`, which would silently stop the rolling windows from trimming on a non-US locale.
* Added an EntryID-based duplicate guard (`seen_ids.txt`), written before the sync marker so an interrupted run fails toward a duplicate rather than a lost email.
* Capped `pull_emails_log.txt` and switched file reads to explicit UTF-8.

**v2.1**
* Added `emails_today.txt`, a rolling 24-hour window used by the daily summary automation instead of the full 10-day `emails_recent.txt`, so daily summaries stay fast regardless of how large your mailbox gets
* Added a one-time progress window shown only on the very first run (the full history pull), since that can take a few minutes on a large mailbox and previously gave no feedback at all
* Fixed a bug where a newly introduced rolling-window file could stay missing forever on an existing install if zero new emails happened to arrive on the first run after updating — it's now seeded from the existing history archive instead

**v2.0**
* Split the single rolling `emails_raw_latest.txt` into a permanent full archive (`emails_history.txt`) and a fast rolling window (`emails_recent.txt`), so summaries stay quick even with high email volume while full history/person/company lookups are still possible
* Self-installing: no more manual Task Scheduler setup, the tool does it on first run
* Runs silently (no console window) in the background
