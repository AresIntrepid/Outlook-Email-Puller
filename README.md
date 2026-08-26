# Outlook Email Puller

A small Windows tool that exports your Outlook inbox to plain text files, useful for feeding into AI tools (like Claude with a connected folder) that can't read your inbox directly.

## What it does

* Connects to your local Outlook desktop app (via COM automation)
* On first run, pulls your **entire inbox history** and writes it to `emails_history.txt` — a permanent archive that's never trimmed
* On every run after that, pulls only what's arrived since the last run and appends it to `emails_history.txt`
* Also maintains two rolling-window files, kept in sync with the archive but automatically trimmed every run so they stay fast to read no matter how large `emails_history.txt` grows:
  * `emails_recent.txt` — last 10 days
  * `emails_today.txt` — last 24 hours (this is the one a daily summary should read, so it only ever has to process roughly a day's worth of data instead of a full 10-day window)
* **Sets itself up to run automatically.** The very first time it runs on a machine, it registers a Windows Scheduled Task (`PullEmailsAutoSync`) that repeats hourly from 6am–midnight, wakes the computer if it's asleep, and catches up if a run gets missed — so after the first double-click, you never have to run it manually again
* **Shows progress on the one run that's slow.** The very first run (the full history pull) can take a few minutes on a large mailbox, so it shows a small progress window for that run only. Every run after that — including all future hourly background runs — is fast and completely silent, no window at all

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
* `pull_emails_log.txt` — a running log of what each run did

## Using it with Claude / Cowork

Connect the `weekly report` folder as a project folder in Claude Desktop (Cowork mode). A sensible default split:

* A **daily** summary automation should read `emails_today.txt` — it's already limited to the last 24 hours, so no further date filtering is needed
* A **weekly** summary automation should read `emails_recent.txt` — filter within it to whatever week window you want (e.g. Saturday through Friday)
* Anything else — a question naming a specific person or company, a request for full history, or anything the recent files don't answer — should fall back to `emails_history.txt`, searched with no date cutoff

**Important:** connected folders aren't watched live — Claude can end up working from a cached view of a file instead of what's actually on disk right now. Build your instructions to explicitly say something like *"always re-read this file fresh — never rely on a cached or previously-read version"* rather than assuming it'll notice the file changed on its own.


### Project instructions

Applies to every chat in the project, including ad-hoc one-off questions.

```
Emails are saved in three files in this folder: emails_today.txt (last 24 hours, the fastest and smallest), emails_recent.txt (last 10 days), and emails_history.txt (the complete, unbounded archive going back to when tracking started). Use emails_today.txt for anything about today or "since I last checked." Use emails_recent.txt for anything else recent — this week, the last several days — and as the default when I don't specify a period. Use emails_history.txt instead whenever I ask about a specific person or company by name — including phrasings like "full picture on X," "everything from X," "history with X," or "how many times has X emailed me" — or whenever I explicitly ask you to look back further than 10 days, or the answer isn't found in the other two files. When using emails_history.txt, search the entire file with no date cutoff unless I specify a range. All three files are updated outside this conversation by a scheduled script — always re-read them fresh each time; never rely on a cached or previously-read version. If necessary, update all three files using the exe/script provided in this folder. Never ask about connecting an email account.
```

### Daily automation

Runs once a day. Reads `emails_today.txt`, which is already pre-filtered to the last 24 hours by the script, so no date filtering happens in the prompt itself.

```
Summarize the emails in the given period. Read emails_today.txt in this folder directly — never ask about connecting an email account. This file is updated outside this conversation by a scheduled script, so always re-read it fresh; never rely on a cached or previously-read version. It's a rolling window of the last 24 hours, appended in chronological order (oldest first, newest last) — include every entry in the file, since it's already limited to the right window. Title it "Inbox summary ([date range])". Group emails into these five sections, in this order, and include every section even if empty (say "None" rather than omitting it): Needs action — awaiting a reply or decision from me. Resolved — was actionable but already handled (note how/when). Overdue — deadline has passed with no action taken. FYI — informational, no action needed. Noise — automated alerts, verification codes, spam. Mention only as a count, don't list individually. 1–2 sentence bullets per item. Output the summary directly as plain text in your reply — do not create or attach a document, artifact, or file for it.
```

### Weekly automation

Runs once a week (Friday). Reads `emails_recent.txt` (the 10-day window) and filters within the prompt to the Saturday–Friday range, since the file itself isn't pre-filtered that tightly.

```
Summarize the emails from the past week. Read emails_recent.txt in this folder directly — never ask about connecting an email account. This file is updated outside this conversation by a scheduled script, so always re-read it fresh; never rely on a cached or previously-read version. It's a rolling window of the last 10 days, appended in chronological order (oldest first, newest last) — only consider entries with a TIME from last Saturday through today (Friday); ignore every older entry still sitting in the file. Title it "Weekly inbox summary ([Sat–Fri date range])". Group emails into these five sections, in this order, and include every section even if empty (say "None" rather than omitting it): Needs action — awaiting a reply or decision from me. Resolved — was actionable but already handled (note how/when). Overdue — deadline has passed with no action taken. FYI — informational, no action needed. Noise — automated alerts, verification codes, spam. Mention only as a count, don't list individually. 1–2 sentence bullets per item. Output the summary directly as plain text in your reply — do not create or attach a document, artifact, or file for it.
```

## Privacy note

This tool only reads your own inbox and writes to a local file on your own machine — it doesn't send anything over the network. Each person who runs it only gets access to their own mail; there's no way to pull someone else's inbox with this.

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

## Version history

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
