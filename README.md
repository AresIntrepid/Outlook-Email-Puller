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

If for some reason the self-install fails (check `pull_emails_log.txt` for a warning), you can set it up manually instead: Task Scheduler → Create Basic Task → Trigger: Daily, repeating hourly → Action: Start a program → `powershell` with arguments `-ExecutionPolicy Bypass -File "C:\path\to\pull_emails_portable.ps1"`.

## Where the folder comes from

The script auto-creates a `weekly report` folder inside `%USERPROFILE%` (`C:\Users\<you>\weekly report`) if it doesn't already exist — no manual setup needed. It contains:

* `emails_history.txt` — the permanent, complete archive
* `emails_recent.txt` — the rolling 10-day window
* `emails_today.txt` — the rolling 24-hour window
* `last_sync_state.txt` — internal bookkeeping (the timestamp of the newest email saved so far)
* `pull_emails_log.txt` — a running log of what each run did

## Using it with Claude / Cowork

Connect the `weekly report` folder as a project folder in Claude Desktop (Cowork mode). A sensible default split:

* A **daily** summary automation should read `emails_today.txt` — it's already limited to the last 24 hours, so no further date filtering is needed
* A **weekly** summary automation should read `emails_recent.txt` — filter within it to whatever week window you want (e.g. Saturday through Friday)
* Anything else — a question naming a specific person or company, a request for full history, or anything the recent files don't answer — should fall back to `emails_history.txt`, searched with no date cutoff

**Important:** connected folders aren't watched live — Claude can end up working from a cached view of a file instead of what's actually on disk right now. Build your instructions to explicitly say something like *"always re-read this file fresh — never rely on a cached or previously-read version"* rather than assuming it'll notice the file changed on its own.

The exact project instructions and automation prompts used in the reference setup are in [`cowork-instructions.md`](./cowork-instructions.md) — copy them in as-is rather than reconstructing them from scratch.

## Privacy note

This tool only reads your own inbox and writes to a local file on your own machine — it doesn't send anything over the network. Each person who runs it only gets access to their own mail; there's no way to pull someone else's inbox with this.

## Building the exe yourself

```
Install-Module -Name ps2exe -Scope CurrentUser -Force
Invoke-ps2exe -inputFile "pull_emails_portable.ps1" -outputFile "PullEmails.exe" -noConsole -STA -title "Pull Outlook Emails" -icon "mail_icon.ico"
```

Notes on the flags: `-noConsole` (not `-noConsole:$false`) because this runs silently once an hour in the background via the scheduled task, and a console window popping up unannounced every hour would be unwelcome. `-STA` is needed because the first-run progress window uses Windows Forms, which requires a single-threaded apartment to behave reliably alongside Outlook's own COM automation.

## Version history

**v2.1**
* Added `emails_today.txt`, a rolling 24-hour window used by the daily summary automation instead of the full 10-day `emails_recent.txt`, so daily summaries stay fast regardless of how large your mailbox gets
* Added a one-time progress window shown only on the very first run (the full history pull), since that can take a few minutes on a large mailbox and previously gave no feedback at all
* Fixed a bug where a newly introduced rolling-window file could stay missing forever on an existing install if zero new emails happened to arrive on the first run after updating — it's now seeded from the existing history archive instead

**v2.0**
* Split the single rolling `emails_raw_latest.txt` into a permanent full archive (`emails_history.txt`) and a fast rolling window (`emails_recent.txt`), so summaries stay quick even with high email volume while full history/person/company lookups are still possible
* Self-installing: no more manual Task Scheduler setup, the tool does it on first run
* Runs silently (no console window) in the background
