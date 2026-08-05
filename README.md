# Outlook Email Puller

A small Windows tool that exports your last 7 days of Outlook inbox mail to a plain text file useful for feeding into AI tools (like Claude with a connected folder) that can't read your inbox directly.

## What it does

- Connects to your local Outlook desktop app (via COM automation)
- Pulls every inbox email received in the last 7 days
- Writes them to `emails_raw_latest.txt` inside a `weekly report` folder in your user profile
- Overwrites that file each time you run it, so it always reflects a rolling 7-day window

Output format per email:

```
---
TIME: 2026-08-05 09:48
FROM: Sender Name
SUBJ: Subject line
BODY: First 1500 characters of the email body, flattened to one line
```

## Requirements

- Windows
- Microsoft Outlook desktop app installed and set up with your mail profile (this does **not** work with Outlook web/OWA only it needs the desktop COM interface)

## Usage

### Option A Run the compiled app (easiest)

1. Download `PullEmails.exe` from the [Releases](../../releases) page
2. Double-click it
3. Approve any Windows SmartScreen prompt ("More info" → "Run anyway")  this is expected for an unsigned exe
4. Output lands at `C:\Users\<you>\weekly report\emails_raw_latest.txt`

### Option B Run the PowerShell script directly

```powershell
powershell -ExecutionPolicy Bypass -File "pull_emails_portable.ps1"
```

### Automating it

To avoid running it manually every day, set up a Windows Task Scheduler job:
1. Open **Task Scheduler** → **Create Basic Task**
2. Trigger: Daily, at a time when your machine is on and Outlook can launch
3. Action: Start a program → `powershell` with arguments `-ExecutionPolicy Bypass -File "C:\path\to\pull_emails_portable.ps1"`

## Where the folder comes from

The script auto-creates `weekly report` inside `%USERPROFILE%` (`C:\Users\<you>\weekly report`) if it doesn't already exist no manual setup needed.

## Using it with Claude / Cowork

Connect the `weekly report` folder as a project folder in Claude Desktop (Cowork mode). Once connected, Claude can read `emails_raw_latest.txt` directly. Re-run the script/exe whenever you want fresh data, then ask Claude to re-check the folder — connected folders aren't watched live, so an explicit "check again, I just updated it" ensures it re-reads instead of using a cached view.

## Privacy note

This tool only reads your own inbox and writes to a local file on your own machine — it doesn't send anything over the network. Each person who runs it only gets access to their own mail; there's no way to pull someone else's inbox with this.

## Building the exe yourself

```powershell
Install-Module -Name ps2exe -Scope CurrentUser -Force
Invoke-ps2exe -inputFile "pull_emails_portable.ps1" -outputFile "PullEmails.exe" -noConsole:$false -title "Pull Outlook Emails" -icon "mail_icon.ico"
```
