# Rolling last-7-days Outlook export. Silent version — no console output, no prompts.
# Works for any user - paths are relative to their own profile.
$untilDt = (Get-Date).Date.AddDays(1)   # today, end of day
$sinceDt = $untilDt.AddDays(-7)          # 7 days back
$outDir = Join-Path $env:USERPROFILE "weekly report"
if (-not (Test-Path $outDir)) {
    New-Item -ItemType Directory -Path $outDir | Out-Null
}
$outFile = Join-Path $outDir "emails_raw_latest.txt"
$logFile = Join-Path $outDir "pull_emails_log.txt"

function Write-Log($msg) {
    Add-Content -Path $logFile -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - $msg"
}

try {
    $o = New-Object -ComObject Outlook.Application
} catch {
    Write-Log "ERROR: Could not connect to Outlook. Make sure Outlook is installed and set up on this machine."
    exit 1
}

$ns = $o.GetNamespace("MAPI")
$inbox = $ns.GetDefaultFolder(6)
$items = $inbox.Items
$items.Sort("[ReceivedTime]", $true)

$sb = New-Object System.Text.StringBuilder
$scanMax = [Math]::Min($items.Count, 600)
$found = 0
for ($i = 1; $i -le $scanMax; $i++) {
    try {
        $it = $items.Item($i)
        $rt = $it.ReceivedTime
        if ($rt -lt $sinceDt) { break }
        if ($rt -gt $untilDt) { continue }
        $time = $rt.ToString("yyyy-MM-dd HH:mm")
        $from = $it.SenderName
        $subj = $it.Subject
        $body = $it.Body
        if ($body -and $body.Length -gt 1500) { $body = $body.Substring(0, 1500) }
        $body = ($body -replace "\r?\n", " ").Trim()
        [void]$sb.AppendLine("---")
        [void]$sb.AppendLine("TIME: $time")
        [void]$sb.AppendLine("FROM: $from")
        [void]$sb.AppendLine("SUBJ: $subj")
        [void]$sb.AppendLine("BODY: $body")
        [void]$sb.AppendLine("")
        $found++
    } catch {
        Write-Log "Error at item $i"
    }
}

[System.IO.File]::WriteAllText($outFile, $sb.ToString(), [System.Text.Encoding]::UTF8)
Write-Log "Done. Output: $outFile ($found items found)"