# Sinkronkan action Task Scheduler ke clone lokal aktif.
$repoDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$scrapeBat = Join-Path $repoDir "Scrape SPSE.bat"

$scrapeAction = New-ScheduledTaskAction `
    -Execute "cmd.exe" `
    -Argument ("/c `"{0}`"" -f $scrapeBat) `
    -WorkingDirectory $repoDir
Set-ScheduledTask -TaskName "POKJA_ScrapeSpse" -Action $scrapeAction | Out-Null

$settings = New-ScheduledTaskSettingsSet `
    -ExecutionTimeLimit (New-TimeSpan -Minutes 15) `
    -StartWhenAvailable `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries
Set-ScheduledTask -TaskName "SyncJadwalTender" -Settings $settings | Out-Null

Write-Output "POKJA_ScrapeSpse -> $scrapeBat"
Write-Output "SyncJadwalTender execution limit -> 15 menit"
