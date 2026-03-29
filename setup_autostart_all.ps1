# Setup Task Scheduler untuk semua Streamlit apps POKJA 2026
# Jalankan sekali: klik kanan > Run with PowerShell

$BASE = "D:\Dokumen\@ POKJA 2026\V19_Scheduler\WPy64-313110"
$WINPYTHON = "$BASE\python\pythonw.exe"
$SYSPYTHON = "C:\Users\MSI\AppData\Local\Programs\Python\Python312\pythonw.exe"

$apps = @(
    @{
        Name    = "V19_Scheduler_Server"
        Desc    = "V19 Tender Scheduler"
        Python  = $WINPYTHON
        AppPy   = "$BASE\V19_Scheduler.py"
        WorkDir = $BASE
        Port    = 8500
    },
    @{
        Name    = "V20_Scrapper_Server"
        Desc    = "V20 SPSE Scrapper"
        Python  = $WINPYTHON
        AppPy   = "$BASE\V20_Scrapper.py"
        WorkDir = $BASE
        Port    = 8501
    },
    @{
        Name    = "V21_InaprocScraper_Server"
        Desc    = "V21 Inaproc Scraper"
        Python  = $WINPYTHON
        AppPy   = "D:\Dokumen\@ POKJA 2026\Inaproc_Scraper\app.py"
        WorkDir = "D:\Dokumen\@ POKJA 2026\Inaproc_Scraper"
        Port    = 8502
    },
    @{
        Name    = "V22_OrderBot_Server"
        Desc    = "V22 Inaproc Order Bot"
        Python  = $SYSPYTHON
        AppPy   = "$BASE\V22_InaprocOrder\app.py"
        WorkDir = "$BASE\V22_InaprocOrder"
        Port    = 8503
    },
    @{
        Name    = "AsistenPokja_Server"
        Desc    = "Asisten Pokja AI"
        Python  = $SYSPYTHON
        AppPy   = "$BASE\Asisten_Pokja\app.py"
        WorkDir = "$BASE\Asisten_Pokja"
        Port    = 8504
    }
)

Write-Host ""
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "  Setup Autostart Semua Streamlit Apps" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""

foreach ($app in $apps) {
    Write-Host "Mendaftarkan: $($app.Desc) (port $($app.Port))..." -ForegroundColor Yellow

    $existing = Get-ScheduledTask -TaskName $app.Name -ErrorAction SilentlyContinue
    if ($existing) {
        Unregister-ScheduledTask -TaskName $app.Name -Confirm:$false
    }

    $trigger = New-ScheduledTaskTrigger -AtLogOn -User $env:USERNAME

    $action = New-ScheduledTaskAction `
        -Execute $app.Python `
        -Argument "-m streamlit run `"$($app.AppPy)`" --server.port $($app.Port) --server.headless true --browser.gatherUsageStats false" `
        -WorkingDirectory $app.WorkDir

    $settings = New-ScheduledTaskSettingsSet `
        -AllowStartIfOnBatteries `
        -DontStopIfGoingOnBatteries `
        -StartWhenAvailable `
        -RestartInterval (New-TimeSpan -Minutes 1) `
        -RestartCount 3 `
        -ExecutionTimeLimit (New-TimeSpan -Days 0)

    Register-ScheduledTask `
        -TaskName $app.Name `
        -Trigger $trigger `
        -Action $action `
        -Settings $settings `
        -Description $app.Desc `
        -RunLevel Limited | Out-Null

    Start-ScheduledTask -TaskName $app.Name
    Write-Host "  OK - http://localhost:$($app.Port)" -ForegroundColor Green
}

Write-Host ""
Write-Host "============================================" -ForegroundColor Green
Write-Host "  Semua task terdaftar dan dijalankan!" -ForegroundColor Green
Write-Host "============================================" -ForegroundColor Green
Write-Host ""
Write-Host "Daftar URL:" -ForegroundColor Cyan
Write-Host "  V19 Tender Scheduler  - http://localhost:8500"
Write-Host "  V20 SPSE Scrapper     - http://localhost:8501"
Write-Host "  V21 Inaproc Scraper   - http://localhost:8502"
Write-Host "  V22 Order Bot         - http://localhost:8503"
Write-Host "  Asisten Pokja         - http://localhost:8504"
Write-Host ""
Write-Host "Tunggu ~20 detik lalu buka di Chrome." -ForegroundColor Yellow
pause
