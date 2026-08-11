param(
    [switch]$Once
)

$ErrorActionPreference = 'SilentlyContinue'

$mutex = New-Object -TypeName System.Threading.Mutex -ArgumentList $false, 'POKJA2026PreviewUnblockWatcher'
$ownsMutex = $false
try {
    $ownsMutex = $mutex.WaitOne(0)
} catch [System.Threading.AbandonedMutexException] {
    $ownsMutex = $true
}
if (-not $ownsMutex) { exit 0 }

$shellFolders = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders'
$downloads = (Get-ItemProperty -Path $shellFolders -ErrorAction Stop).'{374DE290-123F-4565-9164-39C4925E467B}'
$downloads = [Environment]::ExpandEnvironmentVariables($downloads)
if (-not (Test-Path -LiteralPath $downloads -PathType Container)) { exit 1 }

$extensions = @(
    '.pdf', '.doc', '.docx', '.docm', '.xls', '.xlsx', '.xlsm',
    '.ppt', '.pptx', '.pptm', '.txt', '.md', '.markdown', '.rtf', '.csv'
)
$logDir = Join-Path $env:LOCALAPPDATA 'POKJA2026'
New-Item -ItemType Directory -Path $logDir -Force | Out-Null
$logFile = Join-Path $logDir 'preview-download-unblock.log'

function Unblock-PreviewFile {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return }
    $extension = [IO.Path]::GetExtension($Path).ToLowerInvariant()
    if ($extensions -notcontains $extension) { return }

    for ($attempt = 0; $attempt -lt 10; $attempt++) {
        try {
            $zone = Get-Item -LiteralPath $Path -Stream Zone.Identifier -ErrorAction SilentlyContinue
            if ($zone) {
                Unblock-File -LiteralPath $Path -ErrorAction Stop
                Add-Content -LiteralPath $logFile -Value "$(Get-Date -Format s) UNBLOCK $Path"
            }
            return
        } catch {
            Start-Sleep -Milliseconds 500
        }
    }
    Add-Content -LiteralPath $logFile -Value "$(Get-Date -Format s) RETRY_EXHAUSTED $Path"
}

if ($Once) {
    Get-ChildItem -LiteralPath $downloads -Recurse -File -ErrorAction SilentlyContinue |
        ForEach-Object { Unblock-PreviewFile -Path $_.FullName }
    exit 0
}

$watcher = New-Object IO.FileSystemWatcher
$watcher.Path = $downloads
$watcher.Filter = '*'
$watcher.IncludeSubdirectories = $true
$watcher.InternalBufferSize = 65536
$watcher.NotifyFilter = [IO.NotifyFilters]::FileName -bor [IO.NotifyFilters]::CreationTime -bor [IO.NotifyFilters]::LastWrite -bor [IO.NotifyFilters]::Size
$watcher.EnableRaisingEvents = $true

$action = {
    $path = $Event.SourceEventArgs.FullPath
    Unblock-PreviewFile -Path $path
}
$errorAction = {
    Add-Content -LiteralPath $logFile -Value "$(Get-Date -Format s) WATCHER_ERROR $($Event.SourceEventArgs.GetException().Message)"
}

Register-ObjectEvent -InputObject $watcher -EventName Created -Action $action | Out-Null
Register-ObjectEvent -InputObject $watcher -EventName Changed -Action $action | Out-Null
Register-ObjectEvent -InputObject $watcher -EventName Renamed -Action $action | Out-Null
Register-ObjectEvent -InputObject $watcher -EventName Error -Action $errorAction | Out-Null
Add-Content -LiteralPath $logFile -Value "$(Get-Date -Format s) WATCHER_ACTIVE $downloads"

try {
    while ($true) { Wait-Event -Timeout 60 | Out-Null }
} finally {
    $watcher.EnableRaisingEvents = $false
    $watcher.Dispose()
    $mutex.ReleaseMutex()
    $mutex.Dispose()
}
