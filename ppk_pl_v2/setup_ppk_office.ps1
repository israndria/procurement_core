[CmdletBinding()]
param(
    [string]$DriveRoot = '',
    [switch]$VerifyOnly
)

$ErrorActionPreference = 'Stop'
$RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$Generator = Join-Path $PSScriptRoot 'generate_dokumen_ppk.py'
$LauncherSource = Join-Path $PSScriptRoot 'generate_dokumen_ppk.bat'

function Test-PokjaRoot([string]$Path) {
    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path -PathType Container)) {
        return $false
    }
    return (
        (Test-Path (Join-Path $Path 'memory') -PathType Container) -or
        (Test-Path (Join-Path $Path '@ Pejabat Pengadaan 2026') -PathType Container)
    )
}

if (-not (Test-Path -LiteralPath $Generator -PathType Leaf)) {
    throw "Source generator tidak ditemukan: $Generator"
}

$driveCandidates = @(
    $DriveRoot,
    $env:POKJA_DRIVE_ROOT,
    'D:\Dokumen\@ POKJA 2026',
    'G:\Other computers\My Laptop\@ POKJA 2026',
    'C:\POKJA2026'
)
$ResolvedDriveRoot = $driveCandidates |
    Where-Object { Test-PokjaRoot $_ } |
    Select-Object -First 1
if (-not $ResolvedDriveRoot) {
    throw "Root dokumen POKJA tidak ditemukan. Jalankan lagi dengan -DriveRoot 'G:\Other computers\My Laptop\@ POKJA 2026'."
}
$ResolvedDriveRoot = (Resolve-Path -LiteralPath $ResolvedDriveRoot).Path

$pythonCandidates = @(
    $env:POKJA_PYTHON,
    (Join-Path $RepoRoot 'python\python.exe'),
    (Join-Path (Split-Path $RepoRoot -Parent) 'Runtime\WPy64-313110\python\python.exe'),
    'C:\WinPython313\python\python.exe'
)
$PythonExe = $pythonCandidates |
    Where-Object { $_ -and (Test-Path -LiteralPath $_ -PathType Leaf) } |
    Select-Object -First 1
if (-not $PythonExe) {
    $pathPython = Get-Command python.exe -ErrorAction SilentlyContinue
    if ($pathPython) { $PythonExe = $pathPython.Source }
}
if (-not $PythonExe) {
    throw 'Python tidak ditemukan. Sediakan runtime lokal dan set POKJA_PYTHON.'
}
$PythonExe = (Resolve-Path -LiteralPath $PythonExe).Path

$SecretRoot = $env:POKJA_SECRET_ROOT
if ([string]::IsNullOrWhiteSpace($SecretRoot)) {
    $SecretRoot = Join-Path $env:LOCALAPPDATA 'POKJA2026\Secrets'
}
$SecretRoot = [IO.Path]::GetFullPath($SecretRoot)

$PackageRoot = Join-Path $ResolvedDriveRoot 'Paket Experiment - Pengadaan Langsung\V2 - Template PPK PL'
$Workbook = Join-Path $PackageRoot '0. Master_Data_PL_PPK.xlsm'
$RuntimeDir = Join-Path $PackageRoot '__ppk_runtime'
$LauncherTarget = Join-Path $RuntimeDir 'generate_dokumen_ppk.bat'

foreach ($required in @($Workbook, $RuntimeDir, (Join-Path $PackageRoot 'Konstruksi'),
        (Join-Path $PackageRoot 'Perencanaan'), (Join-Path $PackageRoot 'Pengawasan'))) {
    if (-not (Test-Path -LiteralPath $required)) {
        throw "Artefak PPK V2 tidak ditemukan: $required"
    }
}

$launcherNeedsUpdate = (-not (Test-Path -LiteralPath $LauncherTarget -PathType Leaf))
if (-not $launcherNeedsUpdate) {
    $launcherNeedsUpdate = ((Get-FileHash -LiteralPath $LauncherTarget).Hash -ne (Get-FileHash -LiteralPath $LauncherSource).Hash)
}
if ($VerifyOnly -and $launcherNeedsUpdate) {
    throw "Launcher paket belum sama dengan versi portable Git: $LauncherTarget. Jalankan setup tanpa -VerifyOnly untuk memasang dengan backup."
}

if (-not $VerifyOnly) {
    New-Item -ItemType Directory -Force -Path $SecretRoot | Out-Null
    if ((Test-Path -LiteralPath $LauncherTarget) -and
        ((Get-FileHash -LiteralPath $LauncherTarget).Hash -ne (Get-FileHash -LiteralPath $LauncherSource).Hash)) {
        $backup = "$LauncherTarget.backup_$(Get-Date -Format yyyyMMdd_HHmmss)"
        Copy-Item -LiteralPath $LauncherTarget -Destination $backup
        Write-Host "Backup launcher lama: $backup"
    }
    Copy-Item -LiteralPath $LauncherSource -Destination $LauncherTarget -Force

    $env:POKJA_V19_ROOT = $RepoRoot
    $env:POKJA_PYTHON = $PythonExe
    $env:POKJA_DRIVE_ROOT = $ResolvedDriveRoot
    $env:POKJA_SECRET_ROOT = $SecretRoot
    $env:POKJA_PPK_ROOT = $PSScriptRoot
    [Environment]::SetEnvironmentVariable('POKJA_V19_ROOT', $RepoRoot, 'User')
    [Environment]::SetEnvironmentVariable('POKJA_PYTHON', $PythonExe, 'User')
    [Environment]::SetEnvironmentVariable('POKJA_DRIVE_ROOT', $ResolvedDriveRoot, 'User')
    [Environment]::SetEnvironmentVariable('POKJA_SECRET_ROOT', $SecretRoot, 'User')
    [Environment]::SetEnvironmentVariable('POKJA_PPK_ROOT', $PSScriptRoot, 'User')
} else {
    $env:POKJA_V19_ROOT = $RepoRoot
    $env:POKJA_PYTHON = $PythonExe
    $env:POKJA_DRIVE_ROOT = $ResolvedDriveRoot
    $env:POKJA_SECRET_ROOT = $SecretRoot
    $env:POKJA_PPK_ROOT = $PSScriptRoot
}

Write-Host "Python lokal : $PythonExe"
Write-Host "Source lokal : $RepoRoot"
Write-Host "Dokumen Drive: $ResolvedDriveRoot"
Write-Host "Workbook     : $Workbook"
Write-Host "Secret root  : $SecretRoot"

$dependencyCheck = & $PythonExe -c "import openpyxl, docx, win32com.client; print('deps-ok')" 2>&1
if ($LASTEXITCODE -ne 0) {
    throw "Dependency Python/COM tidak lengkap:`n$dependencyCheck"
}
Write-Host "Dependency   : OK"

$officeCheck = & $PythonExe -c "import win32com.client as c; e=c.DispatchEx('Excel.Application'); e.Quit(); w=c.DispatchEx('Word.Application'); w.Quit(); print('office-com-ok')" 2>&1
if ($LASTEXITCODE -ne 0) {
    throw "Excel/Word COM tidak siap:`n$officeCheck"
}
Write-Host "Excel/Word   : COM OK"

$oldPackageRoot = $env:PPK_PACKAGE_ROOT
$oldLogDir = $env:PPK_LOG_DIR
$env:PPK_PACKAGE_ROOT = $PackageRoot
$env:PPK_LOG_DIR = $RuntimeDir
try {
    & $PythonExe $Generator --mode self-check
    if ($LASTEXITCODE -ne 0) { throw 'Self-check generator gagal.' }
} finally {
    $env:PPK_PACKAGE_ROOT = $oldPackageRoot
    $env:PPK_LOG_DIR = $oldLogDir
}

Write-Host ''
if ($VerifyOnly) {
    Write-Host 'VERIFY ONLY OK. Tidak ada environment persisten atau file Drive yang diubah.'
} else {
    Write-Host 'SETUP PPK V2 OK.'
    Write-Host 'Tutup dan buka ulang Excel/Word agar membaca environment baru.'
}
