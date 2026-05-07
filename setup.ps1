param(
    [switch]$ResetData
)

$ErrorActionPreference = "Stop"

$repoRoot = $PSScriptRoot
$venvDir = Join-Path $repoRoot ".venv"
$pythonExe = Join-Path $venvDir "Scripts\python.exe"
$requirementsFile = Join-Path $repoRoot "requirements.txt"
$stateDir = Join-Path $repoRoot "data"
$stateFile = Join-Path $stateDir "actions_state.json"
$launcherPath = Join-Path $repoRoot "actions-tui.cmd"
$iconPath = Join-Path $repoRoot "assets\actions-tui.ico"
$userBinDir = Join-Path $HOME "bin"
$shimPath = Join-Path $userBinDir "actions-tui.cmd"
$desktopShortcutPath = Join-Path ([Environment]::GetFolderPath("Desktop")) "Actions TUI.lnk"

function Write-Step([string]$Message) {
    Write-Host "==> $Message"
}

function Test-PythonModule([string]$PythonPath, [string]$ModuleName) {
    if (-not (Test-Path $PythonPath)) {
        return $false
    }
    & $PythonPath -c "import $ModuleName" *> $null
    return $LASTEXITCODE -eq 0
}

Write-Step "Creating local virtual environment"
if (-not (Test-Path $pythonExe) -or -not (Test-PythonModule $pythonExe "ensurepip")) {
    if (Test-Path $venvDir) {
        Remove-Item -LiteralPath $venvDir -Recurse -Force
    }
    py -3 -m venv $venvDir
}

Write-Step "Bootstrapping pip in the virtual environment"
& $pythonExe -m ensurepip --upgrade

Write-Step "Installing Python dependencies"
& $pythonExe -m pip install --upgrade pip
& $pythonExe -m pip install -r $requirementsFile

Write-Step "Preparing local data directory"
New-Item -ItemType Directory -Force -Path $stateDir | Out-Null
if ($ResetData -or -not (Test-Path $stateFile)) {
    @'
{
  "version": 1,
  "tables": [],
  "active": [],
  "archive": [],
  "graveyard": [],
  "retired_tables": []
}
'@ | Set-Content -LiteralPath $stateFile -Encoding ASCII
}

Write-Step "Creating user command shim"
New-Item -ItemType Directory -Force -Path $userBinDir | Out-Null
@"
@echo off
setlocal
call "$launcherPath" %*
"@ | Set-Content -LiteralPath $shimPath -Encoding ASCII

$currentUserPath = [Environment]::GetEnvironmentVariable("Path", "User")
$pathEntries = @()
if ($currentUserPath) {
    $pathEntries = $currentUserPath.Split(";", [System.StringSplitOptions]::RemoveEmptyEntries)
}
if ($pathEntries -notcontains $userBinDir) {
    Write-Step "Adding $userBinDir to the user PATH"
    $newUserPath = if ($currentUserPath) { "$currentUserPath;$userBinDir" } else { $userBinDir }
    [Environment]::SetEnvironmentVariable("Path", $newUserPath, "User")
    $pathUpdated = $true
} else {
    $pathUpdated = $false
}

Write-Step "Creating desktop shortcut"
$wsh = New-Object -ComObject WScript.Shell
$shortcut = $wsh.CreateShortcut($desktopShortcutPath)
$shortcut.TargetPath = $shimPath
$shortcut.WorkingDirectory = $repoRoot
if (Test-Path $iconPath) {
    $shortcut.IconLocation = $iconPath
}
$shortcut.Save()

Write-Host ""
Write-Host "Actions TUI is set up."
Write-Host "Desktop shortcut: $desktopShortcutPath"
Write-Host "Terminal command: actions-tui"
if ($pathUpdated) {
    Write-Host "Open a new terminal window before using the actions-tui command."
}
