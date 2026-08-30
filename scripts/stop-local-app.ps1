param(
  [Parameter(Mandatory = $true)]
  [string]$ProjectPath
)

$ErrorActionPreference = "Stop"
Add-Type -AssemblyName PresentationFramework

function Show-InfoMessage {
  param([string]$Message)
  [System.Windows.MessageBox]::Show(
    $Message,
    "Approval Circulation",
    [System.Windows.MessageBoxButton]::OK,
    [System.Windows.MessageBoxImage]::Information
  ) | Out-Null
}

function Show-ErrorMessage {
  param([string]$Message)
  [System.Windows.MessageBox]::Show(
    $Message,
    "Approval Circulation",
    [System.Windows.MessageBoxButton]::OK,
    [System.Windows.MessageBoxImage]::Error
  ) | Out-Null
}

$projectRoot = (Resolve-Path -LiteralPath $ProjectPath).Path
$pidFile = Join-Path $projectRoot ".local-app.pid"

if (-not (Test-Path -LiteralPath $pidFile)) {
  Show-InfoMessage "The server started by this launcher is not running."
  exit 0
}

$serverPidText = (Get-Content -Raw -LiteralPath $pidFile).Trim()
$serverPid = 0
if (-not [int]::TryParse($serverPidText, [ref]$serverPid)) {
  Show-ErrorMessage "The server process could not be verified. Ask Codex to stop the local server."
  exit 1
}

$serverProcess = Get-CimInstance Win32_Process -Filter "ProcessId = $serverPid" -ErrorAction SilentlyContinue
if (-not $serverProcess) {
  Remove-Item -LiteralPath $pidFile -Force
  Show-InfoMessage "The server is already stopped."
  exit 0
}

$normalizedCommand = ($serverProcess.CommandLine -replace '/', '\').ToLowerInvariant()
if ($serverProcess.Name -notmatch '^node(\.exe)?$' -or $normalizedCommand -notlike '*server\index.mjs*') {
  Show-ErrorMessage "The server process could not be verified safely. Ask Codex to stop the local server."
  exit 1
}

Stop-Process -Id $serverPid -Force
Remove-Item -LiteralPath $pidFile -Force
Show-InfoMessage "Approval Circulation has stopped."
