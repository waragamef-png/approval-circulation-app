param(
  [Parameter(Mandatory = $true)]
  [string]$ProjectPath,

  [Parameter(Mandatory = $true)]
  [string]$NodePath
)

$ErrorActionPreference = "Stop"
Add-Type -AssemblyName PresentationFramework

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
$serverEntry = Join-Path $projectRoot "server\index.mjs"
$builtPage = Join-Path $projectRoot "dist\index.html"
$pidFile = Join-Path $projectRoot ".local-app.pid"
$appUrl = "http://localhost:5173/"
$healthUrl = "http://127.0.0.1:5173/api/health"

if (-not (Test-Path -LiteralPath $serverEntry) -or -not (Test-Path -LiteralPath $builtPage)) {
  Show-ErrorMessage "Required files are missing. Ask Codex to build the local app."
  exit 1
}

try {
  $health = Invoke-WebRequest -UseBasicParsing -Uri $healthUrl -TimeoutSec 2
  if ($health.StatusCode -eq 200) {
    Start-Process $appUrl
    exit 0
  }
} catch {
  # Start a new server below when no healthy server is running.
}

$portOwner = Get-NetTCPConnection -State Listen -LocalPort 5173 -ErrorAction SilentlyContinue
if ($portOwner) {
  Show-ErrorMessage "Port 5173 is already used by another app. Ask Codex to restore the local server."
  exit 1
}

if (Test-Path -LiteralPath $NodePath) {
  $resolvedNode = (Resolve-Path -LiteralPath $NodePath).Path
} else {
  $nodeCommand = Get-Command $NodePath -ErrorAction SilentlyContinue
  if (-not $nodeCommand) {
    Show-ErrorMessage "Node.js was not found. Ask Codex to restore the local server."
    exit 1
  }
  $resolvedNode = $nodeCommand.Source
}

$env:APP_HOST = "127.0.0.1"
$env:APP_PORT = "5173"
$env:APP_ALLOWED_ORIGINS = "http://localhost:5173,http://127.0.0.1:5173"
$env:NODE_ENV = "production"

$serverProcess = Start-Process `
  -FilePath $resolvedNode `
  -ArgumentList ('"' + $serverEntry + '"') `
  -WorkingDirectory $projectRoot `
  -WindowStyle Hidden `
  -PassThru

Set-Content -LiteralPath $pidFile -Value $serverProcess.Id -Encoding ascii

for ($attempt = 0; $attempt -lt 20; $attempt++) {
  Start-Sleep -Milliseconds 250
  try {
    $health = Invoke-WebRequest -UseBasicParsing -Uri $healthUrl -TimeoutSec 2
    if ($health.StatusCode -eq 200) {
      Start-Process $appUrl
      exit 0
    }
  } catch {
    # Wait briefly for server startup.
  }
}

if (-not $serverProcess.HasExited) {
  Stop-Process -Id $serverProcess.Id -Force -ErrorAction SilentlyContinue
}
if (Test-Path -LiteralPath $pidFile) {
  Remove-Item -LiteralPath $pidFile -Force
}

Show-ErrorMessage "The app could not start. Ask Codex to restore the local server."
exit 1
