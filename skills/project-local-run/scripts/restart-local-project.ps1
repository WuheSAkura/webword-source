param(
  [int]$BackendPort = 8010,
  [int]$FrontendPort = 3000
)

$ErrorActionPreference = "Stop"
$ProjectRoot = Resolve-Path (Join-Path $PSScriptRoot "..\..\..")
$BackendDir = Join-Path $ProjectRoot "backend"
$FrontendDir = Join-Path $ProjectRoot "frontend"
$LogDir = Join-Path $ProjectRoot ".codex-runlogs"
$BackendLog = Join-Path $LogDir "backend.out.log"
$BackendErrLog = Join-Path $LogDir "backend.err.log"
$FrontendLog = Join-Path $LogDir "frontend.out.log"
$FrontendErrLog = Join-Path $LogDir "frontend.err.log"

New-Item -ItemType Directory -Force -Path $LogDir | Out-Null

function Stop-PortProcess([int]$Port) {
  $connections = Get-NetTCPConnection -LocalPort $Port -ErrorAction SilentlyContinue
  $processIds = $connections | Select-Object -ExpandProperty OwningProcess -Unique
  foreach ($processId in $processIds) {
    if (-not $processId) { continue }
    try {
      $process = Get-Process -Id $processId -ErrorAction Stop
      Stop-Process -Id $processId -Force
      "Stopped process $processId ($($process.ProcessName)) on port $Port"
    } catch {
      "Could not stop process ${processId} on port ${Port}: $($_.Exception.Message)"
    }
  }
}

function Test-Http([string]$Url) {
  try {
    $response = Invoke-WebRequest -UseBasicParsing $Url -TimeoutSec 8
    return $response.StatusCode
  } catch {
    return "FAILED: $($_.Exception.Message)"
  }
}

Stop-PortProcess -Port $BackendPort
Stop-PortProcess -Port $FrontendPort
Start-Sleep -Seconds 2

"=== Restart $(Get-Date -Format s) ===" | Set-Content -Path $BackendLog
"=== Restart $(Get-Date -Format s) ===" | Set-Content -Path $BackendErrLog
"=== Restart $(Get-Date -Format s) ===" | Set-Content -Path $FrontendLog
"=== Restart $(Get-Date -Format s) ===" | Set-Content -Path $FrontendErrLog

if (-not (Test-Path (Join-Path $FrontendDir "node_modules"))) {
  Push-Location $FrontendDir
  try {
    npm install
  } finally {
    Pop-Location
  }
}

$backendArgs = @(
  "-m", "uvicorn", "app:app", "--host", "0.0.0.0", "--port", "$BackendPort",
  "--reload", "--reload-dir", $BackendDir,
  "--reload-exclude=temp/*",
  "--reload-exclude=template_library.db"
)
Start-Process -FilePath "python" -ArgumentList $backendArgs -WorkingDirectory $BackendDir -RedirectStandardOutput $BackendLog -RedirectStandardError $BackendErrLog -WindowStyle Hidden

$viteCmd = Join-Path $FrontendDir "node_modules\.bin\vite.cmd"
$frontendArgs = @("--host", "0.0.0.0", "--port", "$FrontendPort", "--strictPort")
Start-Process -FilePath $viteCmd -ArgumentList $frontendArgs -WorkingDirectory $FrontendDir -RedirectStandardOutput $FrontendLog -RedirectStandardError $FrontendErrLog -WindowStyle Hidden

Start-Sleep -Seconds 5

$backendStatus = Test-Http "http://localhost:$BackendPort/api/health"
$frontendStatus = Test-Http "http://localhost:$FrontendPort"
$templateCount = $null
$templateWarning = $null
try {
  $templatePayload = Invoke-RestMethod "http://127.0.0.1:$BackendPort/api/ai/templates" -TimeoutSec 10
  $templateCount = @($templatePayload.templates).Count
  if ($templateCount -lt 1) {
    $templateWarning = "AI 模板库为空，请确认范文目录存在并访问 /api/ai/templates/sync"
  } elseif ($templateCount -ne 15) {
    $templateWarning = "AI 模板库条目数为 $templateCount，期望 15；请检查 template_library.db 是否来自 Windows 路径"
  }
} catch {
  $templateWarning = "无法读取 AI 模板库：$($_.Exception.Message)"
}

[PSCustomObject]@{
  BackendUrl = "http://localhost:$BackendPort"
  BackendHealth = $backendStatus
  FrontendUrl = "http://localhost:$FrontendPort"
  FrontendStatus = $frontendStatus
  AiTemplateCount = $templateCount
  AiTemplateWarning = $templateWarning
  BackendLog = $BackendLog
  BackendErrLog = $BackendErrLog
  FrontendLog = $FrontendLog
  FrontendErrLog = $FrontendErrLog
}
