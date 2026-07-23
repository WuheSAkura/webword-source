param(
  [string]$Service = "webword",
  [string]$BaseUrl = "http://localhost:8000"
)

$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent (Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path))
Set-Location $projectRoot

Write-Host "Checking Docker engine..."
docker info | Out-Null

Write-Host "Building and redeploying $Service..."
docker compose up -d --build $Service

Write-Host "Container status:"
docker ps --filter "name=$Service" --format "table {{.Names}}`t{{.Image}}`t{{.Ports}}`t{{.Status}}"

function Invoke-WithRetry {
  param(
    [string]$Uri,
    [int]$Attempts = 20,
    [int]$DelaySeconds = 2
  )

  $lastError = $null
  for ($i = 1; $i -le $Attempts; $i++) {
    try {
      return Invoke-WebRequest -UseBasicParsing $Uri
    } catch {
      $lastError = $_
      Start-Sleep -Seconds $DelaySeconds
    }
  }
  throw $lastError
}

function Invoke-RestWithRetry {
  param(
    [string]$Uri,
    [int]$Attempts = 20,
    [int]$DelaySeconds = 2
  )

  $lastError = $null
  for ($i = 1; $i -le $Attempts; $i++) {
    try {
      return Invoke-RestMethod -Uri $Uri
    } catch {
      $lastError = $_
      Start-Sleep -Seconds $DelaySeconds
    }
  }
  throw $lastError
}

Write-Host "Checking health endpoint..."
$health = Invoke-WithRetry "$BaseUrl/api/health"
if ($health.StatusCode -ne 200) {
  throw "Health check failed: $($health.StatusCode)"
}

Write-Host "Checking homepage..."
$homeResponse = Invoke-WithRetry $BaseUrl
if ($homeResponse.StatusCode -ne 200) {
  throw "Homepage check failed: $($homeResponse.StatusCode)"
}

Write-Host "Checking template library..."
$templates = Invoke-RestWithRetry "$BaseUrl/api/ai/templates"
$count = @($templates.templates).Count
if ($count -ne 15) {
  throw "Template library expected 15 entries, got $count"
}

Write-Host "Recent container logs:"
docker logs --tail 80 $Service

Write-Host "Docker deployment verified: $BaseUrl"
