# Build and run the Azure Function locally for debugging.
# Requires: .NET 8 SDK, Azure Functions Core Tools v4 (npm i -g azure-functions-core-tools@4),
#           and local.settings.json (copy from local.settings.json.example).

$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

if (-not (Test-Path "local.settings.json")) {
    Write-Host "Missing local.settings.json. Copy from local.settings.json.example and fill in values." -ForegroundColor Yellow
    exit 1
}

# Load Values from local.settings.json into env so the host uses them (avoids host reading from wrong folder/Azurite default)
$settings = Get-Content "local.settings.json" -Raw | ConvertFrom-Json
if ($settings.Values) {
    $settings.Values.PSObject.Properties | ForEach-Object { [Environment]::SetEnvironmentVariable($_.Name, $_.Value, "Process") }
    Write-Host "Loaded settings from local.settings.json into environment." -ForegroundColor Green
}

Write-Host "Building..." -ForegroundColor Cyan
dotnet build --configuration Debug
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

# Ensure the host finds our storage connection string (it reads from bin\output when func runs)
New-Item -ItemType Directory -Force -Path "bin\output" | Out-Null
Copy-Item "local.settings.json" "bin\output\" -Force

Write-Host "Starting Functions host (worker output appears below)..." -ForegroundColor Cyan
func start
