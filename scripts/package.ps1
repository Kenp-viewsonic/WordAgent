# Office Agent Desktop - Package Script (WP1 layout)
# Produces dist/: agent.exe + OfficeAgent.Tray.exe + public/ UI + README.
# Usage: npm run package
$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$distDir = Join-Path $root "dist"

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  Office Agent Desktop - Packager" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan

# [1/5] Clean dist
Write-Host "[1/5] Cleaning dist..." -ForegroundColor Yellow
if (Test-Path $distDir) { Remove-Item -Recurse -Force $distDir }
New-Item -ItemType Directory -Path $distDir | Out-Null

# [2/5] Build UI (production) + local-agent type check
Write-Host "[2/5] Building UI (production)..." -ForegroundColor Yellow
Push-Location $root
npm run build:prod
if ($LASTEXITCODE -ne 0) { Pop-Location; Write-Error "UI build failed"; exit 1 }
Pop-Location

# [3/5] Build agent.exe (esbuild + Node SEA)
Write-Host "[3/5] Building agent.exe..." -ForegroundColor Yellow
& (Join-Path $root "scripts\build-agent.ps1")
if ($LASTEXITCODE -ne 0) { Write-Error "agent.exe build failed"; exit 1 }
Copy-Item (Join-Path $root "build\agent.exe") $distDir

# [4/5] Build tray shell
Write-Host "[4/5] Building tray shell..." -ForegroundColor Yellow
Push-Location $root
dotnet build (Join-Path $root "apps\tray\OfficeAgent.Tray.csproj") -c Release --nologo
if ($LASTEXITCODE -ne 0) { Pop-Location; Write-Error "Tray build failed"; exit 1 }
Pop-Location
Copy-Item (Join-Path $root "apps\tray\bin\Release\net48\OfficeAgent.Tray.exe*") $distDir

# [5/5] Assemble dist
Write-Host "[5/5] Assembling dist..." -ForegroundColor Yellow
$publicDir = Join-Path $distDir "public"
New-Item -ItemType Directory -Path $publicDir | Out-Null
Copy-Item -Recurse (Join-Path $root "apps\word-addin\dist\*") $publicDir
Copy-Item (Join-Path $root "scripts\README-dist.md") (Join-Path $distDir "README.md")

Write-Host ""
Write-Host "========================================" -ForegroundColor Green
Write-Host "  Package complete: $distDir" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host "  dist\agent.exe            - local agent service (Node SEA)" -ForegroundColor Cyan
Write-Host "  dist\OfficeAgent.Tray.exe - tray supervisor (auto-start)" -ForegroundColor Cyan
Write-Host "  dist\public\              - UI served at http://127.0.0.1:8787" -ForegroundColor Cyan
