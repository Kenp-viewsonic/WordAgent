# Office Agent Desktop - Package Script
# Produces dist/: agent.exe + OfficeAgent.Tray.exe + public/ UI + com/ (COM host) + README.
# Usage: npm run package
#
# Native tools (npm/dotnet/npx) write progress to stderr, which Windows
# PowerShell turns into errors under EAP=Stop — native calls run with
# EAP=Continue and are checked via $LASTEXITCODE.

$root = Split-Path -Parent $PSScriptRoot
$distDir = Join-Path $root "dist"

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  Office Agent Desktop - Packager" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan

# [1/6] Clean dist
Write-Host "[1/6] Cleaning dist..." -ForegroundColor Yellow
if (Test-Path $distDir) { Remove-Item -Recurse -Force $distDir }
New-Item -ItemType Directory -Path $distDir | Out-Null

# [2/6] Build UI (production) + local-agent type check
Write-Host "[2/6] Building UI (production)..." -ForegroundColor Yellow
Push-Location $root
$ErrorActionPreference = "Continue"
npm run build:prod
$uiExit = $LASTEXITCODE
$ErrorActionPreference = "Stop"
Pop-Location
if ($uiExit -ne 0) { Write-Error "UI build failed"; exit 1 }

# [3/6] Build agent.exe (esbuild + Node SEA)
Write-Host "[3/6] Building agent.exe..." -ForegroundColor Yellow
& (Join-Path $root "scripts\build-agent.ps1")
Copy-Item (Join-Path $root "build\agent.exe") $distDir

# [4/6] Build COM host (x64 + x86; WebView2Loader differs per bitness)
Write-Host "[4/6] Building COM host..." -ForegroundColor Yellow
Push-Location $root
$ErrorActionPreference = "Continue"
dotnet build (Join-Path $root "apps\com-host\OfficeAgent.ComHost.csproj") -c Release -p:PlatformTarget=x64 -p:OutputPath="bin\x64\" --nologo
$x64Exit = $LASTEXITCODE
dotnet build (Join-Path $root "apps\com-host\OfficeAgent.ComHost.csproj") -c Release -p:PlatformTarget=x86 -p:OutputPath="bin\x86\" --nologo
$x86Exit = $LASTEXITCODE
$ErrorActionPreference = "Stop"
Pop-Location
if ($x64Exit -ne 0 -or $x86Exit -ne 0) { Write-Error "COM host build failed"; exit 1 }
foreach ($arch in @("x64", "x86")) {
    $comDir = Join-Path $distDir "com\$arch"
    New-Item -ItemType Directory -Force -Path $comDir | Out-Null
    Copy-Item (Join-Path $root "apps\com-host\bin\$arch\*.dll") $comDir
}

# [5/6] Build tray shell
Write-Host "[5/6] Building tray shell..." -ForegroundColor Yellow
Push-Location $root
$ErrorActionPreference = "Continue"
dotnet build (Join-Path $root "apps\tray\OfficeAgent.Tray.csproj") -c Release --nologo
$trayExit = $LASTEXITCODE
$ErrorActionPreference = "Stop"
Pop-Location
if ($trayExit -ne 0) { Write-Error "Tray build failed"; exit 1 }
Copy-Item (Join-Path $root "apps\tray\bin\Release\net48\OfficeAgent.Tray.exe*") $distDir

# [6/6] Assemble dist
Write-Host "[6/6] Assembling dist..." -ForegroundColor Yellow
$publicDir = Join-Path $distDir "public"
New-Item -ItemType Directory -Path $publicDir | Out-Null
Copy-Item -Recurse (Join-Path $root "apps\word-addin\dist\*") $publicDir
Copy-Item (Join-Path $root "scripts\README-dist.md") (Join-Path $distDir "README.md")
Copy-Item (Join-Path $root "scripts\register-com.ps1") $distDir
Copy-Item (Join-Path $root "scripts\unregister-com.ps1") $distDir

Write-Host ""
Write-Host "========================================" -ForegroundColor Green
Write-Host "  Package complete: $distDir" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host "  dist\agent.exe            - local agent service (Node SEA)" -ForegroundColor Cyan
Write-Host "  dist\OfficeAgent.Tray.exe - tray supervisor (auto-start)" -ForegroundColor Cyan
Write-Host "  dist\public\              - UI served at http://127.0.0.1:8787" -ForegroundColor Cyan
Write-Host "  dist\com\x64|x86          - COM thin shell (regasm /codebase)" -ForegroundColor Cyan
