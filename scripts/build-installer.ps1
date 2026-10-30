# Builds the distributable installer (WP4). Prerequisites: npm run package.
# Usage: npm run build:installer
$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot

# 1. 优先用 PATH 里的 iscc（最通用，6/7 都行）
$iscc = (Get-Command iscc -ErrorAction SilentlyContinue).Source

# 2. 退而求其次，搜常见安装路径（同时支持 6 和 7）
if (-not $iscc) {
    $iscc = @(
        "C:\Program Files (x86)\Inno Setup 6\ISCC.exe",
        "C:\Program Files\Inno Setup 6\ISCC.exe",
        "C:\Program Files (x86)\Inno Setup 7\ISCC.exe",
        "C:\Program Files\Inno Setup 7\ISCC.exe",
        (Join-Path $env:LOCALAPPDATA "Programs\Inno Setup 6\ISCC.exe"),
        (Join-Path $env:LOCALAPPDATA "Programs\Inno Setup 7\ISCC.exe")
    ) | Where-Object { Test-Path $_ } | Select-Object -First 1
}

# 3. 再退一步，从 winget 安装目录里搜
if (-not $iscc) {
    $iscc = Get-ChildItem (Join-Path $env:LOCALAPPDATA "Microsoft\WinGet\Packages") -Recurse -Filter "ISCC.exe" -ErrorAction SilentlyContinue | Select-Object -First 1 -ExpandProperty FullName
}

# 4. 都找不到才报错
if (-not $iscc) {
    Write-Error "Inno Setup (ISCC.exe) not found - install via: winget install JRSoftware.InnoSetup"
    exit 1
}
if (-not (Test-Path (Join-Path $root "dist\agent.exe"))) {
    Write-Error "dist\ is not built - run 'npm run package' first"
    exit 1
}

Write-Host "[build-installer] compiling installer..." -ForegroundColor Yellow
& $iscc (Join-Path $root "installer\OfficeAgent.iss")
if ($LASTEXITCODE -ne 0) { Write-Error "ISCC failed"; exit 1 }

Get-ChildItem (Join-Path $root "dist-setup") -Filter "*.exe" | ForEach-Object {
    Write-Host "[build-installer] output: $($_.FullName) ($([Math]::Round($_.Length / 1MB, 1)) MB)" -ForegroundColor Green
}
exit 0
