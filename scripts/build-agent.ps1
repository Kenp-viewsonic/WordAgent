# Builds apps/local-agent into a standalone agent.exe via esbuild + Node SEA
# (see plans/com-shell-migration-kickoff.md §6.5). Usage: npm run build:agent-exe
#
# Native tools (npx/node/postject) write progress to stderr, which Windows
# PowerShell turns into errors under EAP=Stop — so native calls run with
# EAP=Continue and are checked via $LASTEXITCODE instead.

$root = Split-Path -Parent $PSScriptRoot
$build = Join-Path $root "build"

Write-Host "[build-agent] bundling local-agent..." -ForegroundColor Yellow
if (Test-Path $build) { Remove-Item -Recurse -Force $build }
New-Item -ItemType Directory -Path $build | Out-Null

Push-Location $root
$ErrorActionPreference = "Continue"
npx esbuild "apps/local-agent/src/server.ts" --bundle --platform=node --target=node20 --format=cjs --outfile="$build/agent-bundle.js"
$bundleExit = $LASTEXITCODE
$ErrorActionPreference = "Stop"
Pop-Location
if ($bundleExit -ne 0) { Write-Error "esbuild bundle failed"; exit 1 }

Write-Host "[build-agent] generating SEA blob..." -ForegroundColor Yellow
$seaConfig = @{
    main                      = "agent-bundle.js"
    output                    = "sea-prep.blob"
    disableExperimentalSEAWarning = $true
    useCodeCache              = $false
} | ConvertTo-Json
Set-Content -Path (Join-Path $build "sea-config.json") -Value $seaConfig -Encoding UTF8

Push-Location $build
$ErrorActionPreference = "Continue"
node --experimental-sea-config sea-config.json
$blobExit = $LASTEXITCODE
if ($blobExit -eq 0) {
    Write-Host "[build-agent] injecting blob into agent.exe..." -ForegroundColor Yellow
    Copy-Item (Get-Command node).Source (Join-Path $build "agent.exe")
    # postject prints "The signature seems corrupted!" for Authenticode-signed
    # node.exe — expected and harmless (re-sign in WP5 if needed).
    npx postject agent.exe NODE_SEA_BLOB sea-prep.blob --sentinel-fuse NODE_SEA_FUSE_fce680ab2cc467b6e072b8b5df1996b2
    $injectExit = $LASTEXITCODE
} else {
    $injectExit = 1
}
$ErrorActionPreference = "Stop"
Pop-Location
if ($blobExit -ne 0) { Write-Error "SEA blob generation failed"; exit 1 }
if ($injectExit -ne 0) { Write-Error "postject injection failed"; exit 1 }

Write-Host "[build-agent] done: $build\agent.exe" -ForegroundColor Green
exit 0
