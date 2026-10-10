# WP1 smoke: verifies agent health + security posture (token auth, CORS
# whitelist, token bootstrap). Usage: npm run smoke
param(
    [string]$Exe = "",
    [int]$Port = 8787
)
$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $PSScriptRoot
if (-not $Exe) {
    $Exe = @(
        (Join-Path $root "dist\agent.exe"),
        (Join-Path $root "build\agent.exe")
    ) | Where-Object { Test-Path $_ } | Select-Object -First 1
}
$base = "http://127.0.0.1:$Port"
$tokenFile = Join-Path $env:LOCALAPPDATA "OfficeAgent\agent.token"
$script:failed = $false
$started = $false
$proc = $null

function Assert($cond, $name) {
    if ($cond) {
        Write-Host "  [PASS] $name" -ForegroundColor Green
    } else {
        Write-Host "  [FAIL] $name" -ForegroundColor Red
        $script:failed = $true
    }
}

try {
    # Stage the UI next to the exe when missing (build\ layout); the packaged
    # dist\ layout already contains public\.
    if ($Exe) {
        $exeDir = Split-Path -Parent $Exe
        $exePublic = Join-Path $exeDir "public"
        $uiDist = Join-Path $root "apps\word-addin\dist"
        if (-not (Test-Path $exePublic) -and (Test-Path $uiDist)) {
            Copy-Item -Recurse $uiDist $exePublic
        }
    }

    $health = $null
    try { $health = Invoke-WebRequest -UseBasicParsing "$base/health" -TimeoutSec 2 } catch { }
    if ($health -and $health.StatusCode -eq 200) {
        Write-Host "  [INFO] reusing running agent on $base"
    } else {
        if (-not $Exe -or -not (Test-Path $Exe)) {
            Write-Error "agent.exe not found - run scripts\build-agent.ps1 first"
        }
        Write-Host "  [INFO] starting $Exe"
        $proc = Start-Process $Exe -PassThru -WindowStyle Hidden
        $started = $true
        $deadline = (Get-Date).AddSeconds(30)
        do {
            Start-Sleep -Milliseconds 500
            try { $health = Invoke-WebRequest -UseBasicParsing "$base/health" -TimeoutSec 2 } catch { $health = $null }
        } until (($health -and $health.StatusCode -eq 200) -or (Get-Date) -gt $deadline)
    }

    $health = $null
    try { $health = Invoke-WebRequest -UseBasicParsing "$base/health" -TimeoutSec 5 } catch { }
    Assert ($health -and [int]$health.StatusCode -eq 200) "/health returns 200"

    $unauth = $null
    try { Invoke-WebRequest -UseBasicParsing "$base/v1/kb/stats" -TimeoutSec 5 | Out-Null } catch { $unauth = $_.Exception.Response }
    Assert ($unauth -and [int]$unauth.StatusCode -eq 401) "/v1/* without token rejected (401)"

    $wrong = $null
    try { Invoke-WebRequest -UseBasicParsing "$base/v1/kb/stats" -Headers @{ "X-Agent-Token" = "wrong-token" } -TimeoutSec 5 | Out-Null } catch { $wrong = $_.Exception.Response }
    Assert ($wrong -and [int]$wrong.StatusCode -eq 401) "/v1/* with wrong token rejected (401)"

    Assert (Test-Path $tokenFile) "token file exists: $tokenFile"
    if (Test-Path $tokenFile) {
        $token = (Get-Content $tokenFile -Raw).Trim()
        $ok = $null
        try { $ok = Invoke-WebRequest -UseBasicParsing "$base/v1/kb/stats" -Headers @{ "X-Agent-Token" = $token } -TimeoutSec 5 } catch { }
        Assert ($ok -and [int]$ok.StatusCode -eq 200) "/v1/* with valid token accepted (200)"
    }

    $cors = $null
    try { $cors = Invoke-WebRequest -UseBasicParsing "$base/health" -Headers @{ "Origin" = "https://evil.example" } -TimeoutSec 5 } catch { }
    $acao = if ($cors) { $cors.Headers["Access-Control-Allow-Origin"] } else { $null }
    Assert ($cors -and ($null -eq $acao -or $acao -ne "https://evil.example")) "CORS does not reflect foreign origins"

    $page = $null
    try { $page = Invoke-WebRequest -UseBasicParsing "$base/" -TimeoutSec 5 } catch { }
    Assert ($page -and $page.Content -match "__AGENT_TOKEN__") "served HTML bootstraps the agent token"
}
finally {
    if ($started -and $proc) {
        Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue
    }
}

if ($script:failed) {
    Write-Host "SMOKE FAILED" -ForegroundColor Red
    exit 1
}
Write-Host "SMOKE PASSED" -ForegroundColor Green
