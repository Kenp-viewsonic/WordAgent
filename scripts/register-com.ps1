# Registers the COM thin shell (WP2) WITHOUT elevation.
#
# regasm /codebase writes HKLM\Software\Classes and needs admin; this script
# writes the identical layout under HKCU\Software\Classes instead (mscoree.dll
# + Assembly/Class/CodeBase), which .NET Framework COM activation honors.
# The Word Addins key (LoadBehavior=3) and the KB948461 lockback bypass key
# are written in the matching registry view (32-bit Office reads WOW6432Node).
# See plans/com-shell-migration-kickoff.md §6.1.
#
# Usage: powershell -ExecutionPolicy Bypass -File scripts\register-com.ps1 [-Arch auto|x86|x64]
param(
    [ValidateSet("auto", "x86", "x64")]
    [string]$Arch = "auto"
)
$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot

$ClsidAddin = "{EC9522B6-2A5E-4E51-9B7D-55CA3C90A5D6}"
$ClsidPane  = "{4342A9CF-0BA6-424E-AEF5-0134DDCD9DFA}"
$ProgIdAddin = "OfficeAgent.Addin"
$ProgIdPane  = "OfficeAgent.TaskPane"
$AssemblyName = "OfficeAgent.ComHost, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null"
$NetCategory = "{62C8FE65-4EBB-45e7-B440-6E39B2CDBF29}"
$LockbackInterface = "{000C0601-0000-0000-C000-000000000046}"

function Get-OfficeArch {
    $c2r = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" -ErrorAction SilentlyContinue
    if ($c2r -and $c2r.Platform) { return $c2r.Platform }
    if (Test-Path "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\16.0\Word\InstallRoot") { return "x86" }
    return "x64"
}

if ($Arch -eq "auto") { $Arch = Get-OfficeArch }
Write-Host "[register-com] target arch: $Arch" -ForegroundColor Cyan

$dll = Join-Path $root "apps\com-host\bin\$Arch\OfficeAgent.ComHost.dll"
if (-not (Test-Path $dll)) { Write-Error "COM host not built: $dll"; exit 1 }
$codeBase = ([System.Uri]$dll).AbsoluteUri

$view = if ($Arch -eq "x64") { [Microsoft.Win32.RegistryView]::Registry64 } else { [Microsoft.Win32.RegistryView]::Registry32 }
$hkcu = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser, $view)
$classes = $hkcu.CreateSubKey("Software\Classes")

function Register-NetComClass($classesKey, $clsid, $progId, $className, $codeBase, $assemblyName) {
    $clsidKey = $classesKey.CreateSubKey("CLSID\$clsid")
    $clsidKey.SetValue("", $className)
    $server = $clsidKey.CreateSubKey("InprocServer32")
    $server.SetValue("", "mscoree.dll")
    $server.SetValue("Class", $className)
    $server.SetValue("Assembly", $assemblyName)
    $server.SetValue("RuntimeVersion", "v4.0.30319")
    $server.SetValue("CodeBase", $codeBase)
    $server.SetValue("ThreadingModel", "Both")
    $server.Close()
    $clsidKey.CreateSubKey("ProgId").SetValue("", $progId)
    $clsidKey.CreateSubKey("Implemented Categories\{62C8FE65-4EBB-45e7-B440-6E39B2CDBF29}") | Out-Null
    $clsidKey.Close()

    $progKey = $classesKey.CreateSubKey($progId)
    $progKey.SetValue("", $className)
    $progKey.CreateSubKey("CLSID").SetValue("", $clsid)
    $progKey.Close()
}

Register-NetComClass $classes $ClsidAddin $ProgIdAddin "OfficeAgent.ComHost.Addin" $codeBase $AssemblyName
Register-NetComClass $classes $ClsidPane  $ProgIdPane  "OfficeAgent.ComHost.AgentTaskPane" $codeBase $AssemblyName

# KB948461: .NET Framework lockback bypass (§6.1 known pitfalls)
$lockback = $classes.CreateSubKey("Interface\$LockbackInterface")
if (-not $lockback.GetValue("")) { $lockback.SetValue("", "Office .NET Framework Lockback Bypass Key") }
$lockback.Close()

# Word Addins key: LoadBehavior=3 (auto-load)
$addins = $hkcu.CreateSubKey("Software\Microsoft\Office\Word\Addins\$ProgIdAddin")
$addins.SetValue("FriendlyName", "Word Agent")
$addins.SetValue("Description", "Word Agent 本地助手（COM 加载项）")
$addins.SetValue("LoadBehavior", 3, [Microsoft.Win32.RegistryValueKind]::DWord)
$addins.Close()
$classes.Close()
$hkcu.Close()

# ── Verification (§6.1) ──────────────────────────────────────────────────────
$failures = 0
$hkcu = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser, $view)
$check = $hkcu.OpenSubKey("Software\Microsoft\Office\Word\Addins\$ProgIdAddin")
if ($check -and [int]$check.GetValue("LoadBehavior") -eq 3) {
    Write-Host "  [OK] Addins key LoadBehavior=3" -ForegroundColor Green
} else {
    Write-Host "  [FAIL] Addins key missing or wrong LoadBehavior" -ForegroundColor Red
    $failures++
}
$server = $hkcu.OpenSubKey("Software\Classes\CLSID\$ClsidAddin\InprocServer32")
if ($server -and $server.GetValue("CodeBase")) {
    Write-Host "  [OK] CLSID InprocServer32 CodeBase=$($server.GetValue('CodeBase'))" -ForegroundColor Green
} else {
    Write-Host "  [FAIL] CLSID registration missing" -ForegroundColor Red
    $failures++
}
$hkcu.Close()

if ($failures -gt 0) { Write-Host "REGISTER FAILED ($failures)" -ForegroundColor Red; exit 1 }
Write-Host "REGISTER OK — 重启 Word 后「开始」选项卡会出现 Word Agent 按钮" -ForegroundColor Green
exit 0
