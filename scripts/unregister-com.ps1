# Unregisters the COM thin shell (WP2) — mirrors scripts\register-com.ps1.
# Usage: powershell -ExecutionPolicy Bypass -File scripts\unregister-com.ps1 [-Arch auto|x86|x64]
param(
    [ValidateSet("auto", "x86", "x64")]
    [string]$Arch = "auto"
)
$ErrorActionPreference = "Stop"

$ClsidAddin = "{EC9522B6-2A5E-4E51-9B7D-55CA3C90A5D6}"
$ClsidPane  = "{4342A9CF-0BA6-424E-AEF5-0134DDCD9DFA}"
$ProgIdAddin = "OfficeAgent.Addin"
$ProgIdPane  = "OfficeAgent.TaskPane"

function Get-OfficeArch {
    $c2r = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" -ErrorAction SilentlyContinue
    if ($c2r -and $c2r.Platform) { return $c2r.Platform }
    if (Test-Path "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\16.0\Word\InstallRoot") { return "x86" }
    return "x64"
}

if ($Arch -eq "auto") { $Arch = Get-OfficeArch }
$view = if ($Arch -eq "x64") { [Microsoft.Win32.RegistryView]::Registry64 } else { [Microsoft.Win32.RegistryView]::Registry32 }
$hkcu = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser, $view)

function Remove-SubKeyTree($parent, $path) {
    try { $parent.DeleteSubKeyTree($path, $false) } catch { }
}

Remove-SubKeyTree $hkcu "Software\Classes\CLSID\$ClsidAddin"
Remove-SubKeyTree $hkcu "Software\Classes\CLSID\$ClsidPane"
Remove-SubKeyTree $hkcu "Software\Classes\$ProgIdAddin"
Remove-SubKeyTree $hkcu "Software\Classes\$ProgIdPane"
Remove-SubKeyTree $hkcu "Software\Microsoft\Office\Word\Addins\$ProgIdAddin"

$check = $hkcu.OpenSubKey("Software\Microsoft\Office\Word\Addins\$ProgIdAddin")
$hkcu.Close()
if ($check) {
    Write-Host "UNREGISTER FAILED — Addins key still present" -ForegroundColor Red
    exit 1
}
Write-Host "UNREGISTER OK" -ForegroundColor Green
exit 0
