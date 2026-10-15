<#
  Crystal Reports (SAP BusinessObjects) cleanup. No registry backup.
  Usage:  .\cleanup_crystal_reports.ps1 -DryRun   (list only, no admin needed)
          .\cleanup_crystal_reports.ps1           (delete, run as Administrator)
  Uses the .NET RegistryKey API directly (much faster than the PowerShell registry provider).
#>
param([switch]$DryRun)

$sw = [Diagnostics.Stopwatch]::StartNew()
if(-not $DryRun){
    $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    if(-not $isAdmin){ throw 'Run as Administrator (or use -DryRun).' }
}

$hklm = [Microsoft.Win32.RegistryKey]::OpenBaseKey('LocalMachine','Registry64')
$hkcu = [Microsoft.Win32.Registry]::CurrentUser
$match = 'SAP BusinessObjects|crystal|crviewer|craxddrt|ReportSourceBridge|EmfGen|ReportPromptEMF|swebrs|sviewhlp|CRVsPackage|vle\.dll|CRDesigner'

$targets = New-Object System.Collections.Generic.List[object]   # @{Root; Path}
$clsids  = New-Object System.Collections.Generic.HashSet[string]
$tlibs   = New-Object System.Collections.Generic.HashSet[string]

function Add-Target($root, $path){ $script:targets.Add([pscustomobject]@{ Root = $root; Path = $path }) }
function Get-Default($root, $path){
    $k = $root.OpenSubKey($path)
    if($k){ try { return [string]$k.GetValue('') } finally { $k.Dispose() } }
}

# 1. vendor keys
foreach($p in 'SOFTWARE\SAP BusinessObjects','SOFTWARE\WOW6432Node\SAP BusinessObjects'){
    $k = $hklm.OpenSubKey($p); if($k){ $k.Dispose(); Add-Target $hklm $p }
}
$k = $hkcu.OpenSubKey('Software\SAP BusinessObjects'); if($k){ $k.Dispose(); Add-Target $hkcu 'Software\SAP BusinessObjects' }

# 2. CLSIDs (native + WOW6432Node views)
foreach($base in 'SOFTWARE\Classes\CLSID','SOFTWARE\Classes\WOW6432Node\CLSID'){
    $b = $hklm.OpenSubKey($base); if(-not $b){ continue }
    foreach($n in $b.GetSubKeyNames()){
        $d = Get-Default $hklm "$base\$n"
        $v = Get-Default $hklm "$base\$n\InprocServer32"
        if("$v $d" -match $match){ Add-Target $hklm "$base\$n"; [void]$clsids.Add($n) }
    }
    $b.Dispose()
}

# 3. TypeLibs
foreach($base in 'SOFTWARE\Classes\TypeLib','SOFTWARE\Classes\WOW6432Node\TypeLib'){
    $b = $hklm.OpenSubKey($base); if(-not $b){ continue }
    foreach($n in $b.GetSubKeyNames()){
        $tl = $b.OpenSubKey($n); $hit = $false
        foreach($ver in $tl.GetSubKeyNames()){
            $vk = $tl.OpenSubKey($ver)
            foreach($lcid in $vk.GetSubKeyNames()){
                foreach($arch in 'win32','win64'){
                    $ak = $vk.OpenSubKey("$lcid\$arch")
                    if($ak){ if(([string]$ak.GetValue('')) -match 'SAP BusinessObjects|crystal'){ $hit = $true }; $ak.Dispose() }
                }
            }
            $vk.Dispose()
        }
        $tl.Dispose()
        if($hit){ Add-Target $hklm "$base\$n"; [void]$tlibs.Add($n) }
    }
    $b.Dispose()
}

# 4. ProgIDs (by name, or pointing to a matched CLSID) and Interfaces (pointing to a matched TypeLib)
foreach($base in 'SOFTWARE\Classes','SOFTWARE\Classes\WOW6432Node'){
    $b = $hklm.OpenSubKey($base); if(-not $b){ continue }
    foreach($n in $b.GetSubKeyNames()){
        if($n -match '^(CrystalReports|CrystalRuntime|CrystalDecisions|CrystalReport|Crystal\.)'){ Add-Target $hklm "$base\$n"; continue }
        $c = Get-Default $hklm "$base\$n\CLSID"
        if($c -and $clsids.Contains($c)){ Add-Target $hklm "$base\$n" }
    }
    $b.Dispose()
    $ib = $hklm.OpenSubKey("$base\Interface"); if($ib){
        foreach($n in $ib.GetSubKeyNames()){
            $t = Get-Default $hklm "$base\Interface\$n\TypeLib"
            if($t -and $tlibs.Contains($t)){ Add-Target $hklm "$base\Interface\$n" }
        }
        $ib.Dispose()
    }
}

# 5. .NET AssemblyFolders
foreach($base in 'SOFTWARE\Microsoft\.NETFramework\AssemblyFolders','SOFTWARE\WOW6432Node\Microsoft\.NETFramework\AssemblyFolders'){
    $b = $hklm.OpenSubKey($base); if(-not $b){ continue }
    foreach($n in $b.GetSubKeyNames()){ if($n -match 'Crystal'){ Add-Target $hklm "$base\$n" } }
    $b.Dispose()
}

$dirs = @(
    'C:\Program Files (x86)\SAP BusinessObjects',
    'C:\Program Files\SAP BusinessObjects',
    'C:\Windows\Microsoft.NET\assembly\GAC_64\CrystalDecisions.VSDesigner',
    'C:\Windows\Microsoft.NET\assembly\GAC_MSIL\CrystalDecisions.CrystalReports.Design'
) | Where-Object { Test-Path -LiteralPath $_ }

Write-Output "Scan finished in $([int]$sw.Elapsed.TotalSeconds)s"
Write-Output "Registry keys to delete: $($targets.Count)"
$targets | ForEach-Object { "  $(if($_.Root -eq $hkcu){'HKCU'}else{'HKLM'})\$($_.Path)" }
Write-Output "Folders to delete: $($dirs.Count)"; $dirs | ForEach-Object { "  $_" }

if($DryRun){ Write-Output "`nDryRun: nothing changed."; return }

foreach($t in $targets){
    try { $t.Root.DeleteSubKeyTree($t.Path, $false) } catch { Write-Warning "Failed: $($t.Path) - $($_.Exception.Message)" }
}
foreach($d in $dirs){
    Remove-Item -LiteralPath $d -Recurse -Force -ErrorAction Continue
    if(Test-Path -LiteralPath $d){ Write-Warning "Not fully removed: $d" } else { Write-Output "Removed: $d" }
}
Write-Output "Done in $([int]$sw.Elapsed.TotalSeconds)s."
