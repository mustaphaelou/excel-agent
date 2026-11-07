# ============================================================================
#  deploy.ps1 - put ExcelAgent.xlam where Excel will load it on every start.
#
#  Target: %APPDATA%\Microsoft\Excel\XLSTART - Excel opens everything in this
#  folder at startup, and it is already one of Excel's Trusted Locations, so
#  the macros run without a security prompt. No registry entries, no new trust
#  settings, nothing to clean up but the file itself.
#
#  (The AddIns folder was tried first: a COM-launched Excel writes
#  OPEN="ExcelAgent.xlam" as a *relative* path, which does not resolve, and the
#  folder is not trusted anyway.)
# ============================================================================
$ErrorActionPreference = 'Continue'

$Src     = '\\wsl.localhost\Ubuntu\home\eloaurdi\hay2010agent\excel-agent'
$Build   = Join-Path $Src 'build\ExcelAgent.xlam'
$AddIns  = Join-Path $env:APPDATA 'Microsoft\AddIns'
$XLStart = Join-Path $env:APPDATA 'Microsoft\Excel\XLSTART'
$Target  = Join-Path $XLStart 'ExcelAgent.xlam'

if (-not (Test-Path $Build)) { throw "Missing $Build - run build.ps1 and inject_ribbon.py first." }

# Drop the old AddIns registration so the add-in cannot load twice.
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
try {
    $ai = $xl.AddIns.Item('ExcelAgent')
    $ai.Installed = $false
    Write-Output "cleared old AddIns registration"
} catch { Write-Output "no old AddIns registration" }
try { $xl.Quit() } catch {}
[void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
[GC]::Collect(); [GC]::WaitForPendingFinalizers()

Remove-Item (Join-Path $AddIns 'ExcelAgent.xlam') -Force -ErrorAction SilentlyContinue
Get-ChildItem $XLStart -Filter '~$ExcelAgent*' -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue

if (-not (Test-Path $XLStart)) { New-Item -ItemType Directory -Path $XLStart -Force | Out-Null }
Copy-Item $Build $Target -Force
Unblock-File $Target -ErrorAction SilentlyContinue

$len = (Get-Item $Target).Length
Write-Output "deployed -> $Target ($len bytes)"
Write-Output "Excel will load it on next start."
