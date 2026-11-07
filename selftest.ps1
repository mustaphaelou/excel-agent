# ============================================================================
#  selftest.ps1 - exercise the deployed add-in headlessly.
#
#  Excel launched over COM deliberately skips add-ins, so the file is opened
#  explicitly here. That still proves the thing that matters: the VBA project
#  compiles, the bridge answers, actions apply to real cells, and the chat
#  form builds and tears down.
# ============================================================================
$ErrorActionPreference = 'Continue'

$Src     = '\\wsl.localhost\Ubuntu\home\eloaurdi\hay2010agent\excel-agent'
$OutDir  = Join-Path $Src 'build'
$XLStart = Join-Path $env:APPDATA 'Microsoft\Excel\XLSTART'
$Target  = Join-Path $XLStart 'ExcelAgent.xlam'
$Rep     = Join-Path $env:TEMP 'excelagent_selftest.txt'
$RepForm = Join-Path $env:TEMP 'excelagent_formtest.txt'
$Trace   = Join-Path $env:TEMP 'excelagent_trace.log'

if (-not (Test-Path $Target)) { throw "Not deployed - run deploy.ps1 first ($Target)" }
foreach ($f in @($Rep, $RepForm, $Trace)) { if (Test-Path $f) { Remove-Item $f -Force } }

$xl = $null; $wb = $null
try {
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $false
    $xl.DisplayAlerts = $false
    $xl.AutomationSecurity = 1

    $wb = $xl.Workbooks.Open($Target)
    Write-Output "opened: $($wb.FullName)"

    $null = $xl.Run('SelfTest', $Rep)
    $null = $xl.Run('FormSelfTest', $RepForm)
}
catch { Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red }
finally {
    if ($wb) { try { $wb.Close($false) } catch {} }
    if ($xl) { try { $xl.Quit() } catch {} }
    if ($xl) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()

    foreach ($pair in @(@($Trace,'trace.log'), @($Rep,'selftest.txt'), @($RepForm,'formtest.txt'))) {
        if (Test-Path $pair[0]) { Copy-Item $pair[0] (Join-Path $OutDir $pair[1]) -Force -ErrorAction SilentlyContinue }
    }
    Write-Host "== self-test ==" -ForegroundColor Cyan
    if (Test-Path $Rep)     { Get-Content $Rep }     else { Write-Host '(none)' -ForegroundColor Yellow }
    Write-Host "== form test ==" -ForegroundColor Cyan
    if (Test-Path $RepForm) { Get-Content $RepForm } else { Write-Host '(none)' -ForegroundColor Yellow }
}
