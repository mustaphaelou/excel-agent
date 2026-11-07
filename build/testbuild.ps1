$ErrorActionPreference='Continue'
$Src = '\\wsl.localhost\Ubuntu\home\eloaurdi\hay2010agent\excel-agent'
$tmp = Join-Path $env:TEMP 'ExcelAgent-test.xlam'
Copy-Item (Join-Path $Src 'build\ExcelAgent.xlam') $tmp -Force
Unblock-File $tmp -ErrorAction SilentlyContinue
$rep  = Join-Path $env:TEMP 'excelagent_selftest.txt'
$repf = Join-Path $env:TEMP 'excelagent_formtest.txt'
foreach ($f in @($rep,$repf)) { if (Test-Path $f) { Remove-Item $f -Force } }
$x = New-Object -ComObject Excel.Application
$x.Visible=$false; $x.DisplayAlerts=$false; $x.AutomationSecurity=1
$wb = $x.Workbooks.Open($tmp)
$null = $x.Run('SelfTest', $rep)
$null = $x.Run('FormSelfTest', $repf)
$wb.Close($false); Start-Sleep -Milliseconds 400; $x.Quit()
Write-Output '== self-test =='
if (Test-Path $rep) { Get-Content $rep }
Write-Output '== form =='
if (Test-Path $repf) { Get-Content $repf }
Copy-Item $rep (Join-Path $Src 'build\selftest.txt') -Force -ErrorAction SilentlyContinue
