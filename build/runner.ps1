param([string]$Xlam, [string]$Macro, [string]$LogPath)
# Opens the throwaway .xlam and runs one probe macro, passing the breadcrumb log
# path as its single argument. Records its own PID so the driver can clean up
# after a hang (Application.Run never returns while VBA is in break mode).
$ErrorActionPreference = 'Continue'
Set-Content -Path (Join-Path $env:TEMP 'runner.pid') -Value $PID

Write-Output ("xlam  : " + $Xlam)
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false; $x.AutomationSecurity = 1
try {
    $w = $x.Workbooks.Open($Xlam)
} catch {
    Write-Output ("RESULT: FAILED opening xlam: " + $_.Exception.Message)
    try { $x.Quit() } catch { }
    exit
}
$names = @()
foreach ($c in $w.VBProject.VBComponents) { $names += ($c.Name + '=' + $c.CodeModule.CountOfLines) }
Write-Output ("parts : " + ($names -join ','))
try {
    $r = $x.Run($Macro, $LogPath)
    Write-Output ("RESULT: OK " + $r)
} catch {
    $e = $_.Exception; $msg = $e.Message
    while ($e.InnerException) { $e = $e.InnerException; $msg += " | " + $e.Message }
    Write-Output ("RESULT: FAILED " + $msg)
    Write-Output ("HResult: 0x" + ('{0:X8}' -f $e.HResult))
}
try { $w.Close($false) } catch { }
try { $x.Quit() } catch { }
