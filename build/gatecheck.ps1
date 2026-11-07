# Validates build/compilegate.ps1 itself: a scratch project with a declaration
# placed after a procedure MUST be reported as a compile error. If this script
# says CLEAN, the gate is not doing its job and the build cannot trust it.
$ErrorActionPreference = 'Continue'
. '\\wsl.localhost\Ubuntu\home\eloaurdi\hay2010agent\excel-agent\build\compilegate.ps1'

$bad = @'
Option Explicit
Public Sub GoodOne()
End Sub
Private gDeclaredTooLate As Long
Public Sub Other()
End Sub
'@

$before = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object Id)
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false
$wb = $x.Workbooks.Add()
$c = $wb.VBProject.VBComponents.Add(1); $c.Name = 'modBad'
$c.CodeModule.AddFromString($bad)
$pid2 = Get-NewExcelPid -Before $before
Write-Output ("excel pid: " + $pid2)

$r = Test-VbaCompile -Excel $x -ProcId $pid2
if ($r.Ok) {
    Write-Output "GATE BROKEN: reported CLEAN for a project that cannot compile"
} else {
    Write-Output ("GATE OK: caught -> " + $r.Message)
    Write-Output ("           line -> " + $r.Line)
}
try { $wb.Close($false) } catch { }
try { $x.Quit() } catch { }
