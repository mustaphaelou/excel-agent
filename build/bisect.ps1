param([string]$Modules = 'modProto,modHttp,modWindow,modContext,modApply,modAgent', [switch]$WithForm)
$ErrorActionPreference = 'Continue'
$Src = '\\wsl.localhost\Ubuntu\home\eloaurdi\hay2010agent\excel-agent'
$tmp = Join-Path $env:TEMP ("bisect-" + [guid]::NewGuid().ToString('N').Substring(0,6) + ".xlam")

$probe = @'
Public Function Ping() As String
    Ping = "pong"
End Function

Public Function FormProbe() As String
    Dim f As Object
    Set f = VBA.UserForms.Add("frmAgent")
    FormProbe = "controls=" & f.Controls.Count
    Unload f
End Function
'@

$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false
$wb = $x.Workbooks.Add()
$proj = $wb.VBProject
if ($WithForm) {
    $form = $proj.VBComponents.Add(3); $form.Name = 'frmAgent'
    $form.CodeModule.AddFromString((Get-Content (Join-Path $Src 'vba\frmAgent.frm.txt') -Raw -Encoding UTF8))
}
foreach ($m in $Modules.Split(',')) {
    if (-not $m) { continue }
    $c = $proj.VBComponents.Add(1); $c.Name = $m
    $c.CodeModule.AddFromString((Get-Content (Join-Path $Src ("vba\" + $m + ".bas")) -Raw -Encoding UTF8))
}
$p = $proj.VBComponents.Add(1); $p.Name = 'modProbe'
$p.CodeModule.AddFromString($probe)
$wb.SaveAs($tmp, 55)
$wb.Close($false)
$x.Quit()

Start-Sleep -Milliseconds 700
$x2 = New-Object -ComObject Excel.Application
$x2.Visible = $false; $x2.DisplayAlerts = $false; $x2.AutomationSecurity = 1
$w2 = $x2.Workbooks.Open($tmp)
$macro = 'Ping'
if ($WithForm) { $macro = 'FormProbe' }
try {
    $r = $x2.Run($macro)
    Write-Output ("COMPILE OK -> " + $macro + " = " + $r)
} catch {
    Write-Output ("FAILED: " + $_.Exception.Message)
}
$w2.Close($false); Start-Sleep -Milliseconds 300; $x2.Quit()
