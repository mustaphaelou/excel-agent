param(
  [string]$Modules = 'modProto,modHttp,modWindow,modContext,modApply,modAgent',
  [switch]$WithForm,
  [string]$Out = ''
)
# Builds a throwaway .xlam from a subset of the real modules plus modProbe.
$ErrorActionPreference = 'Continue'
$Src = '\\wsl.localhost\Ubuntu\home\eloaurdi\hay2010agent\excel-agent'
if (-not $Out) { $Out = Join-Path $env:TEMP ("probe-" + [guid]::NewGuid().ToString('N').Substring(0,6) + ".xlam") }

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
$p.CodeModule.AddFromString((Get-Content (Join-Path $Src 'build\probe.bas') -Raw -Encoding UTF8))

$wb.SaveAs($Out, 55)
$wb.Close($false)
$x.Quit()
Write-Output $Out
