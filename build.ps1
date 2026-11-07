# ============================================================================
#  build.ps1 - assemble ExcelAgent.raw.xlam from the VBA sources.
#
#  Excel's object model is the only thing that can emit a real VBA project, so
#  this script drives Excel over COM: create a workbook, add the form and the
#  standard modules, save as .xlam. The ribbon is grafted on afterwards by
#  inject_ribbon.py (also in this folder).
#
#  Run from Windows:  powershell -NoProfile -ExecutionPolicy Bypass -File build.ps1
# ============================================================================
$ErrorActionPreference = 'Stop'

$Src     = '\\wsl.localhost\Ubuntu\home\eloaurdi\hay2010agent\excel-agent'
$VbaDir  = Join-Path $Src 'vba'
$OutDir  = Join-Path $Src 'build'
$TempOut = Join-Path $env:TEMP 'ExcelAgent.raw.xlam'

$Modules = @('modProto', 'modHttp', 'modWindow', 'modContext', 'modApply', 'modModels', 'Turn', 'modAgent')
$Classes = @('TurnResult')

. (Join-Path $PSScriptRoot 'build\compilegate.ps1')


# Excel rewrites its security settings from the values it cached at startup, so
# a running Excel can silently undo the VBA trust bit. Check before building.
$sec = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Security'
$vbom = (Get-ItemProperty $sec -Name AccessVBOM -ErrorAction SilentlyContinue).AccessVBOM
if ($vbom -ne 1) {
    Write-Host "AccessVBOM is not 1 - re-enabling (close Excel first if this sticks at 0)" -ForegroundColor Yellow
    if (-not (Test-Path $sec)) { New-Item -Path $sec -Force | Out-Null }
    Set-ItemProperty -Path $sec -Name AccessVBOM -Value 1 -Type DWord
    if ((Get-ItemProperty $sec -Name AccessVBOM).AccessVBOM -ne 1) {
        throw "Cannot enable 'Trust access to the VBA project object model'. Close Excel and retry."
    }
}
if (@(Get-Process EXCEL -ErrorAction SilentlyContinue).Count -gt 0) {
    Write-Host "warning: Excel is running; it may rewrite this setting when it exits" -ForegroundColor Yellow
}

Write-Host "== building $TempOut" -ForegroundColor Cyan
if (Test-Path $TempOut) { Remove-Item $TempOut -Force }
if (-not (Test-Path $OutDir)) { New-Item -ItemType Directory -Path $OutDir -Force | Out-Null }

$excel = $null
$wb = $null
$excelPid = 0
try {
    $beforePids = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object Id)
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible        = $false
    $excel.DisplayAlerts  = $false
    $excel.EnableEvents   = $false
    $excel.AutomationSecurity = 3   # msoAutomationSecurityForceDisable (no macros run)

    $excelPid = Get-NewExcelPid -Before $beforePids

    $wb = $excel.Workbooks.Add()
    $proj = $wb.VBProject

    # Class modules first: frmAgent and Turn are compiled against TurnResult.
    foreach ($c in $Classes) {
        $path = Join-Path $VbaDir "$c.cls"
        $code = Get-Content -Path $path -Raw -Encoding UTF8
        $comp = $proj.VBComponents.Add(2)      # vbext_ct_ClassModule
        $comp.Name = $c
        $comp.CodeModule.AddFromString($code)
        Write-Host "   + $c (class, $($code.Length) chars)"
    }

    # The form must exist before modAgent is compiled against it.
    $formCode = Get-Content -Path (Join-Path $VbaDir 'frmAgent.frm.txt') -Raw -Encoding UTF8
    $form = $proj.VBComponents.Add(3)          # vbext_ct_MSForm
    $form.Name = 'frmAgent'
    $form.CodeModule.AddFromString($formCode)
    Write-Host "   + frmAgent (UserForm, $($formCode.Length) chars)"

    foreach ($m in $Modules) {
        $path = Join-Path $VbaDir "$m.bas"
        $code = Get-Content -Path $path -Raw -Encoding UTF8
        $comp = $proj.VBComponents.Add(1)      # vbext_ct_StdModule
        $comp.Name = $m
        $comp.CodeModule.AddFromString($code)
        Write-Host "   + $m ($($code.Length) chars)"
    }

    # Force a real compile before saving. VBA compiles on demand, so a module
    # with a compile error otherwise loads fine and only blows up - as a modal
    # dialog Application.Run never returns from - the first time a real user
    # calls into it.
    Write-Host "== compiling VBA project"
    $gate = Test-VbaCompile -Excel $excel -ProcId $excelPid
    if (-not $gate.Ok) {
        throw ("VBA compile failed: " + $gate.Message + "   at " + $gate.Line)
    }
    Write-Host "   compiled clean" -ForegroundColor Green

    $wb.SaveAs($TempOut, 55)                   # 55 = xlOpenXMLAddIn (.xlam); 18 is the legacy .xla container
    Write-Host "   saved add-in" -ForegroundColor Green

    $copied = Join-Path $OutDir 'ExcelAgent.raw.xlam'
    Copy-Item -Path $TempOut -Destination $copied -Force
    Write-Host "== wrote $copied"
}
finally {
    if ($wb)    { try { $wb.Close($false) } catch {} }
    if ($excel) { try { $excel.Quit() } catch {} }
    if ($wb)    { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) }
    if ($excel) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
