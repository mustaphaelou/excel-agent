$p = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Security'
if (-not (Test-Path $p)) { New-Item -Path $p -Force | Out-Null }
Set-ItemProperty -Path $p -Name AccessVBOM -Value 1 -Type DWord
Write-Output ("AccessVBOM = " + (Get-ItemProperty $p -Name AccessVBOM).AccessVBOM)
$running = @(Get-Process EXCEL -ErrorAction SilentlyContinue)
Write-Output ("excel processes running: " + $running.Count)
