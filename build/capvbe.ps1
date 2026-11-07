param([int]$ProcId, [string]$Out='vbe.png')
Add-Type -AssemblyName System.Drawing
Add-Type @'
using System; using System.Runtime.InteropServices;
public class CV {
  [StructLayout(LayoutKind.Sequential)] public struct RECT { public int L,T,R,B; }
  [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr h, IntPtr hdc, uint f);
  [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
}
'@
$p = Get-Process -Id $ProcId
$h = $p.MainWindowHandle
$r = New-Object CV+RECT
[CV]::GetWindowRect($h, [ref]$r) | Out-Null
$w = $r.R - $r.L; $ht = $r.B - $r.T
$bmp = New-Object System.Drawing.Bitmap $w, $ht
$g = [System.Drawing.Graphics]::FromImage($bmp); $hdc = $g.GetHdc()
[CV]::PrintWindow($h, $hdc, 2) | Out-Null
$g.ReleaseHdc($hdc); $g.Dispose()
$bmp.Save(('\\wsl.localhost\Ubuntu\home\eloaurdi\hay2010agent\excel-agent\build\' + $Out), [System.Drawing.Imaging.ImageFormat]::Png)
$bmp.Dispose(); Write-Output ("captured " + $w + "x" + $ht)
