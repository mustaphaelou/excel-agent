param([int]$ProcId, [string]$HideLike='', [string]$Out='')
# Lists every top-level window (and child Windows) owned by the Excel PID.
# With -HideLike, MDI code windows whose title matches are pushed off-screen so
# whatever is behind them can be captured.
Add-Type -AssemblyName System.Drawing
Add-Type @'
using System; using System.Text; using System.Collections.Generic; using System.Runtime.InteropServices;
public class WP {
  public delegate bool EnumProc(IntPtr h, IntPtr l);
  [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc cb, IntPtr l);
  [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr p, EnumProc cb, IntPtr l);
  [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
  [DllImport("user32.dll", CharSet=CharSet.Auto)] public static extern int GetClassName(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll", CharSet=CharSet.Auto)] public static extern int GetWindowText(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
  [DllImport("user32.dll")] public static extern bool MoveWindow(IntPtr h, int x, int y, int w, int ht, bool rep);
  [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr h, IntPtr hdc, uint f);
  [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
  [StructLayout(LayoutKind.Sequential)] public struct RECT { public int L,T,R,B; }
  public static List<IntPtr> Top(uint target) {
    var res = new List<IntPtr>();
    EnumWindows((h,l) => { uint pid; GetWindowThreadProcessId(h, out pid); if (pid==target) res.Add(h); return true; }, IntPtr.Zero);
    return res;
  }
  public static List<IntPtr> Kids(IntPtr parent) {
    var res = new List<IntPtr>();
    EnumChildWindows(parent, (h,l) => { res.Add(h); return true; }, IntPtr.Zero);
    return res;
  }
  public static string Txt(IntPtr h) { var t = new StringBuilder(512); GetWindowText(h, t, 512); return t.ToString(); }
  public static string Cls(IntPtr h) { var c = new StringBuilder(256); GetClassName(h, c, 256); return c.ToString(); }
  public static string Rect(IntPtr h) { RECT r; GetWindowRect(h, out r); return r.L + "," + r.T + " " + (r.R-r.L) + "x" + (r.B-r.T); }
}
'@
$wins = [WP]::Top([uint32]$ProcId)
foreach ($h in $wins) {
  $t = [WP]::Txt($h); $c = [WP]::Cls($h)
  if ($HideLike -and $t -like $HideLike) {
    [WP]::MoveWindow($h, -6000, -6000, 1200, 800, $true) | Out-Null
    Write-Output ("HID  [" + $c + "] " + $t + "  " + [WP]::Rect($h))
    continue
  }
  Write-Output ("TOP  [" + $c + "] '" + $t + "'  " + [WP]::Rect($h))
  foreach ($k in [WP]::Kids($h)) {
    $kt = [WP]::Txt($k); $kc = [WP]::Cls($k)
    if ($kt -or $HideLike) {
      if ($HideLike -and $kt -like $HideLike) { [WP]::MoveWindow($k, -6000, -6000, 1200, 800, $true) | Out-Null; Write-Output ("  HID [" + $kc + "] " + $kt) }
      elseif ($kt) { Write-Output ("     [" + $kc + "] " + $kt) }
    }
  }
}
if ($Out) {
  Start-Sleep -Milliseconds 700
  $p = Get-Process -Id $ProcId
  $h = $p.MainWindowHandle
  $r = New-Object WP+RECT
  [WP]::GetWindowRect($h, [ref]$r) | Out-Null
  $bmp = New-Object System.Drawing.Bitmap ($r.R-$r.L), ($r.B-$r.T)
  $g = [System.Drawing.Graphics]::FromImage($bmp); $hdc = $g.GetHdc()
  [WP]::PrintWindow($h, $hdc, 2) | Out-Null
  $g.ReleaseHdc($hdc); $g.Dispose()
  $bmp.Save(('\\wsl.localhost\Ubuntu\home\eloaurdi\hay2010agent\excel-agent\build\' + $Out), [System.Drawing.Imaging.ImageFormat]::Png)
  $bmp.Dispose(); Write-Output ("captured -> " + $Out)
}
