param([int]$Flag = 2, [string]$Xlam = '\\wsl.localhost\Ubuntu\home\eloaurdi\hay2010agent\excel-agent\build\ExcelAgent.xlam')
# Opens the add-in, shows the agent panel, and proves the panel is genuinely
# resizable: reads its window style bits, then drives a real Win32 resize and
# captures the panel before and after so the layout can be eyeballed.
Add-Type -AssemblyName System.Drawing
Add-Type @'
using System; using System.Text; using System.Collections.Generic; using System.Runtime.InteropServices;
public class FS {
  public delegate bool EnumProc(IntPtr h, IntPtr l);
  [StructLayout(LayoutKind.Sequential)] public struct RECT { public int L,T,R,B; }
  [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc cb, IntPtr l);
  [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
  [DllImport("user32.dll", CharSet=CharSet.Auto)] public static extern int GetWindowText(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll", CharSet=CharSet.Auto)] public static extern int GetClassName(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
  [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
  [DllImport("user32.dll")] public static extern bool MoveWindow(IntPtr h, int x, int y, int w, int ht, bool rep);
  [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr h, IntPtr hdc, uint f);
  [DllImport("user32.dll")] public static extern int GetWindowLongW(IntPtr h, int idx);
  public static IntPtr Find(uint target, string title) {
    IntPtr found = IntPtr.Zero;
    EnumWindows((h,l) => {
      uint pid; GetWindowThreadProcessId(h, out pid);
      if (pid == target) {
        var t = new StringBuilder(512); GetWindowText(h, t, 512);
        if (t.ToString() == title) { found = h; return false; }
      }
      return true;
    }, IntPtr.Zero);
    return found;
  }
  public static string Describe(IntPtr h) {
    var c = new StringBuilder(256); GetClassName(h, c, 256);
    RECT r; GetWindowRect(h, out r);
    int st = GetWindowLongW(h, -16);
    return "class=" + c + " rect=" + r.L + "," + r.T + " " + (r.R-r.L) + "x" + (r.B-r.T)
         + " visible=" + IsWindowVisible(h)
         + " WS_THICKFRAME=" + ((st & 0x40000) != 0)
         + " WS_MAXIMIZEBOX=" + ((st & 0x10000) != 0);
  }
  [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr p, EnumProc cb, IntPtr l);
  // MSForms controls are real child windows, so their rectangles are far more
  // trustworthy than a PrintWindow capture (which does not recurse reliably).
  public static System.Collections.Generic.List<string> Kids(IntPtr parent) {
    var res = new System.Collections.Generic.List<string>();
    RECT pr; GetWindowRect(parent, out pr);
    EnumChildWindows(parent, (h,l) => {
      RECT r; GetWindowRect(h, out r);
      int x = r.L - pr.L, y = r.T - pr.T, w = r.R - r.L, ht = r.B - r.T;
      if (w > 0 && ht > 0) {
        var c = new StringBuilder(256); GetClassName(h, c, 256);
        var t = new StringBuilder(512); GetWindowText(h, t, 512);
        string txt = t.ToString().Replace("\r", " ").Replace("\n", " ");
        if (txt.Length > 34) txt = txt.Substring(0, 34) + "..";
        res.Add(string.Format("  ({0,4},{1,4}) {2,4}x{3,-4} {4,-24} {5}", x, y, w, ht, c.ToString(), txt));
      }
      return true;
    }, IntPtr.Zero);
    return res;
  }
  public static uint Flag = 2;
  public static void Shot(IntPtr h, string path) {
    RECT r; GetWindowRect(h, out r);
    int w = r.R - r.L, ht = r.B - r.T;
    var bmp = new System.Drawing.Bitmap(w, ht);
    var g = System.Drawing.Graphics.FromImage(bmp);
    IntPtr hdc = g.GetHdc();
    PrintWindow(h, hdc, (uint)Flag);
    g.ReleaseHdc(hdc); g.Dispose();
    bmp.Save(path, System.Drawing.Imaging.ImageFormat.Png);
    bmp.Dispose();
  }
}
'@ -ReferencedAssemblies System.Drawing

$dir = '\\wsl.localhost\Ubuntu\home\eloaurdi\hay2010agent\excel-agent\build'
$before = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object Id)

$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false
$w = $x.Workbooks.Open($Xlam)
Start-Sleep -Milliseconds 1200

$now = @(Get-Process EXCEL -ErrorAction SilentlyContinue)
$newp = @($now | Where-Object { $before -notcontains $_.Id } | Select-Object -First 1)
$pid2 = if ($newp.Count) { $newp[0].Id } else { $now[0].Id }
Write-Output ("excel pid: " + $pid2)

$x.Run("AgentShow")
Start-Sleep -Milliseconds 2500

$h = [FS]::Find([uint32]$pid2, "Excel Agent")
if ($h -eq [IntPtr]::Zero) {
    Write-Output "panel not found while Excel hidden; making Excel visible"
    $x.Visible = $true
    Start-Sleep -Milliseconds 1200
    $h = [FS]::Find([uint32]$pid2, "Excel Agent")
}
if ($h -eq [IntPtr]::Zero) { Write-Output "RESULT: FAILED - no panel window"; $x.Quit(); exit 1 }

Write-Output ("before  : " + [FS]::Describe($h))
[FS]::Shot($h, (Join-Path $dir 'panel_small.png'))
Write-Output "before controls:"
[FS]::Kids($h) | ForEach-Object { Write-Output $_ }

# A real resize, the same thing dragging a corner does.
[FS]::MoveWindow($h, 60, 60, 880, 720, $true) | Out-Null
Start-Sleep -Milliseconds 2200      # the layout poll runs once a second
Write-Output ("after   : " + [FS]::Describe($h))
[FS]::Shot($h, (Join-Path $dir 'panel_large.png'))
Write-Output "after controls:"
[FS]::Kids($h) | ForEach-Object { Write-Output $_ }

# (AgentReset is deliberately not called: it raises a modal MsgBox, which would
# block this script forever. Nothing here depends on clearing the conversation.)
try { $w.Close($false) } catch { }
try { $x.Quit() } catch { }
Write-Output "RESULT: OK"
