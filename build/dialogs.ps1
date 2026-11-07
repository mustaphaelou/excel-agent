Add-Type @'
using System; using System.Text; using System.Collections.Generic; using System.Runtime.InteropServices;
public class Dlg {
  public delegate bool EnumProc(IntPtr h, IntPtr l);
  [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc cb, IntPtr l);
  [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr p, EnumProc cb, IntPtr l);
  [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
  [DllImport("user32.dll", CharSet=CharSet.Auto)] public static extern int GetClassName(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll", CharSet=CharSet.Auto)] public static extern int GetWindowText(IntPtr h, StringBuilder s, int n);
  public static List<string> Find(uint target) {
    var res = new List<string>();
    EnumWindows((h,l) => {
      uint pid; GetWindowThreadProcessId(h, out pid);
      if (pid == target) {
        var c = new StringBuilder(256); GetClassName(h, c, 256);
        if (c.ToString() == "#32770") {
          var t = new StringBuilder(512); GetWindowText(h, t, 512);
          res.Add("DIALOG: " + t);
          EnumChildWindows(h, (ch, cl) => {
            var cc = new StringBuilder(256); GetClassName(ch, cc, 256);
            var ct = new StringBuilder(512); GetWindowText(ch, ct, 512);
            res.Add("   child " + cc + " :: " + ct);
            return true;
          }, IntPtr.Zero);
        }
      }
      return true;
    }, IntPtr.Zero);
    return res;
  }
}
'@
foreach ($p in Get-Process EXCEL -ErrorAction SilentlyContinue) {
  Write-Output ("pid " + $p.Id + " :: " + $p.MainWindowTitle)
  [Dlg]::Find([uint32]$p.Id) | ForEach-Object { Write-Output ("  " + $_) }
}
