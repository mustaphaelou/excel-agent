# ============================================================================
#  compilegate.ps1 - dot-source and call Test-VbaCompile.
#
#  VBA compiles on demand, so a module containing a compile error loads fine
#  and only fails the first time something calls into it - as a modal
#  "Compile error" dialog that Application.Run never returns from. That is
#  exactly how a broken add-in shipped once. Force the compile here instead,
#  detect the dialog, and fail the build loudly.
# ============================================================================

if (-not ([System.Management.Automation.PSTypeName]'VbaDlgProbe').Type) {
Add-Type @'
using System; using System.Text; using System.Collections.Generic; using System.Runtime.InteropServices;
public class VbaDlgProbe {
  public delegate bool EnumProc(IntPtr h, IntPtr l);
  [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc cb, IntPtr l);
  [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr p, EnumProc cb, IntPtr l);
  [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
  [DllImport("user32.dll", CharSet=CharSet.Auto)] public static extern int GetClassName(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll", CharSet=CharSet.Auto)] public static extern int GetWindowText(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll", CharSet=CharSet.Auto)] public static extern IntPtr SendMessage(IntPtr h, uint m, IntPtr w, IntPtr l);
  public static List<string> Dialogs(uint target) {
    var res = new List<string>();
    EnumWindows((h,l) => {
      uint pid; GetWindowThreadProcessId(h, out pid);
      if (pid == target) {
        var c = new StringBuilder(256); GetClassName(h, c, 256);
        if (c.ToString() == "#32770") {
          var parts = new List<string>();
          EnumChildWindows(h, (ch, cl) => {
            var t = new StringBuilder(1024); GetWindowText(ch, t, 1024);
            if (t.Length > 0) parts.Add(t.ToString());
            return true;
          }, IntPtr.Zero);
          res.Add(string.Join(" | ", parts));
        }
      }
      return true;
    }, IntPtr.Zero);
    return res;
  }
  public static void CloseAll(uint target) {
    EnumWindows((h,l) => {
      uint pid; GetWindowThreadProcessId(h, out pid);
      if (pid == target) {
        var c = new StringBuilder(256); GetClassName(h, c, 256);
        if (c.ToString() == "#32770") SendMessage(h, 0x0010, IntPtr.Zero, IntPtr.Zero);
      }
      return true;
    }, IntPtr.Zero);
  }
}
'@
}

# The PID of the hidden automation instance we just created.
function Get-NewExcelPid {
    param([int[]]$Before)
    $now = @(Get-Process EXCEL -ErrorAction SilentlyContinue)
    $new = @($now | Where-Object { $Before -notcontains $_.Id } | Select-Object -First 1)
    if ($new.Count -gt 0) { return $new[0].Id }
    if ($now.Count -gt 0) { return $now[0].Id }
    return 0
}

# Returns @{ Ok = $bool; Message = 'dialog text'; Line = 'module Lnnn: source' }
function Test-VbaCompile {
    param($Excel, [int]$ProcId)

    $ctl = $Excel.VBE.CommandBars.FindControl(1, 578, $null, $null)   # Debug > Compile VBAProject
    if (-not $ctl) { return @{ Ok = $false; Message = 'Compile command not found'; Line = '' } }
    $ctl.Execute()

    # The IDE raises its dialog asynchronously; poll briefly for it.
    $found = $null
    for ($i = 0; $i -lt 30; $i++) {
        $d = [VbaDlgProbe]::Dialogs([uint32]$ProcId)
        if ($d.Count -gt 0) { $found = $d; break }
        Start-Sleep -Milliseconds 100
    }

    # Where the IDE parked the cursor is the offending line, even if the
    # dialog has already been dismissed.
    $line = ''
    try {
        $cp = $Excel.VBE.ActiveCodePane
        $b = 0; $c = 0; $e = 0; $f = 0
        $cp.GetSelection([ref]$b, [ref]$c, [ref]$e, [ref]$f)
        $line = $cp.CodeModule.Parent.Name + " L$b C$c : " + $cp.CodeModule.Lines($b, 1).Trim()
    } catch { }

    if ($found) {
        [VbaDlgProbe]::CloseAll([uint32]$ProcId)
        Start-Sleep -Milliseconds 300
        return @{ Ok = $false; Message = ($found -join ' ;; '); Line = $line }
    }
    return @{ Ok = $true; Message = ''; Line = $line }
}
