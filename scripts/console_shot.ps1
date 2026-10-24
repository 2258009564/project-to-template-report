param(
  [Parameter(Mandatory=$true,ParameterSetName='Script')][string]$ScriptFile,
  [Parameter(Mandatory=$true,ParameterSetName='Command')][string]$Command,
  [Parameter(Mandatory=$true)][string]$Out,
  [string]$WorkingDirectory = (Get-Location).Path,
  [ValidateSet('pwsh','powershell')][string]$Shell = 'pwsh',
  [string]$Title = 'CONSOLE-CAPTURE',
  [int]$Left = 60,
  [int]$Top = 40,
  [int]$Width = 960,
  [int]$Height = 760,
  [int]$TimeoutSec = 30,
  [int]$SettleMs = 2000,
  [switch]$KeepOpen
)
# Capture a real console session as a PNG (Windows Terminal window client area).
# The caller supplies a demo script (preferred) or a command line; the terminal is
# launched with a dedicated title, sized, and captured through PrintWindow so an
# occluded window still yields a full image. Screen copy is the fallback.
# See references/console-screenshots.md.
$ErrorActionPreference='Stop'
Add-Type -AssemblyName System.Drawing
Add-Type @"
using System;
using System.Text;
using System.Collections.Generic;
using System.Runtime.InteropServices;

public class ConsoleShot {
  public delegate bool EnumProc(IntPtr hWnd, IntPtr lParam);
  [DllImport("user32.dll")] public static extern bool SetProcessDpiAwarenessContext(IntPtr ctx);
  [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc cb, IntPtr lParam);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetWindowText(IntPtr hWnd, StringBuilder sb, int max);
  [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr hWnd);
  [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
  [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int cmd);
  [DllImport("user32.dll")] public static extern bool MoveWindow(IntPtr hWnd, int x, int y, int w, int h, bool repaint);
  [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr hWnd, uint msg, IntPtr w, IntPtr l);
  [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr hWnd, out RECT r);
  [DllImport("user32.dll")] public static extern bool GetClientRect(IntPtr hWnd, out RECT r);
  [DllImport("user32.dll")] public static extern bool ClientToScreen(IntPtr hWnd, ref POINT p);
  [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr hWnd, IntPtr hdc, uint flags);

  [StructLayout(LayoutKind.Sequential)] public struct RECT { public int Left; public int Top; public int Right; public int Bottom; }
  [StructLayout(LayoutKind.Sequential)] public struct POINT { public int X; public int Y; }

  public static IntPtr Find(string part) {
    IntPtr found = IntPtr.Zero;
    EnumWindows(delegate(IntPtr h, IntPtr l) {
      if (!IsWindowVisible(h)) return true;
      var sb = new StringBuilder(512);
      GetWindowText(h, sb, sb.Capacity);
      if (sb.ToString().IndexOf(part, StringComparison.OrdinalIgnoreCase) >= 0) { found = h; return false; }
      return true;
    }, IntPtr.Zero);
    return found;
  }

  public static int[] WindowRect(IntPtr h) {
    RECT r;
    if (!GetWindowRect(h, out r)) return null;
    return new int[] { r.Left, r.Top, r.Right - r.Left, r.Bottom - r.Top };
  }

  public static int[] ClientRect(IntPtr h) {
    RECT cr; POINT p; p.X = 0; p.Y = 0;
    if (!GetClientRect(h, out cr)) return null;
    ClientToScreen(h, ref p);
    return new int[] { p.X, p.Y, cr.Right - cr.Left, cr.Bottom - cr.Top };
  }
}
"@
try { [ConsoleShot]::SetProcessDpiAwarenessContext([IntPtr](-4)) | Out-Null } catch {}

function Resolve-Shell([string]$name){
  if($name -eq 'pwsh'){
    $cmd = Get-Command pwsh.exe -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
    if($cmd){ return $cmd.Source }
    $candidate = Join-Path $env:ProgramFiles 'PowerShell\7\pwsh.exe'
    if(Test-Path -LiteralPath $candidate){ return $candidate }
    throw 'PowerShell 7 (pwsh.exe) was not found. Install it or pass -Shell powershell explicitly.'
  }
  return (Get-Command powershell.exe -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1).Source
}

function Resolve-Terminal(){
  $wt = Get-Command wt.exe -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
  if($wt){ return $wt.Source }
  $candidate = Join-Path $env:LOCALAPPDATA 'Microsoft\WindowsApps\wt.exe'
  if(Test-Path -LiteralPath $candidate){ return $candidate }
  throw 'Windows Terminal (wt.exe) was not found. It is required for console captures.'
}

function Test-Blank($bitmap){
  $colors = New-Object System.Collections.Generic.List[int]
  $w = $bitmap.Width; $h = $bitmap.Height
  foreach($pt in @(@(5,5),@([int]($w/2),[int]($h/2)),@([int]($w/2),20),@(20,[int]($h/2)),@([int]($w-6),[int]($h-6)))){
    [int]$px = $pt[0]; [int]$py = $pt[1]
    if($px -ge 0 -and $py -ge 0 -and $px -lt $w -and $py -lt $h){ $colors.Add($bitmap.GetPixel($px,$py).ToArgb()) }
  }
  return (($colors | Select-Object -Unique).Count -lt 2)
}

$shellPath = Resolve-Shell $Shell
$terminal  = Resolve-Terminal
$WorkingDirectory = (Resolve-Path -LiteralPath $WorkingDirectory).Path
$Out = [IO.Path]::GetFullPath($Out)
[IO.Directory]::CreateDirectory((Split-Path -Parent $Out)) | Out-Null

$quote = { param($s) "'" + ([string]$s).Replace("'","''") + "'" }
$lines = New-Object System.Collections.Generic.List[string]
$lines.Add('$Host.UI.RawUI.WindowTitle = ' + (& $quote $Title))
$lines.Add('Set-Location -LiteralPath ' + (& $quote $WorkingDirectory))
$lines.Add('Clear-Host')
if($PSCmdlet.ParameterSetName -eq 'Script'){
  $ScriptFile = (Resolve-Path -LiteralPath $ScriptFile).Path
  $lines.Add('& ' + (& $quote $ScriptFile))
}else{
  $lines.Add('Write-Host (' + (& $quote ('PS ' + $WorkingDirectory + '> ')) + ' + ' + (& $quote $Command) + ')')
  $lines.Add('Invoke-Expression ' + (& $quote $Command))
  $lines.Add('Write-Host (' + (& $quote ('PS ' + $WorkingDirectory + '> ')) + ')')
}
$lines.Add('while ($true) { Start-Sleep -Seconds 1 }')
$wrapper = Join-Path ([IO.Path]::GetTempPath()) ('console-shot-' + [Guid]::NewGuid().ToString('N') + '.ps1')
[IO.File]::WriteAllLines($wrapper, $lines, [Text.UTF8Encoding]::new($true))

$hwnd = [IntPtr]::Zero
try{
  $arguments = '-d "' + $WorkingDirectory + '" "' + $shellPath + '" -NoLogo -NoProfile -NoExit -File "' + $wrapper + '"'
  Start-Process -FilePath $terminal -ArgumentList $arguments

  $deadline = (Get-Date).AddSeconds($TimeoutSec)
  while((Get-Date) -lt $deadline){
    $hwnd = [ConsoleShot]::Find($Title)
    if($hwnd -ne [IntPtr]::Zero){ break }
    Start-Sleep -Milliseconds 300
  }
  if($hwnd -eq [IntPtr]::Zero){ throw "Terminal window '$Title' was not found within $TimeoutSec seconds." }

  [ConsoleShot]::ShowWindow($hwnd,9) | Out-Null
  Start-Sleep -Milliseconds 300
  [ConsoleShot]::SetForegroundWindow($hwnd) | Out-Null
  Start-Sleep -Milliseconds 400
  [ConsoleShot]::MoveWindow($hwnd,$Left,$Top,$Width,$Height,$true) | Out-Null
  Start-Sleep -Milliseconds $SettleMs

  $client = [ConsoleShot]::ClientRect($hwnd)
  $window = [ConsoleShot]::WindowRect($hwnd)
  if($null -eq $client -or $null -eq $window){ throw 'Window geometry could not be read.' }
  [int]$cw = $client[2]; [int]$ch = $client[3]
  if($cw -le 0 -or $ch -le 0){ throw "Invalid client area $cw x $ch." }

  # Primary path: PrintWindow renders the window even when another window covers it.
  $bmp = $null
  $printed = $false
  $winBmp = New-Object System.Drawing.Bitmap($window[2],$window[3])
  $g = [System.Drawing.Graphics]::FromImage($winBmp)
  $hdc = $g.GetHdc()
  try { $printed = [ConsoleShot]::PrintWindow($hwnd,$hdc,2) } finally { $g.ReleaseHdc($hdc); $g.Dispose() }
  if($printed){
    [int]$sx = $client[0] - $window[0]
    [int]$sy = $client[1] - $window[1]
    $crop = New-Object System.Drawing.Rectangle($sx,$sy,$cw,$ch)
    $bmp = $winBmp.Clone($crop,$winBmp.PixelFormat)
  }
  $winBmp.Dispose()

  # Fallback: screen copy needs the window unobstructed.
  if($null -eq $bmp -or (Test-Blank $bmp)){
    if($bmp){ $bmp.Dispose() }
    $bmp = New-Object System.Drawing.Bitmap($cw,$ch)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    try { $g.CopyFromScreen($client[0],$client[1],0,0,(New-Object System.Drawing.Size($cw,$ch))) } finally { $g.Dispose() }
  }

  $blank = Test-Blank $bmp
  $bmp.Save($Out,[System.Drawing.Imaging.ImageFormat]::Png)
  $bmp.Dispose()
  if($blank){ throw "Capture looks blank ($Out). Keep the terminal window visible and try again." }

  [ordered]@{ path=$Out; width=$cw; height=$ch; shell=$Shell; shell_path=$shellPath; terminal=$terminal; method=$(if($printed){'PrintWindow'}else{'ScreenCopy'}) }
}finally{
  Remove-Item -LiteralPath $wrapper -Force -ErrorAction SilentlyContinue
  if($hwnd -ne [IntPtr]::Zero -and -not $KeepOpen){
    [ConsoleShot]::PostMessage($hwnd,0x0010,[IntPtr]::Zero,[IntPtr]::Zero) | Out-Null
  }
}
