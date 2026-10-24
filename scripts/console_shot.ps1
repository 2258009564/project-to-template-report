param(
 [Parameter(Mandatory=$true,ParameterSetName='Script')][string]$ScriptFile,
 [Parameter(Mandatory=$true,ParameterSetName='Command')][string]$Command,
 [Parameter(Mandatory=$true)][string]$Out,
 [string]$WorkingDirectory=(Get-Location).Path,
 [ValidateSet('Auto','pwsh','powershell')][string]$Shell='Auto',
 [string]$Title='CONSOLE-CAPTURE',
 [int]$Left=60,[int]$Top=40,
 [ValidateRange(320,8192)][int]$Width=960,
 [ValidateRange(200,8192)][int]$Height=760,
 [ValidateRange(1,3600)][int]$TimeoutSec=30,
 [ValidateRange(0,30000)][int]$SettleMs=2000,
 [switch]$KeepOpen
)
$ErrorActionPreference='Stop'
if($env:OS -ne 'Windows_NT'){throw 'Console capture requires Windows and Windows Terminal.'}
. (Join-Path $PSScriptRoot 'console_shot_helpers.ps1')
$shellPath=Resolve-ConsoleShell $Shell
$terminal=Get-Command wt.exe -CommandType Application -ErrorAction SilentlyContinue|Select-Object -First 1
if(-not $terminal){throw 'Windows Terminal (wt.exe) is required.'}
# Semicolons are WT command separators; only fixed, safe tokens cross that parser.
if($shellPath.Contains(';')){throw 'The shell executable path must not contain a Windows Terminal command separator.'}
$WorkingDirectory=(Resolve-Path -LiteralPath $WorkingDirectory).Path
if(-not (Test-Path -LiteralPath $WorkingDirectory -PathType Container)){throw 'WorkingDirectory must be a directory.'}
if($ScriptFile){$ScriptFile=(Resolve-Path -LiteralPath $ScriptFile).Path;if(-not (Test-Path -LiteralPath $ScriptFile -PathType Leaf)){throw 'ScriptFile must be a file.'}}
$Out=[IO.Path]::GetFullPath($Out)
if([IO.Path]::GetExtension($Out) -ine '.png'){throw 'Out must be a PNG path.'}
Add-Type -AssemblyName System.Drawing
Initialize-ConsoleCapture
$sessionId=[guid]::NewGuid().ToString('N')
$safeTitle=($Title -replace '[^\p{L}\p{N}_.-]','_')
if($safeTitle.Length -gt 48){$safeTitle=$safeTitle.Substring(0,48)}
$windowTitle=$safeTitle+'-'+$sessionId
$statePath=Join-Path ([IO.Path]::GetTempPath()) ('console-shot-'+$sessionId+'.json')
$runner=Join-Path $PSScriptRoot 'console_shot_runner.ps1'
$code='& '+(Quote-ConsoleLiteral $runner)+' -WorkingDirectory '+(Quote-ConsoleLiteral $WorkingDirectory)+' -StatePath '+(Quote-ConsoleLiteral $statePath)+' -SessionId '+(Quote-ConsoleLiteral $sessionId)
if($ScriptFile){$code+=' -ScriptFile '+(Quote-ConsoleLiteral $ScriptFile)}else{$code+=' -Command '+(Quote-ConsoleLiteral $Command)}
$encoded=[Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($code))
$arguments='-w new new-tab --title '+$windowTitle+' --suppressApplicationTitle "'+$shellPath+'" -NoLogo -NoProfile -NoExit -EncodedCommand '+$encoded
if($arguments.Length -gt 30000){throw 'Command is too long; put it in a file and use ScriptFile.'}
$hwnd=[IntPtr]::Zero;$oldDpi=[IntPtr]::Zero;$bmp=$null;$tempImage=$null
$previousWindow=[WordReportCapture.Native]::GetForegroundWindow()
$clock=[Diagnostics.Stopwatch]::StartNew()
try {
 # An explicitly requested terminal capture is a visible, foreground operation.
 Start-Process -FilePath $terminal.Source -ArgumentList $arguments|Out-Null
 $state=$null
 while($clock.Elapsed.TotalSeconds -lt $TimeoutSec){
  if(Test-Path -LiteralPath $statePath){
   $state=Get-Content -LiteralPath $statePath -Raw -Encoding UTF8|ConvertFrom-Json
   if($state.session_id -cne $sessionId){throw 'Capture session does not match its completion record.'}
   break
  }
  Start-Sleep -Milliseconds 100
 }
 if($null -eq $state){throw "The run did not finish within $TimeoutSec seconds; no screenshot was saved."}
 if(-not $state.success -or $state.exit_code -ne 0){throw "Run failed (exit code $($state.exit_code)): $($state.error)"}
 while($clock.Elapsed.TotalSeconds -lt $TimeoutSec){
  $hwnd=[WordReportCapture.Native]::Find($sessionId)
  if($hwnd -ne [IntPtr]::Zero){break}
  Start-Sleep -Milliseconds 100
 }
 if($hwnd -eq [IntPtr]::Zero){throw 'The completed run has no matching terminal window.'}
 $oldDpi=[WordReportCapture.Native]::SetThreadDpiAwarenessContext([IntPtr](-4))
 if($oldDpi -eq [IntPtr]::Zero){throw 'Per-monitor DPI awareness could not be enabled.'}
 [WordReportCapture.Native]::ShowWindow($hwnd,9)|Out-Null
 if(-not [WordReportCapture.Native]::MoveWindow($hwnd,$Left,$Top,$Width,$Height,$true)){throw 'Cannot size the terminal window.'}
 [WordReportCapture.Native]::SetForegroundWindow($hwnd)|Out-Null
 Start-Sleep -Milliseconds $SettleMs
 if([WordReportCapture.Native]::GetForegroundWindow() -ne $hwnd){throw 'The capture window could not be brought to the foreground.'}
 $rect=[WordReportCapture.Native]::ClientRect($hwnd)
 if(-not [WordReportCapture.Native]::Unobscured($hwnd,$rect)){throw 'The terminal is covered or outside the visible desktop; no screenshot was saved.'}
 $w=$rect[2];$h=$rect[3]
 $bmp=New-Object Drawing.Bitmap($w,$h,[Drawing.Imaging.PixelFormat]::Format32bppArgb)
 $graphics=[Drawing.Graphics]::FromImage($bmp)
 try{$graphics.CopyFromScreen($rect[0],$rect[1],0,0,(New-Object Drawing.Size($w,$h)))}finally{$graphics.Dispose()}
 $data=$bmp.LockBits((New-Object Drawing.Rectangle(0,0,$w,$h)),[Drawing.Imaging.ImageLockMode]::ReadOnly,[Drawing.Imaging.PixelFormat]::Format32bppArgb)
 try{
  $pixels=New-Object byte[] ($data.Stride*$h)
  [Runtime.InteropServices.Marshal]::Copy($data.Scan0,$pixels,0,$pixels.Length)
  $skipTop=[int][Math]::Ceiling(56*[WordReportCapture.Native]::GetDpiForWindow($hwnd)/96.0)
  $hasContent=[WordReportCapture.Native]::HasContent($pixels,$w,$h,$data.Stride,$skipTop)
 }finally{$bmp.UnlockBits($data)}
 if(-not $hasContent){throw 'The terminal body is blank; no screenshot was saved.'}
 $parent=Split-Path -Parent $Out
 [IO.Directory]::CreateDirectory($parent)|Out-Null
 $tempImage=Join-Path $parent ('.console-shot-'+$sessionId+'.png')
 $bmp.Save($tempImage,[Drawing.Imaging.ImageFormat]::Png)
 Move-Item -LiteralPath $tempImage -Destination $Out -Force
 $tempImage=$null
 [ordered]@{path=$Out;width=$w;height=$h;shell_version=$state.shell_version;shell_path=$shellPath;session_id=$sessionId;exit_code=$state.exit_code}
}finally{
 if($bmp){$bmp.Dispose()}
 # Find by our per-run token, including on timeout; never close a shared WT window.
 if($hwnd -eq [IntPtr]::Zero){$hwnd=[WordReportCapture.Native]::Find($sessionId)}
 if(-not $KeepOpen -and [WordReportCapture.Native]::Owns($hwnd,$sessionId)){
  $restore=([WordReportCapture.Native]::GetForegroundWindow() -eq $hwnd)
  [WordReportCapture.Native]::PostMessage($hwnd,0x0010,[IntPtr]::Zero,[IntPtr]::Zero)|Out-Null
  if($restore -and $previousWindow -ne [IntPtr]::Zero){[WordReportCapture.Native]::SetForegroundWindow($previousWindow)|Out-Null}
 }
 if($oldDpi -ne [IntPtr]::Zero){[WordReportCapture.Native]::SetThreadDpiAwarenessContext($oldDpi)|Out-Null}
 foreach($path in @($statePath,($statePath+'.tmp'),$tempImage)){
  if($path -and (Test-Path -LiteralPath $path)){Remove-Item -LiteralPath $path -Force -ErrorAction SilentlyContinue}
 }
}
