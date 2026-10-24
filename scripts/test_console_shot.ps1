param([Parameter(Mandatory=$true)][string]$WorkDir,[switch]$IncludeWindowCapture)
$ErrorActionPreference='Stop'
$WorkDir=[IO.Path]::GetFullPath($WorkDir)
[IO.Directory]::CreateDirectory($WorkDir)|Out-Null
. (Join-Path $PSScriptRoot 'console_shot_helpers.ps1')
Initialize-ConsoleCapture
$pixels=New-Object byte[] (100*100*4)
for($y=0;$y -lt 20;$y++){for($x=0;$x -lt 100;$x++){
 $i=($y*100+$x)*4;$pixels[$i]=255;$pixels[$i+1]=255;$pixels[$i+2]=255
}}
if([WordReportCapture.Native]::HasContent($pixels,100,100,400,40)){throw 'Tab-strip pixels were mistaken for terminal output'}
for($y=60;$y -lt 65;$y++){for($x=35;$x -lt 40;$x++){
 $i=($y*100+$x)*4;$pixels[$i]=255;$pixels[$i+1]=255;$pixels[$i+2]=255
}}
if(-not [WordReportCapture.Native]::HasContent($pixels,100,100,400,40)){throw 'Sparse terminal output was rejected as blank'}
$runner=Join-Path $PSScriptRoot 'console_shot_runner.ps1'
$caseDir=Join-Path $WorkDir "case's [中文]; test"
[IO.Directory]::CreateDirectory($caseDir)|Out-Null
$demo=Join-Path $caseDir "demo's [一]; test.ps1"
[IO.File]::WriteAllText($demo,"Write-Output 'SCRIPT_OK'",[Text.UTF8Encoding]::new($true))
$checked=0
foreach($shellName in @('pwsh','powershell')){
 try{$exe=Resolve-ConsoleShell $shellName}catch{Write-Warning "Skipping unavailable shell $shellName";continue}
 $nativeFailure='& '+(Quote-ConsoleLiteral $exe)+" -NoProfile -NonInteractive -Command 'exit 7'"
 foreach($mode in @('success','failure','native-failure','script')){
  $session=[guid]::NewGuid().ToString('N')
  $statePath=Join-Path $caseDir ($shellName+'-'+$mode+'.json')
  $payload='& '+(Quote-ConsoleLiteral $runner)+' -WorkingDirectory '+(Quote-ConsoleLiteral $caseDir)+' -StatePath '+(Quote-ConsoleLiteral $statePath)+' -SessionId '+(Quote-ConsoleLiteral $session)
  if($mode -eq 'script'){$payload+=' -ScriptFile '+(Quote-ConsoleLiteral $demo)}
  else{
   $command=if($mode -eq 'success'){"Start-Sleep -Milliseconds 300; Write-Output '中文 ; literal `$()'"}elseif($mode -eq 'failure'){"throw 'expected failure'"}else{$nativeFailure}
   $payload+=' -Command '+(Quote-ConsoleLiteral $command)
  }
  $encoded=[Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($payload))
  $clock=[Diagnostics.Stopwatch]::StartNew()
  & $exe -NoProfile -NonInteractive -EncodedCommand $encoded *>&1 | Out-File -LiteralPath (Join-Path $caseDir ($shellName+'-'+$mode+'.log')) -Encoding utf8
  if(-not (Test-Path -LiteralPath $statePath)){throw 'Runner exited without a completion record'}
  $state=Get-Content -LiteralPath $statePath -Raw -Encoding UTF8|ConvertFrom-Json
  if($state.session_id -cne $session){throw 'Wrong completion session'}
  $expected=if($mode -eq 'failure'){1}elseif($mode -eq 'native-failure'){7}else{0}
  if($state.exit_code -ne $expected -or [bool]$state.success -ne ($expected -eq 0)){throw "Incorrect completion status for $mode"}
  if($mode -eq 'success' -and $clock.Elapsed.TotalMilliseconds -lt 300){throw 'Completion preceded the command'}
  if([IO.File]::ReadAllBytes($statePath)[0] -eq 239){throw 'Completion JSON contains a BOM'}
  $checked++
 }
 if($IncludeWindowCapture){
  $out=Join-Path $WorkDir ($shellName+'-real.png')
  $capture=& (Join-Path $PSScriptRoot 'console_shot.ps1') -Command "Start-Sleep -Seconds 2; Write-Output 'CAPTURE_READY_OK'" -Shell $shellName -Out $out -TimeoutSec 30 -Width 960 -Height 560
  if(-not (Test-Path -LiteralPath $out) -or $capture.exit_code -ne 0){throw "Real capture failed for $shellName"}
 }
}
if($checked -eq 0){throw 'No shell runner was tested'}
Write-Output ('TEST PASS: content detection, {0} real background executions; visible capture requested: {1}' -f $checked,[bool]$IncludeWindowCapture)
