param([Parameter(Mandatory=$true)][string]$WorkDir)
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'word_runtime.ps1')
$WorkDir=[IO.Path]::GetFullPath($WorkDir)
[IO.Directory]::CreateDirectory($WorkDir)|Out-Null
$cases=@(
 @{Runtime='Auto';CurrentMajor=5;PwshPath='pwsh';WindowsPowerShellPath='legacy';Name='PowerShell7';Executable='pwsh'},
 @{Runtime='Auto';CurrentMajor=5;PwshPath='';WindowsPowerShellPath='legacy';Name='WindowsPowerShell';Executable=$null},
 @{Runtime='Auto';CurrentMajor=7;PwshPath='pwsh';WindowsPowerShellPath='legacy';Name='PowerShell7';Executable=$null},
 @{Runtime='PowerShell7';CurrentMajor=5;PwshPath='pwsh';WindowsPowerShellPath='legacy';Name='PowerShell7';Executable='pwsh'},
 @{Runtime='PowerShell7';CurrentMajor=7;PwshPath='';WindowsPowerShellPath='legacy';Name='PowerShell7';Executable=$null},
 @{Runtime='WindowsPowerShell';CurrentMajor=7;PwshPath='pwsh';WindowsPowerShellPath='legacy';Name='WindowsPowerShell';Executable='legacy'},
 @{Runtime='WindowsPowerShell';CurrentMajor=5;PwshPath='pwsh';WindowsPowerShellPath='legacy';Name='WindowsPowerShell';Executable=$null}
)
foreach($case in $cases){
 $actual=Select-WordRuntime -Runtime $case.Runtime -CurrentMajor $case.CurrentMajor -PwshPath $case.PwshPath -WindowsPowerShellPath $case.WindowsPowerShellPath
 if($actual.Name -ne $case.Name -or $actual.Executable -ne $case.Executable){throw 'Runtime selection mismatch'}
}
foreach($runtime in @('PowerShell7','WindowsPowerShell')){
 $major=if($runtime -eq 'PowerShell7'){5}else{7}
 $rejected=$false
 try{Select-WordRuntime -Runtime $runtime -CurrentMajor $major -PwshPath '' -WindowsPowerShellPath ''|Out-Null}catch{$rejected=$_.Exception.Message -match 'explicitly requested'}
 if(-not $rejected){throw 'An explicitly missing runtime must fail'}
}
$nested=@{a=@{b=@{c=@{text='中文 深层数据'}}}}
$json=Join-Path $WorkDir 'nested [中文].json'
Write-WordJson -Value $nested -Path $json
$bytes=[IO.File]::ReadAllBytes($json)
if($bytes[0] -eq 239 -and $bytes[1] -eq 187 -and $bytes[2] -eq 191){throw 'JSON contains a BOM'}
$decoded=Get-Content -LiteralPath $json -Raw -Encoding UTF8|ConvertFrom-Json
if($decoded.a.b.c.text -cne '中文 深层数据'){throw 'Nested Unicode JSON was truncated'}

# Verify real process handoffs as well as pure selection, without starting Word.
$fixture=Join-Path $WorkDir 'runtime 参数 [测试].ps1'
$fixtureText=@'
param([string]$HelperPath,[string]$Out,[string]$Value,[switch]$Flag,[switch]$Fail,[string]$Runtime='Auto')
$ErrorActionPreference='Stop'
. $HelperPath
Invoke-WordRuntime -Runtime $Runtime -ScriptPath $PSCommandPath -BoundParameters $PSBoundParameters
Write-WordJson -Value @{major=$PSVersionTable.PSVersion.Major;value=$Value;flag=[bool]$Flag} -Path $Out
if($Fail){exit 23}
'@
[IO.File]::WriteAllText($fixture,$fixtureText,[Text.UTF8Encoding]::new($true))
$helper=Join-Path $PSScriptRoot 'word_runtime.ps1'
$legacy=Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
$pwsh=Get-Command pwsh.exe -CommandType Application -ErrorAction SilentlyContinue|Select-Object -First 1
$handoffs=0
if($pwsh -and (Test-Path -LiteralPath $legacy)){
 foreach($mode in @('Auto','WindowsPowerShell')){
  $exe=if($mode -eq 'Auto'){$legacy}else{$pwsh.Source}
  $expected=if($mode -eq 'Auto'){7}else{5}
  $out=Join-Path $WorkDir ($mode+' result [中文].json')
  $value='中文 path [一]; literal $()'
  & $exe -NoProfile -NonInteractive -File $fixture -HelperPath $helper -Out $out -Value $value -Flag -Runtime $mode
  if($LASTEXITCODE -ne 0){throw 'Runtime handoff failed'}
  $record=Get-Content -LiteralPath $out -Raw -Encoding UTF8|ConvertFrom-Json
  if($record.major -ne $expected -or $record.value -cne $value -or -not $record.flag){throw 'Handoff changed runtime, Unicode arguments or switch'}
  & $exe -NoProfile -NonInteractive -File $fixture -HelperPath $helper -Out $out -Value $value -Fail -Runtime $mode
  if($LASTEXITCODE -ne 23){throw 'Child exit code was lost'}
  $handoffs++
 }
}else{Write-Warning 'Both runtimes are needed for the process-handoff checks; those checks were skipped.'}
Write-Output ('TEST PASS: 7 selection cases, 2 unavailable-runtime errors, nested BOM-free JSON, {0} process handoff directions' -f $handoffs)
