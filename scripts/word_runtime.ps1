# Shared by the Word engine and its tests; compatible with Windows PowerShell 5.1.
function Select-WordRuntime {
 param(
  [ValidateSet('Auto','PowerShell7','WindowsPowerShell')][string]$Runtime,
  [int]$CurrentMajor,
  [AllowNull()][AllowEmptyString()][string]$PwshPath,
  [AllowNull()][AllowEmptyString()][string]$WindowsPowerShellPath
 )
 if($CurrentMajor -lt 5){throw 'PowerShell 5.1 or PowerShell 7+ is required.'}
 if($Runtime -eq 'WindowsPowerShell'){
  if($CurrentMajor -eq 5){return @{Name='WindowsPowerShell';Executable=$null}}
  if(-not $WindowsPowerShellPath){throw 'Windows PowerShell 5.1 was explicitly requested but is unavailable.'}
  return @{Name='WindowsPowerShell';Executable=$WindowsPowerShellPath}
 }
 if($CurrentMajor -ge 7){return @{Name='PowerShell7';Executable=$null}}
 if($PwshPath){return @{Name='PowerShell7';Executable=$PwshPath}}
 if($Runtime -eq 'PowerShell7'){throw 'PowerShell 7 was explicitly requested but pwsh.exe was not found. Install it or use -Runtime WindowsPowerShell.'}
 return @{Name='WindowsPowerShell';Executable=$null}
}

function Invoke-WordRuntime {
 param([string]$Runtime,[string]$ScriptPath,[System.Collections.IDictionary]$BoundParameters)
 if($env:OS -ne 'Windows_NT'){throw 'Word COM automation requires Windows and Microsoft Word.'}
 $pwsh=Get-Command pwsh.exe -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
 $pwshPath=if($pwsh){$pwsh.Source}else{$null}
 if(-not $pwshPath -and $env:ProgramFiles){
  $candidate=Join-Path $env:ProgramFiles 'PowerShell\7\pwsh.exe'
  if(Test-Path -LiteralPath $candidate){$pwshPath=$candidate}
 }
 $legacy=Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
 if(-not (Test-Path -LiteralPath $legacy)){$legacy=$null}
 $choice=Select-WordRuntime -Runtime $Runtime -CurrentMajor $PSVersionTable.PSVersion.Major -PwshPath $pwshPath -WindowsPowerShellPath $legacy
 if($choice.Executable){
  # Pass arguments as data to -File; never interpolate them into executable code.
  $arguments=@('-NoLogo','-NoProfile','-NonInteractive','-STA','-File',$ScriptPath)
  foreach($key in $BoundParameters.Keys){
   if($key -eq 'Runtime'){continue}
   $value=$BoundParameters[$key]
   if($value -is [System.Management.Automation.SwitchParameter]){
    if($value.IsPresent){$arguments+='-'+$key}
   }else{$arguments+=@("-$key",[string]$value)}
  }
  $arguments+=@('-Runtime',$choice.Name)
  & $choice.Executable @arguments
  exit $LASTEXITCODE
 }
 if($choice.Name -eq 'WindowsPowerShell'){
  Write-Warning 'Using Windows PowerShell 5.1 compatibility mode. PowerShell 7+ is recommended.'
 }
 Write-Host ('RUNTIME PowerShell {0} ({1})' -f $PSVersionTable.PSVersion,$choice.Name)
}

function Write-WordJson {
 param([object]$Value,[string]$Path,[int]$Depth=30)
 $fullPath=[IO.Path]::GetFullPath($Path)
 $parent=Split-Path -Parent $fullPath
 if($parent){[IO.Directory]::CreateDirectory($parent)|Out-Null}
 $json=ConvertTo-Json -InputObject $Value -Depth $Depth
 [IO.File]::WriteAllText($fullPath,$json,[Text.UTF8Encoding]::new($false))
}

function Save-WordDocx {
 param([object]$Document,[string]$Path)
 # SaveAs2 declares by-reference object arguments. Typed references avoid the
 # Windows PowerShell 5.1 COM binder's problematic implicit conversion path.
 [int]$format=16
 $Document.SaveAs2([ref]$Path,[ref]$format)
}
