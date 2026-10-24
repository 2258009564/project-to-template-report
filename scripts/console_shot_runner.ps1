# Internal runner. Keep this file beside console_shot.ps1.
param(
 [string]$ScriptFile,[string]$Command,
 [Parameter(Mandatory=$true)][string]$WorkingDirectory,
 [Parameter(Mandatory=$true)][string]$StatePath,
 [Parameter(Mandatory=$true)][string]$SessionId
)
$ErrorActionPreference='Stop'
$success=$false;$exitCode=1;$message='Execution did not finish normally.'
try {
 Set-Location -LiteralPath $WorkingDirectory
 $global:LASTEXITCODE=0
 if($ScriptFile){
  Write-Host ("PS {0}> & '{1}'" -f $WorkingDirectory,$ScriptFile.Replace("'","''"))
  & $ScriptFile
 }else{
  Write-Host ("PS {0}> {1}" -f $WorkingDirectory,$Command)
  & ([scriptblock]::Create($Command))
 }
 $invocationSucceeded=$?
 $exitCode=if($LASTEXITCODE -ne 0){[int]$LASTEXITCODE}elseif($invocationSucceeded){0}else{1}
 if($exitCode -ne 0){throw "Execution failed (exit code $exitCode)."}
 $success=$true;$message=''
}catch{
 if($exitCode -eq 0){$exitCode=1}
 $message=$_.Exception.Message
 Write-Host $message -ForegroundColor Red
}finally{
 $state=[ordered]@{session_id=$SessionId;success=$success;exit_code=$exitCode;error=$message;pid=$PID;shell_version=$PSVersionTable.PSVersion.ToString()}
 $temp=$StatePath+'.tmp'
 [IO.File]::WriteAllText($temp,($state|ConvertTo-Json -Depth 4),[Text.UTF8Encoding]::new($false))
 Move-Item -LiteralPath $temp -Destination $StatePath -Force
}
