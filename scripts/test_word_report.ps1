param(
 [Parameter(Mandatory=$true)][string]$WorkDir,
 [ValidateSet('Auto','PowerShell7','WindowsPowerShell')][string]$Runtime='Auto'
)
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'word_runtime.ps1')
Invoke-WordRuntime -Runtime $Runtime -ScriptPath $PSCommandPath -BoundParameters $PSBoundParameters
$WorkDir=[IO.Path]::GetFullPath($WorkDir)
[IO.Directory]::CreateDirectory($WorkDir)|Out-Null
$engine=Join-Path $PSScriptRoot 'word_report.ps1'
$source=Join-Path $WorkDir 'fixture.docx';$profile=Join-Path $WorkDir 'profile.json'
$app=$null;$doc=$null
try{
 $app=New-Object -ComObject Word.Application;$app.Visible=$false;$app.DisplayAlerts=0
 $doc=$app.Documents.Add();$doc.Content.Text="封面文本`r目录`r一级样例`r二级样例`r正文样例`r"
 $doc.Content.Font.NameFarEast='宋体';$doc.Content.Font.NameAscii='Times New Roman';$doc.Content.Font.NameOther='Times New Roman';$doc.Content.Font.Size=12
 $doc.Paragraphs.Item(1).Range.Font.Size=24
 $doc.Paragraphs.Item(3).Range.Font.Size=16;$doc.Paragraphs.Item(3).Range.Font.Bold=-1
 $doc.Paragraphs.Item(3).Format.OutlineLevel=1
 $doc.Paragraphs.Item(4).Range.Font.Size=14;$doc.Paragraphs.Item(4).Format.OutlineLevel=2
 $doc.Paragraphs.Item(5).Format.FirstLineIndent=24
 Save-WordDocx -Document $doc -Path $source
}finally{if($doc){$doc.Close(0)};if($app){$app.Quit()}}
$hash=(Get-FileHash $source).Hash
& $engine -Mode inspect -Source $source -Out $profile -Runtime $Runtime
$p=Get-Content $profile -Raw -Encoding UTF8|ConvertFrom-Json
if($p.paragraphs[2].font.Size -ne 16 -or $p.paragraphs[4].paragraph.FirstLineIndent -ne 24){throw 'Direct format extraction failed'}
$cache=& $engine -Mode inspect -Source $source -Out $profile -Runtime $Runtime
if($cache -notcontains 'CACHE HIT'){throw 'Cache reuse failed'}
# A profile cached by the old 5.1 writer must remain Python-readable on a cache hit.
[IO.File]::WriteAllText($profile,[IO.File]::ReadAllText($profile),[Text.UTF8Encoding]::new($true))
$legacyCache=& $engine -Mode inspect -Source $source -Out $profile -Runtime $Runtime
if($legacyCache -notcontains 'CACHE HIT' -or [IO.File]::ReadAllBytes($profile)[0] -eq 239){throw 'Legacy cache BOM normalization failed'}
$plan=@{
 replace_from='目录';roles=@{title=@{sample=2;outline=10};h1=@{sample=3;outline=1};body=@{sample=5;outline=10}}
 cover_cells=@();blocks=@(@{kind='toc';role='title';text='目录'},@{kind='text';role='h1';text='一、验收'},@{kind='text';role='body';text='中文直接格式保真验证。'})
 required=@('封面文本','中文直接格式保真验证');forbidden=@('未填占位符')
}
$path=Join-Path $WorkDir 'plan.json';Write-WordJson -Value $plan -Path $path
$output=Join-Path $WorkDir 'result.docx'
& $engine -Mode build -Source $source -Plan $path -Out $output -Runtime $Runtime
$qa=Get-Content ($output+'.qa.json') -Raw -Encoding UTF8|ConvertFrom-Json
if($qa.errors.Count -or $qa.toc -ne 1 -or $qa.checked_text_blocks -ne 2){throw 'Build verification failed'}
$plan.roles.body.font=@{Size=9999999}
Write-WordJson -Value $plan -Path $path
$rejected=$false
try{& $engine -Mode build -Source $source -Plan $path -Out (Join-Path $WorkDir 'must_not_exist.docx') -Runtime $Runtime}catch{$rejected=$_.Exception.Message -match 'mixed font'}
if(-not $rejected -or (Test-Path (Join-Path $WorkDir 'must_not_exist.docx'))){throw 'Mixed-format fail-closed check failed'}
if((Get-FileHash $source).Hash -ne $hash){throw 'Source mutation'}
Write-WordJson -Value @{passed=5;runtime=$PSVersionTable.PSVersion.ToString();checks=@('effective direct formatting','SHA256 cache reuse','real TOC and role verification','mixed-value rejection','source hash unchanged')} -Path (Join-Path $WorkDir 'test_result.json')
Write-Output 'TEST PASS: 5 behavioral checks'
