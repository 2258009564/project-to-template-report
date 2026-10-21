param(
 [Parameter(Mandatory=$true)][ValidateSet('inspect','build')][string]$Mode,
 [Parameter(Mandatory=$true)][string]$Source,
 [Parameter(Mandatory=$true)][string]$Out,
 [string]$Plan,[switch]$Force
)
$ErrorActionPreference='Stop'
$schema=2
$Source=(Resolve-Path -LiteralPath $Source).Path
$Out=[IO.Path]::GetFullPath($Out)
if($Source -eq $Out){throw 'Output must differ from source'}
$hash=(Get-FileHash -LiteralPath $Source -Algorithm SHA256).Hash
$fontNames=@('NameFarEast','NameAscii','NameOther','Size','Bold','Italic','Color','Underline','Scaling','Spacing')
$paraNames=@('Alignment','SpaceBefore','SpaceAfter','LineSpacingRule','LineSpacing','LeftIndent','RightIndent','FirstLineIndent','KeepWithNext','KeepTogether','WidowControl','PageBreakBefore','OutlineLevel')
function Props($obj,$names){$h=[ordered]@{};foreach($n in $names){$h[$n]=$obj.$n};return $h}
function Geometry($doc){$a=@();foreach($s in $doc.Sections){$a+=,(Props $s.PageSetup @('PageWidth','PageHeight','TopMargin','BottomMargin','LeftMargin','RightMargin','HeaderDistance','FooterDistance','Orientation','DifferentFirstPageHeaderFooter'))};return ,$a}
function JsonSave($obj,$path){$dir=Split-Path -Parent $path;if($dir){[IO.Directory]::CreateDirectory($dir)|Out-Null};$obj|ConvertTo-Json -Depth 30|Set-Content -LiteralPath $path -Encoding utf8}
function IsMixed($v){return ($v -is [ValueType] -and [double]$v -eq 9999999)}
function SetWordProperty($obj,$key,$v){
 # Explicit conversion avoids PowerShell COM setter caching numeric types across dynamic member names.
 if($key -in @('Size','Spacing','SpaceBefore','SpaceAfter','LineSpacing','LeftIndent','RightIndent','FirstLineIndent')){$v=[float]$v}
 elseif($v -is [ValueType]){$v=[int]$v}
 $obj.GetType().InvokeMember($key,[Reflection.BindingFlags]::SetProperty,$null,$obj,@($v))|Out-Null
}
if($Mode -eq 'inspect' -and (Test-Path -LiteralPath $Out) -and -not $Force){
 $old=Get-Content -LiteralPath $Out -Raw -Encoding UTF8|ConvertFrom-Json
 if($old.schema -eq $schema -and $old.sha256 -eq $hash){Write-Output 'CACHE HIT';exit 0}
}
$app=$null;$doc=$null;$target=$null;$tmp=$null
try{
 $app=New-Object -ComObject Word.Application
 $app.Visible=$false;$app.DisplayAlerts=0;$app.AutomationSecurity=3;$app.ScreenUpdating=$false
 $doc=$app.Documents.Open($Source,$false,$true,$false)
 $geometry=Geometry $doc
 if($Mode -eq 'inspect'){
  $rows=@();$i=0
  foreach($p in $doc.Paragraphs){
   $i++;$text=$p.Range.Text.Trim([char]13,[char]7)
   $row=[ordered]@{index=$i;text=$text;style=$p.Style.NameLocal;start=$p.Range.Start;end=$p.Range.End;in_table=$p.Range.Information(12);font=(Props $p.Range.Font $fontNames);paragraph=(Props $p.Format $paraNames)}
   if(@($row.font.Values|Where-Object {IsMixed $_}).Count -gt 0){
    $runs=@();$last='';$count=0
    foreach($ch in $p.Range.Characters){$f=Props $ch.Font $fontNames;$key=$f|ConvertTo-Json -Compress;if($key -ne $last){$runs+=,[ordered]@{text=$ch.Text;font=$f};$last=$key;$count++};if($count -ge 8){break}}
    $row['run_samples']=$runs
   }
   $rows+=,$row
  }
  $tables=@();foreach($t in $doc.Tables){$tables+=,[ordered]@{rows=$t.Rows.Count;columns=$t.Columns.Count;start=$t.Range.Start;end=$t.Range.End}}
  $fields=@();foreach($f in $doc.Fields){$fields+=$f.Code.Text.Trim()}
  JsonSave ([ordered]@{schema=$schema;source=$Source;sha256=$hash;sections=$geometry;pages=$doc.ComputeStatistics(2);paragraphs=$rows;tables=$tables;fields=$fields;inline_images=$doc.InlineShapes.Count;floating_shapes=$doc.Shapes.Count}) $Out
  Write-Output ('INSPECT paragraphs={0} sections={1}' -f $rows.Count,$doc.Sections.Count)
 }else{
  if(-not $Plan){throw 'Build requires -Plan'}
  $planData=Get-Content -LiteralPath $Plan -Raw -Encoding UTF8|ConvertFrom-Json
  $samples=@{};$count=0;foreach($p in $doc.Paragraphs){$count++;$samples[$count]=$p}
  $roles=@{}
  foreach($entry in $planData.roles.PSObject.Properties){
   $r=$entry.Value;if(-not $samples.ContainsKey([int]$r.sample)){throw "Invalid role sample $($entry.Name)"}
   $p=$samples[[int]$r.sample];$font=Props $p.Range.Font $fontNames;$pf=Props $p.Format $paraNames
   foreach($v in $r.font.PSObject.Properties){$font[$v.Name]=$v.Value}
   foreach($v in $r.paragraph.PSObject.Properties){$pf[$v.Name]=$v.Value}
   if($null -ne $r.outline){$pf.OutlineLevel=[int]$r.outline}
   foreach($key in @($font.Keys)){if((IsMixed $font[$key]) -or ($key -like 'Name*' -and -not $font[$key])){throw "Unresolved mixed font $($entry.Name).$key"}}
   foreach($key in @($pf.Keys)){if(IsMixed $pf[$key]){throw "Unresolved mixed paragraph $($entry.Name).$key"}}
   $roles[$entry.Name]=@{font=$font;paragraph=$pf}
  }
  $boundaries=@();foreach($p in $doc.Paragraphs){if($p.Range.Text.Trim([char]13,[char]7) -ceq $planData.replace_from){$boundaries+=$p.Range.Start}}
  if($boundaries.Count -ne 1){throw "Boundary expected 1, found $($boundaries.Count)"}
  $start=[int]$boundaries[0]
  [IO.Directory]::CreateDirectory((Split-Path -Parent $Out))|Out-Null
  $tmp=Join-Path (Split-Path -Parent $Out) ([IO.Path]::GetFileNameWithoutExtension($Out)+'.'+[Guid]::NewGuid().ToString('N')+'.building.docx')
  # SaveAs on the read-only source opens a new output without changing the original file.
  $doc.SaveAs2($tmp,16);$target=$doc;$doc=$null
  foreach($edit in $planData.cover_cells){
   $r=$target.Tables.Item([int]$edit.table).Cell([int]$edit.row,[int]$edit.column).Range.Duplicate
   $r.End=$r.End-1;$oldFont=$r.Font.Duplicate;$oldPF=$r.ParagraphFormat.Duplicate
   $r.Text=[string]$edit.text;$r.Font=$oldFont;$r.ParagraphFormat=$oldPF
  }
  # Cover edits may shift character offsets; re-resolve the exact boundary.
  $found=@();foreach($p in $target.Paragraphs){if($p.Range.Text.Trim([char]13,[char]7) -ceq $planData.replace_from){$found+=$p.Range.Start}}
  if($found.Count -ne 1){throw 'Boundary changed unexpectedly'};$start=[int]$found[0]
  $blocks=@($planData.blocks);$texts=@();foreach($b in $blocks){
   if(-not $roles.ContainsKey([string]$b.role)){throw "Unknown role $($b.role)"}
   if($b.kind -notin @('text','toc','image','table')){throw "Unsupported block $($b.kind)"}
   if($b.kind -in @('image','table')){$texts+=' '}
   elseif($b.kind -eq 'toc'){$texts+=([string]$b.text+"`r ")}
   else{$texts+=[string]$b.text}
  }
  $rng=$target.Range($start,$target.Content.End-1);$rng.Text=($texts -join "`r")+"`r"
  Write-Output ('INSERTED blocks='+$blocks.Count)
  # Materialize each role once. Reapplying 25 style properties for every paragraph is unnecessary COM work.
  foreach($roleName in $roles.Keys){
   $role=$roles[$roleName];$styleName='WR_'+$roleName
   try{$st=$target.Styles.Item($styleName)}catch{$st=$target.Styles.Add($styleName,1)}
   foreach($key in $role.font.Keys){try{SetWordProperty $st.Font $key $role.font[$key]}catch{throw "Role $styleName Font.$key=$($role.font[$key]): $_"}}
   foreach($key in $role.paragraph.Keys){try{SetWordProperty $st.ParagraphFormat $key $role.paragraph[$key]}catch{throw "Role $styleName Paragraph.$key=$($role.paragraph[$key]): $_"}}
  }
  $ranges=@();$pos=$start
  for($i=0;$i -lt $blocks.Count;$i++){
   $end=$pos+$texts[$i].Length+1;$r=$target.Range($pos,$end);$name='WRB'+$i
   $target.Bookmarks.Add($name,$r)|Out-Null;$ranges+=,$name;$pos=$end
   $styleName='WR_'+[string]$blocks[$i].role
   $r.Style=$target.Styles.Item($styleName)
   # Remove inherited hint overrides only inside newly generated blocks.
   $r.Font.Reset();$r.ParagraphFormat.Reset();$r.HighlightColorIndex=0
   if($blocks[$i].page_break_before){$r.ParagraphFormat.PageBreakBefore=-1}
  }
  Write-Output 'ROLES APPLIED'
  # Objects are inserted backwards so bookmark positions remain unambiguous.
  for($i=$blocks.Count-1;$i -ge 0;$i--){
   $b=$blocks[$i];$r=$target.Bookmarks.Item($ranges[$i]).Range.Duplicate
   if($b.kind -eq 'image'){
    $r.End=$r.End-1;$r.Text='';$pic=$target.InlineShapes.AddPicture([string]$b.path,$false,$true,$r)
    $pic.LockAspectRatio=-1;$pic.Width=[double]$b.width_cm*28.3464567
   }elseif($b.kind -eq 'table'){
    $rs=@($b.rows);$cols=@($rs[0]).Count;if($cols -lt 1){throw 'Empty table'}
    foreach($row in $rs){if(@($row).Count -ne $cols){throw 'Ragged table'}}
    $r.End=$r.End-1;$r.Text='';$tb=$target.Tables.Add($r,$rs.Count,$cols)
    $tb.AllowAutoFit=$false;$tb.PreferredWidthType=2;$tb.PreferredWidth=100
    $tb.Borders.Enable=1;$tb.Shading.BackgroundPatternColor=-16777216
    $tb.TopPadding=4;$tb.BottomPadding=4;$tb.LeftPadding=5;$tb.RightPadding=5
    $tb.Rows.Item(1).HeadingFormat=-1
    for($ri=0;$ri -lt $rs.Count;$ri++){for($ci=0;$ci -lt $cols;$ci++){
     $cell=$tb.Cell($ri+1,$ci+1);$cell.Range.Text=[string]$rs[$ri][$ci];$cell.VerticalAlignment=1
     $cell.Range.Style=$target.Styles.Item('WR_'+[string]$b.role)
    }}
    $target.Bookmarks.Add($ranges[$i],$tb.Range)|Out-Null
   }elseif($b.kind -eq 'toc'){
    # Insert strictly inside the reserved TOC block, never at the next heading's bookmark boundary.
    $r.Start=$r.Start+([string]$b.text).Length+1;$r.End=$r.Start+1;$r.Text='';$r.Collapse(1)
    $toc=$target.TablesOfContents.Add($r,$true,1,2,$false,'',$true,$true,'',$true,$true,$true)
   }
  }
  Write-Output 'OBJECTS INSERTED'
  $target.Repaginate();foreach($toc in $target.TablesOfContents){$toc.Update()}
  $null=$target.Fields.Update();$target.Repaginate()
  # Word TOC styles are independent from report headings, so never rewrite them as Heading styles.
  foreach($toc in $target.TablesOfContents){$toc.Range.Font.NameFarEast='宋体';$toc.Range.Font.Size=14;$toc.Range.Font.Color=0;$toc.Range.ParagraphFormat.SpaceBefore=0;$toc.Range.ParagraphFormat.SpaceAfter=3;$toc.Range.ParagraphFormat.LineSpacingRule=0}
  $target.Repaginate();foreach($toc in $target.TablesOfContents){$toc.UpdatePageNumbers()}
  $errors=@();$all=$target.Content.Text
  foreach($t in $planData.required){if(-not $all.Contains([string]$t)){$errors+="Missing: $t"}}
  foreach($t in $planData.forbidden){if($all.Contains([string]$t)){$errors+="Forbidden: $t"}}
  if(($blocks|Where-Object kind -eq 'toc') -and $target.TablesOfContents.Count -ne 1){$errors+='Expected one real TOC'}
  if((Get-FileHash -LiteralPath $Source -Algorithm SHA256).Hash -ne $hash){$errors+='Source changed'}
  if((ConvertTo-Json -InputObject (Geometry $target) -Compress) -ne (ConvertTo-Json -InputObject $geometry -Compress)){$errors+='Section geometry changed'}
  $checked=0;$locations=@()
  for($i=0;$i -lt $blocks.Count;$i++){
   $b=$blocks[$i];$r=$target.Bookmarks.Item($ranges[$i]).Range
   $locations+=,[ordered]@{block=$i;role=$b.role;kind=$b.kind;page=$r.Information(3);text=([string]$b.text).Substring(0,[Math]::Min(60,([string]$b.text).Length))}
   if($b.kind -ne 'text'){continue};$role=$roles[[string]$b.role]
   foreach($key in $role.font.Keys){$actual=$r.Font.$key;$expected=$role.font[$key];if($actual -is [string]){$bad=$actual -ne $expected}else{$bad=[Math]::Abs([double]$actual-[double]$expected) -gt 0.6};if($bad){$errors+="Block $i font.$key expected=$expected actual=$actual"}}
   foreach($key in $role.paragraph.Keys){if($key -eq 'PageBreakBefore' -and $b.page_break_before){continue};$actual=$r.ParagraphFormat.$key;$expected=$role.paragraph[$key];if([Math]::Abs([double]$actual-[double]$expected) -gt 0.6){$errors+="Block $i paragraph.$key expected=$expected actual=$actual"}}
   $checked++
  }
  $qa=[ordered]@{source_sha256=$hash;pages=$target.ComputeStatistics(2);sections=$target.Sections.Count;toc=$target.TablesOfContents.Count;checked_text_blocks=$checked;errors=$errors;locations=$locations}
  JsonSave $qa ($Out+'.qa.json')
  if($errors.Count){throw ('Verification: '+($errors -join '; '))}
  $target.Save();if($planData.pdf){$target.ExportAsFixedFormat([string]$planData.pdf,17)}
  $target.Close(0);$target=$null
  Move-Item -LiteralPath $tmp -Destination $Out -Force;$tmp=$null
  Write-Output ('BUILD PASS pages={0} checked={1} errors=0' -f $qa.pages,$checked)
 }
}catch{Write-Output $_.ScriptStackTrace;throw}finally{if($target){$target.Close(0)};if($doc){$doc.Close(0)};if($app){$app.Quit()}}
