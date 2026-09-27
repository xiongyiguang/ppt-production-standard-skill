param([Parameter(Mandatory=$true)][string]$InputPptx,
      [Parameter(Mandatory=$true)][string]$OutputDir)
$ErrorActionPreference='Stop'
$inputPath=(Resolve-Path -LiteralPath $InputPptx).Path
$outPath=[IO.Path]::GetFullPath($OutputDir)
if(Test-Path -LiteralPath $outPath){throw 'Use a new verification directory.'}
[IO.Directory]::CreateDirectory($outPath)|Out-Null
$app=$null;$deck=$null
$findings=[Collections.Generic.List[object]]::new()
$measurements=@{};$operations=[Collections.Generic.List[object]]::new()
try {
 $app=New-Object -ComObject PowerPoint.Application
 foreach($openDeck in $app.Presentations){if($openDeck.FullName -eq $inputPath){throw 'Input is already open. Use a saved working copy.'}}
 $deck=$app.Presentations.Open($inputPath,$true,$false,$false)
 $count=$deck.Slides.Count
 for($i=1;$i -le $count;$i++){
  $slide=$deck.Slides.Item($i)
  $slide.Export((Join-Path $outPath "slide-$i.png"),'PNG',1600,900)
  $candidates=[Collections.Generic.List[int]]::new();$position=0
  foreach($shape in $slide.Shapes){
   $position++
   if($shape.Name.StartsWith('library:') -and $shape.Type -eq 1){$candidates.Add($position)}
   if($shape.HasTextFrame -eq -1 -and $shape.TextFrame.HasText -eq -1){
    $tr=$shape.TextFrame.TextRange;$pars=@()
    for($j=1;$j -le $tr.Paragraphs().Count;$j++){
     $para=$tr.Paragraphs($j,1)
     if([string]::IsNullOrWhiteSpace($para.Text)){continue}
     $fmt=$para.ParagraphFormat
     $pars+=@{multiple=($fmt.LineRuleWithin -eq -1);within=[double]$fmt.SpaceWithin;before=[double]$fmt.SpaceBefore;after=[double]$fmt.SpaceAfter}
    }
    $measurements["${i}:$($shape.Id)"]=@{lines=[int]$tr.Lines().Count;paragraphs=$pars}
    $height=[double]$shape.Height-$shape.TextFrame.MarginTop-$shape.TextFrame.MarginBottom
    $width=[double]$shape.Width-$shape.TextFrame.MarginLeft-$shape.TextFrame.MarginRight
    if($tr.BoundHeight -gt $height+2 -or $tr.BoundWidth -gt $width+2){
     $findings.Add(@{slide=$i;id=$shape.Id;name=$shape.Name;kind='text-overflow';text=$tr.Text;available=@($width,$height);actual=@($tr.BoundWidth,$tr.BoundHeight)})
    }
   }
  }
  # Exercise native objects on a disposable duplicate, then discard without saving.
  if($candidates.Count -ge 2){
   $duplicate=$slide.Duplicate().Item(1)
   try {
    $duplicate.Shapes.Item($candidates[0]).Name='qa-native-one'
    $duplicate.Shapes.Item($candidates[1]).Name='qa-native-two'
    $range=$duplicate.Shapes.Range(@('qa-native-one','qa-native-two'));$group=$range.Group()
    $group.IncrementLeft(1);$group.IncrementLeft(-1)
    $ungroup=$group.Ungroup();$shape=$ungroup.Item(1)
    if($shape.HasTextFrame -eq -1){$old=$shape.TextFrame.TextRange.Text;$shape.TextFrame.TextRange.Text=$old+' ';$shape.TextFrame.TextRange.Text=$old}
    $operations.Add(@{slide=$i;group_ungroup_move_text='pass'})
   } catch {$findings.Add(@{slide=$i;kind='edit-operation';error=$_.Exception.Message})}
   finally {$duplicate.Delete()}
  }
 }
 $result=@{input=$inputPath;source='PowerPoint.TextRange';errors=@();sha256=(Get-FileHash -LiteralPath $inputPath -Algorithm SHA256).Hash.ToLowerInvariant();slides=$count;findings=@($findings.ToArray());measurements=$measurements;operations=@($operations.ToArray());scope='PowerPoint open/export; shape text bounds; duplicate-slide group/ungroup/move/text. Charts and placeholders have separate edit restrictions.'}
 $result|ConvertTo-Json -Depth 15|Set-Content -LiteralPath (Join-Path $outPath 'qa.json') -Encoding utf8
 Write-Output "Slides=$count Findings=$($findings.Count) Native-edit-tests=$($operations.Count)"
 if($findings.Count -gt 0){exit 2}
}finally{
 if($deck){$deck.Close();[void][Runtime.InteropServices.Marshal]::ReleaseComObject($deck)}
 if($app){if($app.Presentations.Count -eq 0){$app.Quit()};[void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)}
}
