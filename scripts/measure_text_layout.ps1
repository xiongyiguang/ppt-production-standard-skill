param(
    [Parameter(Mandatory = $true)] [string] $InputPptx,
    [Parameter(Mandatory = $true)] [string] $OutputJson
)
$ErrorActionPreference = 'Stop'
$inputPath = (Resolve-Path -LiteralPath $InputPptx).ProviderPath
$outputPath = [IO.Path]::GetFullPath($OutputJson)
if (Test-Path -LiteralPath $outputPath) { throw 'Use a new evidence filename; output already exists.' }
$beforeHash = (Get-FileHash -LiteralPath $inputPath -Algorithm SHA256).Hash.ToLowerInvariant()
$measurements = @{}
$errors = [Collections.Generic.List[string]]::new()
function Read-Text($shape, [string] $key) {
    if ($shape.HasTextFrame -ne -1 -or $shape.TextFrame.HasText -ne -1) { return }
    try {
        $range = $shape.TextFrame.TextRange
        $formats = @()
        for ($i = 1; $i -le $range.Paragraphs().Count; $i++) {
            $paragraph = $range.Paragraphs($i, 1)
            if ([string]::IsNullOrWhiteSpace($paragraph.Text)) { continue }
            $format = $paragraph.ParagraphFormat
            $formats += @{multiple=($format.LineRuleWithin -eq -1); within=[double]$format.SpaceWithin;
                          before=[double]$format.SpaceBefore; after=[double]$format.SpaceAfter}
        }
        $measurements[$key] = @{lines=[int]$range.Lines().Count; paragraphs=@($formats)}
    } catch { $errors.Add("${key}: $($_.Exception.Message)") }
}
function Read-Shape($shape, [int] $page) {
    $key = "${page}:$($shape.Id)"
    if ($shape.Type -eq 6) {
        for ($i = 1; $i -le $shape.GroupItems.Count; $i++) { Read-Shape $shape.GroupItems.Item($i) $page }
    } elseif ($shape.HasTable -eq -1) {
        for ($r = 1; $r -le $shape.Table.Rows.Count; $r++) {
            for ($c = 1; $c -le $shape.Table.Columns.Count; $c++) {
                Read-Text $shape.Table.Cell($r, $c).Shape "${key}/r${r}c${c}"
            }
        }
    } else { Read-Text $shape $key }
}
$localDir = Join-Path ([IO.Path]::GetTempPath()) ('ppt-standard-measure-' + [Guid]::NewGuid().ToString('N'))
$localFile = Join-Path $localDir 'working-copy.pptx'
$app = $null
$deck = $null
try {
    $app = New-Object -ComObject PowerPoint.Application
    [IO.Directory]::CreateDirectory($localDir) | Out-Null
    Copy-Item -LiteralPath $inputPath -Destination $localFile
    if ((Get-FileHash -LiteralPath $localFile -Algorithm SHA256).Hash.ToLowerInvariant() -ne $beforeHash) { throw 'Source changed while making the independent copy.' }
    # Parallels caller must use --current-user; local copy + visible read-only window.
    $deck = $app.Presentations.Open($localFile, -1, 0, -1)
    foreach ($slide in $deck.Slides) {
        foreach ($shape in $slide.Shapes) { Read-Shape $shape $slide.SlideIndex }
    }
    $afterHash = (Get-FileHash -LiteralPath $inputPath -Algorithm SHA256).Hash.ToLowerInvariant()
    if ($beforeHash -ne $afterHash) { throw 'Input changed during measurement; rerun on a stable file.' }
    $report = @{sha256=$beforeHash; source='PowerPoint.TextRange'; powerpointVersion=$app.Version; slideCount=$deck.Slides.Count;
               openMode=@{readOnly=$true; withWindow=$true; localWorkingCopy=$true}; measurements=$measurements; errors=@($errors.ToArray())}
    $json = $report | ConvertTo-Json -Depth 12
    $stream = [IO.File]::Open($outputPath, [IO.FileMode]::CreateNew)
    try {
        $bytes = [Text.UTF8Encoding]::new($false).GetBytes($json)
        $stream.Write($bytes, 0, $bytes.Length)
    } finally { $stream.Dispose() }
    Write-Output $outputPath
} finally {
    if ($null -ne $deck) { $deck.Saved = -1; $deck.Close(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($deck) }
    # Do not Quit: COM may attach to the user's existing PowerPoint process.
    if ($null -ne $app) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($app) }
    # Remove only this invocation's generated local copy.
    if (Test-Path -LiteralPath $localFile) { Remove-Item -LiteralPath $localFile -Force }
    if (Test-Path -LiteralPath $localDir) { Remove-Item -LiteralPath $localDir }
}
