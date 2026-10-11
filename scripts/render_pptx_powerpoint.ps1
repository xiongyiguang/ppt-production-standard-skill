param(
    [Parameter(Mandatory = $true)] [string] $InputPptx,
    [Parameter(Mandatory = $true)] [string] $OutputDir,
    [int] $Width = 1920,
    [int] $Height = 1080
)

$ErrorActionPreference = 'Stop'

$inputPath = [System.IO.Path]::GetFullPath($InputPptx)
$outputPath = [System.IO.Path]::GetFullPath($OutputDir)
if (-not [System.IO.File]::Exists($inputPath)) {
    throw "Input presentation does not exist: $inputPath"
}
[System.IO.Directory]::CreateDirectory($outputPath) | Out-Null
if ((Get-ChildItem -LiteralPath $outputPath -Force | Measure-Object).Count -ne 0) {
    throw "Output directory must be new or empty to prevent stale preview files: $outputPath"
}

$beforeHash = (Get-FileHash -LiteralPath $inputPath -Algorithm SHA256).Hash.ToLowerInvariant()
$localDir = Join-Path ([IO.Path]::GetTempPath()) ('ppt-standard-render-' + [Guid]::NewGuid().ToString('N'))
$localFile = Join-Path $localDir 'working-copy.pptx'
$powerPoint = $null
$presentation = $null
try {
    $powerPoint = New-Object -ComObject PowerPoint.Application
    [IO.Directory]::CreateDirectory($localDir) | Out-Null
    Copy-Item -LiteralPath $inputPath -Destination $localFile
    if ((Get-FileHash -LiteralPath $localFile -Algorithm SHA256).Hash.ToLowerInvariant() -ne $beforeHash) { throw 'Source changed while making the independent copy.' }
    # Parallels caller must use --current-user; local copy + visible read-only window.
    $presentation = $powerPoint.Presentations.Open($localFile, -1, 0, -1)
    foreach ($slide in $presentation.Slides) {
        $file = Join-Path $outputPath ("slide-{0}.png" -f $slide.SlideIndex)
        $slide.Export($file, "PNG", $Width, $Height)
        if (-not [System.IO.File]::Exists($file)) {
            throw "PowerPoint did not export slide $($slide.SlideIndex): $file"
        }
        Write-Output $file
    }
    if ((Get-FileHash -LiteralPath $inputPath -Algorithm SHA256).Hash.ToLowerInvariant() -ne $beforeHash) { throw 'Source changed during rendering; evidence is stale.' }
    $report = @{sha256=$beforeHash; powerpointVersion=$powerPoint.Version; slideCount=$presentation.Slides.Count;
                openMode=@{readOnly=$true; withWindow=$true; localWorkingCopy=$true}; width=$Width; height=$Height;
                visualReview='not_performed'}
    $reportPath = Join-Path $outputPath 'powerpoint-render.json'
    [IO.File]::WriteAllText($reportPath, ($report | ConvertTo-Json -Depth 6), [Text.UTF8Encoding]::new($false))
}
finally {
    if ($presentation -ne $null) { $presentation.Saved = -1; $presentation.Close() }
    # Do not Quit: COM may attach to a user session containing other decks.
    if ($presentation -ne $null) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($presentation) }
    if ($powerPoint -ne $null) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($powerPoint) }
    # Remove only this invocation's generated local copy; never remove the source or output.
    if (Test-Path -LiteralPath $localFile) { Remove-Item -LiteralPath $localFile -Force }
    if (Test-Path -LiteralPath $localDir) { Remove-Item -LiteralPath $localDir }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
