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

$powerPoint = $null
$presentation = $null
try {
    $powerPoint = New-Object -ComObject PowerPoint.Application
    foreach ($openDeck in $powerPoint.Presentations) {
        if ($openDeck.FullName -eq $inputPath) { throw 'Input is already open; render a saved working copy.' }
    }
    $presentation = $powerPoint.Presentations.Open($inputPath, $true, $false, $false)
    foreach ($slide in $presentation.Slides) {
        $file = Join-Path $outputPath ("slide-{0}.png" -f $slide.SlideIndex)
        $slide.Export($file, "PNG", $Width, $Height)
        if (-not [System.IO.File]::Exists($file)) {
            throw "PowerPoint did not export slide $($slide.SlideIndex): $file"
        }
        Write-Output $file
    }
}
finally {
    if ($presentation -ne $null) { $presentation.Close() }
    # Do not Quit: COM may attach to a user session containing other decks.
    if ($presentation -ne $null) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($presentation) }
    if ($powerPoint -ne $null) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($powerPoint) }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
