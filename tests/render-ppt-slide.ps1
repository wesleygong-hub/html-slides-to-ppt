param(
    [Parameter(Mandatory = $true)]
    [string]$PptxPath,

    [Parameter(Mandatory = $true)]
    [string]$PngPath
)

$ErrorActionPreference = "Stop"
$powerPoint = $null
$presentation = $null

try {
    $powerPoint = New-Object -ComObject PowerPoint.Application
    $presentation = $powerPoint.Presentations.Open(
        [System.IO.Path]::GetFullPath($PptxPath),
        1,
        0,
        0
    )
    $presentation.Slides.Item(1).Export(
        [System.IO.Path]::GetFullPath($PngPath),
        "PNG",
        1920,
        1080
    )
} finally {
    if ($null -ne $presentation) {
        try { $presentation.Close() } catch {}
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($presentation)
    }
    if ($null -ne $powerPoint) {
        try { $powerPoint.Quit() } catch {}
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($powerPoint)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
