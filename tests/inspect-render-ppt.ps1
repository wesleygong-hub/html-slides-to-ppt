param(
    [Parameter(Mandatory = $true)][string]$PptxPath,
    [Parameter(Mandatory = $true)][string]$OutputDirectory,
    [int]$Width = 1920,
    [int]$Height = 1080
)
$ErrorActionPreference = "Stop"
$app = $null
$deck = $null
try {
    [void][IO.Directory]::CreateDirectory([IO.Path]::GetFullPath($OutputDirectory))
    $app = New-Object -ComObject PowerPoint.Application
    $deck = $app.Presentations.Open([IO.Path]::GetFullPath($PptxPath), 1, 0, 0)
    $slides = @()
    for ($i = 1; $i -le $deck.Slides.Count; $i++) {
        $slide = $deck.Slides.Item($i)
        $slide.Export((Join-Path $OutputDirectory ("slide-{0:D3}.png" -f $i)), "PNG", $Width, $Height)
        $text = @()
        foreach ($shape in $slide.Shapes) {
            if ($shape.HasTextFrame -eq -1 -and $shape.TextFrame2.HasText -eq -1) {
                $range = $shape.TextFrame2.TextRange
                $text += [pscustomobject]@{
                    name = $shape.Name
                    text = $range.Text
                    font = $range.Font.Name
                    size = [double]$range.Font.Size
                    x = [double]$shape.Left
                    y = [double]$shape.Top
                    w = [double]$shape.Width
                    h = [double]$shape.Height
                }
            }
        }
        $slides += [pscustomobject]@{ slide = $i; shapes = $slide.Shapes.Count; effects = $slide.TimeLine.MainSequence.Count; text = $text }
    }
    ConvertTo-Json -InputObject $slides -Depth 5 -Compress
} finally {
    if ($null -ne $deck) {
        try { $deck.Close() } catch {}
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($deck)
    }
    if ($null -ne $app) {
        try { $app.Quit() } catch {}
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
