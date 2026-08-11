param(
    [Parameter(Mandatory = $true)]
    [string]$PptxPath,

    [Parameter(Mandatory = $true)]
    [string]$ManifestPath
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

if (-not (Test-Path -LiteralPath $PptxPath -PathType Leaf)) {
    throw "PPTX file not found: $PptxPath"
}
if (-not (Test-Path -LiteralPath $ManifestPath -PathType Leaf)) {
    throw "Animation manifest not found: $ManifestPath"
}

$manifest = Get-Content -LiteralPath $ManifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
$powerPoint = $null
$presentation = $null
$applied = 0

try {
    $powerPoint = New-Object -ComObject PowerPoint.Application
    # ReadOnly, Untitled, WithWindow. Opening without a window keeps the
    # post-processing step unobtrusive when conversion is launched by double-click.
    $presentation = $powerPoint.Presentations.Open(
        [System.IO.Path]::GetFullPath($PptxPath),
        0,
        0,
        0
    )

    foreach ($slideSpec in @($manifest.slides)) {
        if (@($slideSpec.effects).Count -eq 0) {
            continue
        }
        $slide = $presentation.Slides.Item([int]$slideSpec.slide)
        $sequence = $slide.TimeLine.MainSequence

        foreach ($effectSpec in @($slideSpec.effects)) {
            $shape = $null
            for ($shapeIndex = 1; $shapeIndex -le $slide.Shapes.Count; $shapeIndex++) {
                $candidate = $slide.Shapes.Item($shapeIndex)
                if ($candidate.Name -eq [string]$effectSpec.objectName) {
                    $shape = $candidate
                    break
                }
            }
            if ($null -eq $shape) {
                throw "Animation target not found on slide $($slideSpec.slide): $($effectSpec.objectName)"
            }

            # Level 0 targets the whole shape. Trigger 2 is With Previous, so
            # CSS delays stay relative to slide entry instead of requiring clicks.
            $effect = $sequence.AddEffect(
                $shape,
                [int]$effectSpec.effectType,
                0,
                2
            )
            if ([int]$effectSpec.direction -ne 0) {
                try {
                    $effect.EffectParameters.Direction = [int]$effectSpec.direction
                } catch {
                    # Some PowerPoint effects expose no direction parameter.
                    # The effect itself is still valid and should be retained.
                }
            }
            # Set timing after direction: PowerPoint resets an effect to its
            # preset duration whenever Direction changes.
            $effect.Timing.Duration = [single]$effectSpec.duration
            $effect.Timing.TriggerDelayTime = [single]$effectSpec.delay
            $applied++
        }
    }

    $presentation.Save()
    Write-Output "$applied effects saved"
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
