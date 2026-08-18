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
$verified = 0
$timingTolerance = 0.02

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
            $requestedDirection = [int]$effectSpec.direction
            $directionRequired = [bool]$effectSpec.directionRequired
            $directionWasSet = $false
            if ($requestedDirection -ne 0) {
                try {
                    $effect.EffectParameters.Direction = $requestedDirection
                    $directionWasSet = $true
                } catch {
                    if ($directionRequired) {
                        throw "Direction could not be set on slide $($slideSpec.slide), shape $($effectSpec.objectName): $($_.Exception.Message)"
                    }
                }
            }
            # Set timing after direction: PowerPoint resets an effect to its
            # preset duration whenever Direction changes.
            $effect.Timing.Duration = [single]$effectSpec.duration
            $effect.Timing.TriggerDelayTime = [single]$effectSpec.delay

            # COM readback catches unsupported directions and PowerPoint preset
            # resets before the presentation is saved.
            if ($directionWasSet) {
                try {
                    $actualDirection = [int]$effect.EffectParameters.Direction
                    if ($directionRequired -and $actualDirection -ne $requestedDirection) {
                        throw "Direction verification failed: expected $requestedDirection, got $actualDirection"
                    }
                } catch {
                    if ($directionRequired) {
                        throw "Direction verification failed on slide $($slideSpec.slide), shape $($effectSpec.objectName): $($_.Exception.Message)"
                    }
                }
            }
            $actualDuration = [double]$effect.Timing.Duration
            $actualDelay = [double]$effect.Timing.TriggerDelayTime
            if ([Math]::Abs($actualDuration - [double]$effectSpec.duration) -gt $timingTolerance) {
                throw "Duration verification failed on slide $($slideSpec.slide), shape $($effectSpec.objectName): expected $($effectSpec.duration), got $actualDuration"
            }
            if ([Math]::Abs($actualDelay - [double]$effectSpec.delay) -gt $timingTolerance) {
                throw "Delay verification failed on slide $($slideSpec.slide), shape $($effectSpec.objectName): expected $($effectSpec.delay), got $actualDelay"
            }
            $verified++
            $applied++
        }
    }

    $presentation.Save()
    Write-Output "$applied effects saved, $verified verified"
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
