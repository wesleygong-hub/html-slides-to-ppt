param(
    [Parameter(Mandatory = $true)]
    [string]$PptxPath
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
    $items = @()
    for ($slideIndex = 1; $slideIndex -le $presentation.Slides.Count; $slideIndex++) {
        $sequence = $presentation.Slides.Item($slideIndex).TimeLine.MainSequence
        for ($effectIndex = 1; $effectIndex -le $sequence.Count; $effectIndex++) {
            $effect = $sequence.Item($effectIndex)
            $items += [pscustomobject]@{
                slide = $slideIndex
                shape = [string]$effect.Shape.Name
                effectType = [int]$effect.EffectType
                direction = [int]$effect.EffectParameters.Direction
                duration = [double]$effect.Timing.Duration
                delay = [double]$effect.Timing.TriggerDelayTime
                trigger = [int]$effect.Timing.TriggerType
            }
        }
    }
    @($items) | ConvertTo-Json -Compress
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
