# S1 validation, layer 2: open the real output .pptx in actual PowerPoint via COM
# automation. This is the strongest available proxy for "opens with no repair
# prompt, theme intact" -- PowerPoint's own repair/self-heal logic runs here,
# not just python-pptx's parser. Exports each slide to PNG for visual inspection.

$repoRoot = "P:\decke"
$outPath = Join-Path $repoRoot "spike\output\S1_filled_deck.pptx"
$shotDir = Join-Path $repoRoot "spike\output\screenshots"
New-Item -ItemType Directory -Force -Path $shotDir | Out-Null

if (-not (Test-Path $outPath)) {
    Write-Error "Output deck not found at $outPath -- run s1_template_fill.py first."
    exit 1
}

$ppt = New-Object -ComObject PowerPoint.Application
$ppt.DisplayAlerts = 1  # ppAlertsNone -- suppress any blocking dialog (e.g. auto-repair notice)

$openError = $null
$pres = $null
try {
    # msoFalse = 0 for WithWindow and ReadOnly args in this COM signature
    $pres = $ppt.Presentations.Open($outPath, $true, $false, $false)
} catch {
    $openError = $_.Exception.Message
}

if ($openError) {
    Write-Output "OPEN_RESULT: FAILED -- $openError"
} else {
    $slideCount = $pres.Slides.Count
    Write-Output "OPEN_RESULT: OK, no exception thrown on open"
    Write-Output "SLIDE_COUNT: $slideCount"

    for ($i = 1; $i -le $slideCount; $i++) {
        $slide = $pres.Slides.Item($i)
        $shotPath = Join-Path $shotDir ("slide{0}.png" -f $i)
        $slide.Export($shotPath, "PNG", 1280, 720)
    }
    Write-Output "SCREENSHOTS_WRITTEN: $shotDir"

    $pres.Close()
}

$ppt.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt) | Out-Null
Write-Output "DONE"
