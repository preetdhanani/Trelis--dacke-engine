# Open S1_negative_control.pptx (already_filled / naive_reclone / fixed_reclone)
# in real PowerPoint and see what actually happens to the naive re-clone: does
# PowerPoint show a repair prompt, a broken-image icon, or something else.

$outPath = "P:\decke\spike\output\S1_negative_control.pptx"
$shotDir = "P:\decke\spike\output\screenshots_negative_control"
New-Item -ItemType Directory -Force -Path $shotDir | Out-Null

$ppt = New-Object -ComObject PowerPoint.Application
$ppt.DisplayAlerts = 1

$openError = $null
try {
    $pres = $ppt.Presentations.Open($outPath, $true, $false, $false)
} catch {
    $openError = $_.Exception.Message
}

if ($openError) {
    Write-Output "OPEN_RESULT: FAILED -- $openError"
} else {
    Write-Output "OPEN_RESULT: OK, no exception thrown on open"
    Write-Output "SLIDE_COUNT: $($pres.Slides.Count)"
    for ($i = 1; $i -le $pres.Slides.Count; $i++) {
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
