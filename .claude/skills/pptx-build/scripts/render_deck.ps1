# render_deck.ps1 — pptx를 PowerPoint COM으로 슬라이드별 PNG 내보내기
# 사용: pwsh -File render_deck.ps1 -Pptx <경로> -OutDir <폴더> [-Width 1920]
param(
    [Parameter(Mandatory = $true)][string]$Pptx,
    [Parameter(Mandatory = $true)][string]$OutDir,
    [int]$Width = 1920
)
$ErrorActionPreference = "Stop"
$PptxAbs = (Resolve-Path $Pptx).Path
New-Item -ItemType Directory -Force $OutDir | Out-Null
$OutAbs = (Resolve-Path $OutDir).Path

# PowerPoint COM은 단일 인스턴스 — 동시 렌더가 서로의 Quit에 죽지 않게 전역 뮤텍스로 직렬화
$mutex = New-Object System.Threading.Mutex($false, "Global\pptmaker_render_deck")
[void]$mutex.WaitOne()

$pp = New-Object -ComObject PowerPoint.Application
try {
    # Open(FileName, ReadOnly, Untitled, WithWindow)
    $pres = $pp.Presentations.Open($PptxAbs, $true, $false, $false)
    $h = [int]($Width * $pres.PageSetup.SlideHeight / $pres.PageSetup.SlideWidth)
    foreach ($slide in $pres.Slides) {
        $n = $slide.SlideIndex
        $out = Join-Path $OutAbs ("s{0:d2}.png" -f $n)
        $slide.Export($out, "PNG", $Width, $h)
    }
    $pres.Close()
    Write-Output "exported $($OutAbs)"
}
finally {
    $pp.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($pp)
    $mutex.ReleaseMutex()
    $mutex.Dispose()
}
