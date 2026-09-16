param([string]$VsdxPath, [string]$OutDir)
$ErrorActionPreference = "Stop"
if (-not (Test-Path -LiteralPath $OutDir)) { New-Item -ItemType Directory -Path $OutDir -Force | Out-Null }

$visio = New-Object -ComObject Visio.Application
$visio.Visible = $false
$visio.AlertResponse = 1
try {
    $doc = $visio.Documents.Open($VsdxPath)
    $i = 0
    foreach ($pg in $doc.Pages) {
        $i++
        $name = ("P{0}.png" -f $i)
        $p = Join-Path $OutDir $name
        $pg.Export($p)
        Write-Output ("exported " + $p + "  (" + $pg.Name + ")")
    }
    $doc.Close()
} finally {
    $visio.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($visio) | Out-Null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
