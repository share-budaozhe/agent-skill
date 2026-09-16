# Build Visio diagram from JSON layout data (COM automation, headless).
param(
    [string]$JsonPath,
    [string]$OutPath
)

$ErrorActionPreference = "Stop"

function RgbStr($arr) {
    return ("RGB({0},{1},{2})" -f [int]$arr[0], [int]$arr[1], [int]$arr[2])
}

function SetText($shp, $fs, $fc, $bold) {
    $shp.CellsU("Char.Size").FormulaU = ("{0} pt" -f $fs)
    $shp.CellsU("Char.Color").FormulaU = (RgbStr $fc)
    if ($script:FontId -ne $null) { $shp.CellsU("Char.Font").FormulaU = [string]$script:FontId }
    if ($bold) { $shp.CellsU("Char.Style").FormulaU = "1" } else { $shp.CellsU("Char.Style").FormulaU = "0" }
}

function DrawBox($pg, $s, $pH) {
    $x = [double]$s.x
    $w = [double]$s.w
    $h = [double]$s.h
    $top = [double]$s.y
    $y = $pH - $top - $h
    # oval=true 时画椭圆（JSON 里可选字段，缺省为矩形）
    if ($s.PSObject.Properties.Name -contains 'oval' -and [bool]$s.oval) {
        $shp = $pg.DrawOval($x, $y, ($x + $w), ($y + $h))
    } else {
        $shp = $pg.DrawRectangle($x, $y, ($x + $w), ($y + $h))
    }

    $shp.CellsU("FillPattern").FormulaU = "1"
    $shp.CellsU("FillForegnd").FormulaU = (RgbStr $s.fill)
    if ([int]$s.dash -eq 0) {
        $shp.CellsU("LinePattern").FormulaU = "0"
    } elseif ([int]$s.dash -ge 2) {
        # dash>=2: 直接透传 Visio LinePattern 编号（2=虚线 3=点线 ...），向后兼容
        $shp.CellsU("LinePattern").FormulaU = [string][int]$s.dash
        $shp.CellsU("LineColor").FormulaU = (RgbStr $s.line)
    } else {
        $shp.CellsU("LinePattern").FormulaU = "1"
        $shp.CellsU("LineColor").FormulaU = (RgbStr $s.line)
    }
    $shp.CellsU("LineWeight").FormulaU = ("{0} pt" -f $s.lw)
    $rd = [double]$s.rnd
    if ($rd -gt 0) { $shp.CellsU("Rounding").FormulaU = ("{0} in" -f $rd) }

    if ($s.text -ne $null -and [string]$s.text -ne "") {
        $shp.Text = [string]$s.text
    }
    SetText $shp $s.fs $s.fc ([bool]$s.bold)
    $shp.CellsU("VerticalAlign").FormulaU = [string]$s.valign
    $shp.CellsU("Para.HorzAlign").FormulaU = [string]$s.align
    $shp.CellsU("LeftMargin").FormulaU = "0.03 in"
    $shp.CellsU("RightMargin").FormulaU = "0.03 in"
    $shp.CellsU("TopMargin").FormulaU = "0.015 in"
    $shp.CellsU("BottomMargin").FormulaU = "0.015 in"
    return $shp
}

function DrawConn($pg, $c, $pH) {
    $x1 = [double]$c.x1; $x2 = [double]$c.x2
    $y1 = $pH - [double]$c.y1; $y2 = $pH - [double]$c.y2
    $shp = $pg.DrawLine($x1, $y1, $x2, $y2)
    $shp.CellsU("LineColor").FormulaU = (RgbStr $c.color)
    $shp.CellsU("LineWeight").FormulaU = ("{0} pt" -f $c.lw)
    $shp.CellsU("LinePattern").FormulaU = [string]$c.dash
    $a = [string]$c.arrow
    $endV = "0"; $begV = "0"
    if ($a -eq "end")   { $endV = "4" }
    if ($a -eq "start") { $begV = "4" }
    if ($a -eq "both")  { $endV = "4"; $begV = "4" }
    $shp.CellsU("EndArrow").FormulaU = $endV
    $shp.CellsU("BeginArrow").FormulaU = $begV
    $shp.CellsU("EndArrowSize").FormulaU = "2"
    $shp.CellsU("BeginArrowSize").FormulaU = "2"
    if ($c.label -ne $null -and [string]$c.label -ne "") {
        $shp.Text = [string]$c.label
        SetText $shp $c.fs $c.fc $false
        $shp.CellsU("TxtAngle").FormulaU = "0 deg"
        $shp.CellsU("VerticalAlign").FormulaU = "0"
        $shp.CellsU("Para.HorzAlign").FormulaU = "1"
    }
    return $shp
}

# --- load data ---
if (-not (Test-Path -LiteralPath $JsonPath)) { throw "JSON not found: $JsonPath" }
$raw = Get-Content -LiteralPath $JsonPath -Raw -Encoding UTF8
$data = $raw | ConvertFrom-Json

$visio = New-Object -ComObject Visio.Application
$visio.Visible = $false
$visio.AlertResponse = 1

try {
    $doc = $visio.Documents.Add("")
    $first = $true

    # resolve CJK-friendly font id (Char.Font takes a numeric font id, not a name)
    $script:FontId = $null
    try {
        $fnt = $doc.Fonts.Item("Microsoft YaHei")
        if ($fnt -ne $null) { $script:FontId = $fnt.ID }
    } catch { }
    if ($script:FontId -eq $null) {
        foreach ($f in $doc.Fonts) {
            if ($f.Name -like "*YaHei*" -or $f.Name -like "*DengXian*" -or $f.Name -like "*SimSun*") {
                $script:FontId = $f.ID; break
            }
        }
    }
    Write-Output ("font id = " + $script:FontId)

    foreach ($pgData in $data.pages) {
        if ($first) { $pg = $doc.Pages.Item(1); $first = $false } else { $pg = $doc.Pages.Add() }
        $pg.Name = [string]$pgData.name
        $pg.PageSheet.CellsU("PageWidth").FormulaU  = ("{0} in" -f $data.pageW)
        $pg.PageSheet.CellsU("PageHeight").FormulaU = ("{0} in" -f $data.pageH)

        foreach ($s in $pgData.shapes) { [void](DrawBox $pg $s ([double]$data.pageH)) }
        foreach ($c in $pgData.conns)  { [void](DrawConn $pg $c ([double]$data.pageH)) }
        Write-Output ("page done: " + $pg.Name + "  shapes=" + $pgData.shapes.Count + " conns=" + $pgData.conns.Count)
    }

    # jump to first page
    if ($doc.Pages.Count -gt 0) {
        $w = $visio.ActiveWindow
        if ($w -ne $null) { $w.Page = $doc.Pages.Item(1) }
    }

    if (Test-Path -LiteralPath $OutPath) { Remove-Item -LiteralPath $OutPath -Force }
    $doc.SaveAs($OutPath)
    Write-Output ("SAVED: " + $OutPath)
    $doc.Close()
} finally {
    $visio.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($visio) | Out-Null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
