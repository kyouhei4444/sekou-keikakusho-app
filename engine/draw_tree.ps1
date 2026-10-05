<#
組織表の系統線（親の枠から子の枠へつなぐ線）を引く。
下請業者を色分けして示す場合などに使う。セル位置は結合を考慮して取る。

  powershell -ExecutionPolicy Bypass -File draw_tree.ps1 -Json 系統線.json

線は「親の下端→中間の高さ」「横バー」「各子の上端へ」に分けて引く（bracket）。
direction を right にすると「親の右端→中間→各子の左端へ」（横書きの組織表向き）。

JSON（詳細は scripts/README.md）:
{
  "files": [
    { "file": "出力\\（3）現場組織表.xls",
      "sheet": "（3）現場組織表",
      "deleteNamePrefix": "下請業者線",
      "trees": [
        { "name": "下請業者線", "parent": "F24",
          "children": ["A26:B26", "E26:F26", "I26:J26"],
          "color": "blue", "weight": 1.25, "direction": "down" }
      ] }
  ]
}
color は "blue" "red" "black" "green" / [r,g,b] / Excel の色値（数値）。
#>
param([Parameter(Mandatory = $true)][string]$Json)

$ErrorActionPreference = 'Stop'
$cfg = Get-Content -Path $Json -Raw -Encoding UTF8 | ConvertFrom-Json

$NAMED = @{ black = 0; blue = 16711680; red = 255; green = 32768; orange = 26367; gray = 8421504 }

function To-Color($c) {
    if ($null -eq $c) { return 0 }
    if ($c -is [array]) { return ([int]$c[0]) + ([int]$c[1]) * 256 + ([int]$c[2]) * 65536 }   # RGB→0xBBGGRR
    if ($c -is [int] -or $c -is [long] -or $c -is [double]) { return [int]$c }
    $k = ([string]$c).ToLower()
    if ($NAMED.ContainsKey($k)) { return $NAMED[$k] }
    if ($k -match '^#?([0-9a-f]{2})([0-9a-f]{2})([0-9a-f]{2})$') {
        return [int]('0x' + $matches[3]) + [int]('0x' + $matches[2]) * 256 + [int]('0x' + $matches[1]) * 65536
    }
    throw ("色を解釈できません: " + $c)
}

function Get-Sheet($bk, $s) {
    if ($s -is [int] -or ([string]$s) -match '^\d+$') { return $bk.Worksheets.Item([int]$s) }
    try { return $bk.Worksheets.Item([string]$s) } catch {}
    foreach ($w in $bk.Worksheets) { if ($w.Name.Trim() -eq ([string]$s).Trim()) { return $w } }
    throw ("シートが見つかりません: '" + $s + "'")
}

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
try {
    foreach ($f in @($cfg.files)) {
        if ($null -eq $f) { continue }
        $bk = $xl.Workbooks.Open((Resolve-Path $f.file).Path)
        $ws = Get-Sheet $bk $f.sheet

        # 引き直しのため、前回引いた線を消す（名前の接頭辞で判別）
        $prefixes = @()
        if ($f.deleteNamePrefix) { $prefixes += [string]$f.deleteNamePrefix }
        foreach ($t in @($f.trees)) { if ($t.name) { $prefixes += [string]$t.name } }
        foreach ($s in @($ws.Shapes)) {
            foreach ($p in $prefixes) { if ($s.Name -like ($p + '*')) { $s.Delete(); break } }
        }

        foreach ($t in @($f.trees)) {
            if ($null -eq $t) { continue }
            $col = To-Color $t.color
            $w   = if ($null -ne $t.weight) { [double]$t.weight } else { 1.25 }
            $nm  = if ($t.name) { [string]$t.name } else { '系統線' }
            $dir = if ($t.direction) { [string]$t.direction } else { 'down' }

            $p = $ws.Range($t.parent).MergeArea
            $segs = @()

            if ($dir -eq 'right') {
                $xP = $p.Left + $p.Width
                $yP = $p.Top + $p.Height / 2
                $ys = @(); $xLeft = 0
                foreach ($k in @($t.children)) {
                    $r = $ws.Range($k)
                    $ys += ($r.Top + $r.Height / 2)
                    $xLeft = $r.Left
                }
                $xMid = ($xP + $xLeft) / 2
                $segs += , @($xP, $yP, $xMid, $yP)
                $segs += , @($xMid, ($ys | Measure-Object -Minimum).Minimum, $xMid, ($ys | Measure-Object -Maximum).Maximum)
                foreach ($y in $ys) { $segs += , @($xMid, $y, $xLeft, $y) }
            }
            else {
                $xP = $p.Left + $p.Width / 2
                $yP = $p.Top + $p.Height
                $xs = @(); $yTop = 0
                foreach ($k in @($t.children)) {
                    $r = $ws.Range($k)
                    $xs += ($r.Left + $r.Width / 2)
                    $yTop = $r.Top
                }
                $yMid = ($yP + $yTop) / 2
                $segs += , @($xP, $yP, $xP, $yMid)
                $segs += , @(($xs | Measure-Object -Minimum).Minimum, $yMid, ($xs | Measure-Object -Maximum).Maximum, $yMid)
                foreach ($x in $xs) { $segs += , @($x, $yMid, $x, $yTop) }
            }

            $i = 0
            foreach ($g in $segs) {
                $i++
                $ln = $ws.Shapes.AddLine($g[0], $g[1], $g[2], $g[3])
                $ln.Name = $nm + $i
                $ln.Line.ForeColor.RGB = $col
                $ln.Line.Weight = $w
                $ln.Placement = 1      # xlMoveAndSize
            }
            Write-Output ("  " + $nm + " : " + $segs.Count + "本（親 " + $t.parent + " → 子 " + @($t.children).Count + "）")
        }

        $bk.Worksheets.Item(1).Activate() | Out-Null
        $bk.Save(); $bk.Close($false)
        Write-Output ((Split-Path -Leaf $f.file) + " 更新")
    }
}
finally {
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}
