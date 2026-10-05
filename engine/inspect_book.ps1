<#
雛形・成果物の構造を調べる。セル番地は雛形が変わればずれるので、編集の前に必ずこれで確認する。

  # シート一覧と印刷設定
  powershell -ExecutionPolicy Bypass -File inspect_book.ps1 -File "…\（3）現場組織表.xls"

  # 1シートを詳しく（結合・塗りつぶし・図形・行高列幅・入力規則）
  powershell -ExecutionPolicy Bypass -File inspect_book.ps1 -File "…\（10）緊急時の体制及び対応.xls" `
      -Sheet "緊急時の体制及び対応" -Colors -Shapes -Validation -Grid 40

  # 特定セルの結合範囲・位置・値
  powershell -ExecutionPolicy Bypass -File inspect_book.ps1 -File "…" -Sheet "…" -Cells "B7,F24,A26,L1"

-Colors は「塗りつぶしがあり白でないセル」を色ごとにまとめて出す。
雛形では**記入すべき欄が色分けされている**ことが多く、記入位置の特定に使える。
#>
param(
    [Parameter(Mandatory = $true)][string]$File,
    [string]$Sheet,
    [string]$Cells,
    [int]$Grid = 0,          # 行高・列幅を出す範囲（行数）
    [int]$MaxColors = 200,
    [switch]$Colors,
    [switch]$Shapes,
    [switch]$Validation,
    [string]$Out
)

$ErrorActionPreference = 'Stop'
$lines = New-Object System.Collections.ArrayList
function L($s) { [void]$lines.Add([string]$s) }

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
    $bk = $xl.Workbooks.Open((Resolve-Path $File).Path)
    L ('===== ' + (Split-Path -Leaf $File) + ' =====')

    # ---- 名前定義
    $nm = @()
    foreach ($n in $bk.Names) { $nm += ($n.Name + ' = ' + $n.RefersTo) }
    if ($nm.Count) { L ('名前定義: ' + ($nm -join ' / ')) }

    $targets = @()
    if ($Sheet) { $targets += (Get-Sheet $bk $Sheet) } else { foreach ($w in $bk.Worksheets) { $targets += $w } }

    foreach ($ws in $targets) {
        $ps = $ws.PageSetup
        $ur = $ws.UsedRange
        L ''
        L ('--- [' + $ws.Name + ']  visible=' + $ws.Visible + '  used=' + $ur.Address($false, $false) +
           '  rows=' + $ur.Rows.Count + ' cols=' + $ur.Columns.Count + '  shapes=' + $ws.Shapes.Count)
        L ('    PrintArea=' + $ps.PrintArea + '  TitleRows=' + $ps.PrintTitleRows +
           '  Orientation=' + $ps.Orientation + '  Zoom=' + $ps.Zoom +
           '  FitW=' + $ps.FitToPagesWide + ' FitH=' + $ps.FitToPagesTall +
           '  改ページ(横)=' + $ws.HPageBreaks.Count)

        if ($Grid -gt 0) {
            $nc = [math]::Min($ur.Column + $ur.Columns.Count - 1, 26)
            L ('    列幅: ' + ((1..$nc | ForEach-Object {
                        $letter = ($ws.Cells(1, $_).Address($false, $false) -replace '\d', '')
                        $letter + '=' + [math]::Round($ws.Columns($_).ColumnWidth, 1)
                    }) -join ' '))
            L ('    行高: ' + ((1..$Grid | ForEach-Object { "$_=" + [math]::Round($ws.Rows($_).RowHeight, 0) }) -join ' '))
        }

        if ($Cells) {
            foreach ($a in ($Cells -split ',')) {
                $a = $a.Trim()
                if (-not $a) { continue }
                $c = $ws.Range($a)
                L ('    cell ' + $a + ' merge=' + $c.MergeArea.Address($false, $false) +
                   ' L=' + [math]::Round($c.Left, 1) + ' T=' + [math]::Round($c.Top, 1) +
                   ' W=' + [math]::Round($c.MergeArea.Width, 1) + ' H=' + [math]::Round($c.MergeArea.Height, 1) +
                   ' color=' + $c.Interior.Color + ' wrap=' + $c.WrapText +
                   " val='" + $c.Value2 + "'")
            }
        }

        if ($Colors) {
            $r2 = $ur.Row + $ur.Rows.Count - 1
            $c2 = $ur.Column + $ur.Columns.Count - 1
            $byColor = @{}
            $n = 0
            for ($r = $ur.Row; $r -le $r2; $r++) {
                for ($c = $ur.Column; $c -le $c2; $c++) {
                    $cell = $ws.Cells($r, $c)
                    if ($cell.Interior.Pattern -eq -4142) { continue }      # xlNone（塗りなし）
                    $col = [int]$cell.Interior.Color
                    if ($col -eq 16777215) { continue }                     # 白
                    $anchor = $cell.MergeArea.Cells(1, 1)
                    if ($anchor.Row -ne $r -or $anchor.Column -ne $c) { continue }   # 結合の非アンカーは飛ばす
                    if (-not $byColor.ContainsKey($col)) { $byColor[$col] = New-Object System.Collections.ArrayList }
                    [void]$byColor[$col].Add($cell.Address($false, $false) + "='" + $cell.Value2 + "'")
                    $n++
                    if ($n -ge $MaxColors) { break }
                }
                if ($n -ge $MaxColors) { break }
            }
            foreach ($k in ($byColor.Keys | Sort-Object)) {
                $b = $k -band 0xFF; $g = ($k -shr 8) -band 0xFF; $rr = ($k -shr 16) -band 0xFF
                # Excel の色値は 0xBBGGRR
                L ('    color=' + $k + ' RGB(' + $b + ',' + $g + ',' + $rr + ') 件数=' + $byColor[$k].Count)
                L ('       ' + (($byColor[$k] | Select-Object -First 40) -join ' / '))
            }
            if ($n -ge $MaxColors) { L '    …(打ち切り。-MaxColors で増やせる)' }
        }

        if ($Shapes) {
            $i = 0
            foreach ($s in $ws.Shapes) {
                $i++
                if ($i -gt 60) { L '    …(以下省略)'; break }
                $t = ''
                try { $t = $s.TextFrame2.TextRange.Text } catch {}
                $lc = ''
                try { $lc = $s.Line.ForeColor.RGB } catch {}
                # Type: 13=画像 9=線 17=テキストボックス
                L ("    shape '" + $s.Name + "' type=" + $s.Type + ' line=' + $lc +
                   ' L=' + [math]::Round($s.Left, 1) + ' T=' + [math]::Round($s.Top, 1) +
                   ' W=' + [math]::Round($s.Width, 1) + ' H=' + [math]::Round($s.Height, 1) +
                   " txt='" + $t + "'")
            }
        }

        if ($Validation) {
            try {
                $v = $ws.Cells.SpecialCells(-4174)      # xlCellTypeAllValidation
                L ('    入力規則: ' + $v.Address($false, $false))
            }
            catch { L '    入力規則: なし' }
        }
    }
    $bk.Close($false)
}
finally {
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}

if ($Out) {
    $lines -join "`r`n" | Out-File -FilePath $Out -Encoding utf8
    Write-Output ('wrote ' + $Out + ' (' + $lines.Count + ' lines)')
}
else { $lines | ForEach-Object { Write-Output $_ } }
