<#
印刷体裁を整える。折り返し・行高の自動計算・縮小表示・印刷タイトル行・改ページ。
結合セルは AutoFit が効かないので、非表示の作業シートで必要高さを測って設定する。

  powershell -ExecutionPolicy Bypass -File fit_print.ps1 -Json 体裁.json

範囲文字列の {LAST} は、そのシートの使用範囲の最終行に置き換わる（例 "L5:L{LAST}"）。

JSON（詳細は scripts/README.md）:
{
  "files": [
    { "file": "出力\\（7）施工方法.xls",
      "sheets": [
        { "sheet": "*", "except": ["施工方法", "（7-2）仮設備計画", "詳細図"],
          "titleRows": "$1:$4", "fitWide": 1, "fitTall": false,
          "shrink": ["L5:L{LAST}"],
          "measureRows": [ { "row": 2, "cells": ["A2","D2","H2","J2"], "min": 34.5, "pad": 6 } ] },

        { "sheet": "（7-2）仮設備計画",
          "printArea": "$A$1:$C$43", "fitWide": 1, "fitTall": 1,
          "wrap": ["C4:C{LAST}"], "valignTop": ["B4:C{LAST}"],
          "font": [ { "range": "B4:C{LAST}", "color": 0 } ],
          "autoFitRows": { "range": "4:{LAST}", "min": 18 },
          "rowRules": [ { "range": "4:43", "ifEmpty": "C", "height": 13.5 },
                        { "range": "4:43", "ifValue": "B", "ifEmpty": "C", "height": 22 } ] },

        { "sheet": "(2)出来形管理基準",
          "titleRows": "$1:$3", "printArea": "$A$1:$O$30", "fitWide": 1, "fitTall": false,
          "wrap": ["A4:O30"], "valignTop": ["A4:O30"],
          "resetPageBreaks": true, "pageBreaks": [8, 17, 24, 30],
          "measureBlocks": [ { "row": 4, "rows": 3, "perRowCols": ["G","H","K","L"],
                               "blockCols": ["F","M","O"], "minRow": 18, "pad": 2,
                               "minBlock": 130 } ] }
      ] }
  ]
}
#>
param([Parameter(Mandatory = $true)][string]$Json)

$ErrorActionPreference = 'Stop'
$cfg = Get-Content -Path $Json -Raw -Encoding UTF8 | ConvertFrom-Json
$xlTop = -4160

function Get-Sheet($bk, $s) {
    if ($s -is [int] -or ([string]$s) -match '^\d+$') { return $bk.Worksheets.Item([int]$s) }
    try { return $bk.Worksheets.Item([string]$s) } catch {}
    foreach ($w in $bk.Worksheets) { if ($w.Name.Trim() -eq ([string]$s).Trim()) { return $w } }
    throw ("シートが見つかりません: '" + $s + "'")
}

function Last-Row($ws) {
    $ur = $ws.UsedRange
    return ($ur.Row + $ur.Rows.Count - 1)
}

# 関数名 R は Invoke-History の別名 r と衝突するので使えない
function Resolve-Addr($ws, [string]$addr) {
    return ($addr -replace '\{LAST\}', [string](Last-Row $ws))
}

function Set-RowHeight($ws, [int]$row, [double]$h) {
    if ($h -gt 409) { $h = 409 }      # Excel の行高上限
    if ($h -gt 0) { $ws.Rows($row).RowHeight = $h }
}

# 結合セルの必要高さを、非表示の作業シートで測る
function Measure-H($m, [string]$text, [double]$wpt, [string]$fn, [double]$fs) {
    if ([string]::IsNullOrWhiteSpace($text)) { return 0 }
    if ($wpt -le 0) { return 0 }
    $m.Visible = -1
    $m.Columns(1).ColumnWidth = 50
    for ($i = 0; $i -lt 4; $i++) {
        # ColumnWidth(文字数) と Width(pt) は非線形なので反復して合わせる
        $cur = $m.Columns(1).Width
        if ($cur -le 0) { break }
        $m.Columns(1).ColumnWidth = $m.Columns(1).ColumnWidth * $wpt / $cur
    }
    $c = $m.Range('A1')
    $c.Font.Name = $fn; $c.Font.Size = $fs; $c.WrapText = $true
    $c.Value2 = $text
    $m.Rows(1).AutoFit() | Out-Null
    $h = $m.Rows(1).RowHeight
    $c.ClearContents() | Out-Null
    $m.Visible = 0
    return $h
}

function Need-Cell($ws, $m, [string]$addr, [bool]$skipVMerge) {
    $c = $ws.Range($addr)
    $t = [string]$c.Value2
    if ([string]::IsNullOrWhiteSpace($t)) { return 0 }
    $ma = $c.MergeArea
    if ($skipVMerge -and $ma.Rows.Count -gt 1) { return 0 }     # 縦結合は1行分では測れない
    return (Measure-H $m $t $ma.Width $c.Font.Name $c.Font.Size)
}

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.ScreenUpdating = $false
try {
    foreach ($f in @($cfg.files)) {
        if ($null -eq $f) { continue }
        $bk = $xl.Workbooks.Open((Resolve-Path $f.file).Path)

        # 測定用の作業シート（最後に削除する）
        $m = $bk.Worksheets.Add()
        $m.Name = '_fit'
        # 作業シートを隠す前に、別の表示シートをアクティブにしておく
        foreach ($w in $bk.Worksheets) {
            if ($w.Name -ne '_fit' -and $w.Visible -eq -1) { $w.Activate() | Out-Null; break }
        }
        $m.Visible = 0

        foreach ($spec in @($f.sheets)) {
            if ($null -eq $spec) { continue }

            # 対象シートの決定（"*" は全シート、except で除外）
            $targets = @()
            if ([string]$spec.sheet -eq '*') {
                $ex = @($spec.except) | ForEach-Object { ([string]$_).Trim() }
                foreach ($w in $bk.Worksheets) {
                    if ($w.Name -eq '_fit') { continue }
                    if ($ex -contains $w.Name.Trim()) { continue }
                    $targets += $w
                }
            }
            else { $targets += (Get-Sheet $bk $spec.sheet) }

            foreach ($ws in $targets) {
                # ---- 文字の扱い
                foreach ($a in @($spec.wrap))      { if ($a) { $ws.Range((Resolve-Addr $ws $a)).WrapText = $true } }
                foreach ($a in @($spec.valignTop)) { if ($a) { $ws.Range((Resolve-Addr $ws $a)).VerticalAlignment = $xlTop } }
                foreach ($a in @($spec.shrink))    { if ($a) { $ws.Range((Resolve-Addr $ws $a)).ShrinkToFit = $true } }
                foreach ($g in @($spec.font)) {
                    if ($null -eq $g) { continue }
                    $rg = $ws.Range((Resolve-Addr $ws $g.range))
                    if ($null -ne $g.color) { $rg.Font.Color = [int]$g.color }
                    if ($null -ne $g.size)  { $rg.Font.Size = [double]$g.size }
                }

                # ---- 行高の自動調整（結合の無い範囲向け）
                if ($spec.autoFitRows) {
                    $rr = (Resolve-Addr $ws $spec.autoFitRows.range)
                    $ws.Rows($rr).AutoFit() | Out-Null
                    if ($null -ne $spec.autoFitRows.min) {
                        $p = $rr.Split(':')
                        for ($r = [int]$p[0]; $r -le [int]$p[1]; $r++) {
                            if ($ws.Rows($r).RowHeight -lt [double]$spec.autoFitRows.min) {
                                $ws.Rows($r).RowHeight = [double]$spec.autoFitRows.min
                            }
                        }
                    }
                }

                # ---- 条件つきの行高（空行を戻す／図の見出し行を高くする 等）
                foreach ($rule in @($spec.rowRules)) {
                    if ($null -eq $rule) { continue }
                    $p = (Resolve-Addr $ws $rule.range).Split(':')
                    for ($r = [int]$p[0]; $r -le [int]$p[1]; $r++) {
                        $ok = $true
                        if ($rule.ifEmpty) {
                            $v = [string]$ws.Range([string]$rule.ifEmpty + $r).Value2
                            if (-not [string]::IsNullOrWhiteSpace($v)) { $ok = $false }
                        }
                        if ($ok -and $rule.ifValue) {
                            $v = [string]$ws.Range([string]$rule.ifValue + $r).Value2
                            if ([string]::IsNullOrWhiteSpace($v)) { $ok = $false }
                        }
                        if ($ok) { Set-RowHeight $ws $r ([double]$rule.height) }
                    }
                }

                # ---- 1行の必要高さを測って設定（結合セルを含む見出し行など）
                foreach ($mr in @($spec.measureRows)) {
                    if ($null -eq $mr) { continue }
                    $need = if ($null -ne $mr.min) { [double]$mr.min } else { 0 }
                    foreach ($a in @($mr.cells)) {
                        $h = Need-Cell $ws $m $a $true
                        if ($h -gt $need) { $need = $h }
                    }
                    $pad = if ($null -ne $mr.pad) { [double]$mr.pad } else { 0 }
                    Set-RowHeight $ws ([int]$mr.row) ([math]::Ceiling($need) + $pad)
                    if (@($spec.valignTop).Count -eq 0) {
                        $ws.Rows([int]$mr.row).VerticalAlignment = -4108   # xlCenter
                    }
                }

                # ---- 工種ブロック（縦結合を含む表）の行高
                foreach ($mb in @($spec.measureBlocks)) {
                    if ($null -eq $mb) { continue }
                    $r0 = [int]$mb.row; $n = [int]$mb.rows
                    $minRow = if ($null -ne $mb.minRow) { [double]$mb.minRow } else { 18 }
                    $pad    = if ($null -ne $mb.pad)    { [double]$mb.pad }    else { 2 }
                    $rowNeed = @()
                    for ($i = 0; $i -lt $n; $i++) {
                        $r = $r0 + $i
                        $h = $minRow
                        foreach ($col in @($mb.perRowCols)) {
                            $v = Need-Cell $ws $m ([string]$col + $r) $true
                            if ($v -gt $h) { $h = $v }
                        }
                        $rowNeed += [math]::Ceiling($h) + $pad
                    }
                    $blockNeed = if ($null -ne $mb.minBlock) { [double]$mb.minBlock } else { 0 }
                    foreach ($col in @($mb.blockCols)) {
                        $v = Need-Cell $ws $m ([string]$col + $r0) $false
                        if ($v -gt $blockNeed) { $blockNeed = $v }
                    }
                    $sum = ($rowNeed | Measure-Object -Sum).Sum
                    if ($sum -lt $blockNeed) {
                        $add = [math]::Ceiling(($blockNeed - $sum) / $n)
                        for ($i = 0; $i -lt $n; $i++) { $rowNeed[$i] = $rowNeed[$i] + $add }
                    }
                    for ($i = 0; $i -lt $n; $i++) { Set-RowHeight $ws ($r0 + $i) ([double]$rowNeed[$i]) }
                    Write-Output ("    block r" + $r0 + " (" + $n + "行) = " + (($rowNeed | ForEach-Object { [int]$_ }) -join ','))
                }

                # ---- 印刷設定
                $ps = $ws.PageSetup
                if ($null -ne $spec.orientation) { $ps.Orientation = [int]$spec.orientation }   # 1=縦 2=横
                if ($null -ne $spec.fitWide -or $null -ne $spec.fitTall) {
                    $ps.Zoom = $false                       # FitToPages を使うには必須
                    if ($null -ne $spec.fitWide) { $ps.FitToPagesWide = $spec.fitWide }
                    # false を渡すと「指定なし（縦は伸ばす）」になる。1 にすると全体が縮小されて潰れる
                    if ($null -ne $spec.fitTall) { $ps.FitToPagesTall = $spec.fitTall }
                }
                elseif ($null -ne $spec.zoom) { $ps.Zoom = [int]$spec.zoom }
                if ($spec.titleRows) { $ps.PrintTitleRows = [string]$spec.titleRows }
                if ($spec.titleCols) { $ps.PrintTitleColumns = [string]$spec.titleCols }
                if ($spec.printArea) { $ps.PrintArea = (Resolve-Addr $ws $spec.printArea) }

                if ($spec.resetPageBreaks) { $ws.ResetAllPageBreaks() }
                foreach ($r in @($spec.pageBreaks)) {
                    if ($null -eq $r) { continue }
                    $ws.HPageBreaks.Add($ws.Rows([int]$r)) | Out-Null    # その行の「前」で改ページ
                }

                Write-Output ("  " + $ws.Name + " : title=" + $ps.PrintTitleRows + " fitW=" + $ps.FitToPagesWide + " fitH=" + $ps.FitToPagesTall + " area=" + $ps.PrintArea)
            }
        }

        $m.Delete()
        try { $bk.Worksheets.Item(1).Activate() | Out-Null } catch {}
        $bk.Save(); $bk.Close($false)
        Write-Output ((Split-Path -Leaf $f.file) + " 更新")
    }
}
finally {
    $xl.ScreenUpdating = $true
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}
