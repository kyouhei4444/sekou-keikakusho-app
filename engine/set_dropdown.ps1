<#
組織表・連絡組織表の「自社人員の氏名欄」を、印刷範囲外に置いた元リストを参照する
ドロップダウン（入力規則）にする。

  powershell -ExecutionPolicy Bypass -File set_dropdown.ps1 -Json 氏名.json

・元リストは印刷範囲外の列に縦に書き、名前定義を OFFSET+COUNTA にするので
  「行を足すだけで選択肢が増える」。
・一覧に無い氏名も手入力できるよう、入力規則のエラー表示はOFFにする。
・names を指定すると元リストを作り直す。appendNames を指定すると
  既存リストの insertBefore（既定：今後追加分）の直前に足す（重複はスキップ）。

JSON（詳細は scripts/README.md）:
{
  "listName": "氏名リスト",
  "header":   "▼氏名リスト（印刷範囲外・行を足すとリストも増えます）",
  "names":    ["山田　太郎", "…", "今後追加分"],
  "appendNames": ["鈴木　一郎"],
  "insertBefore": "今後追加分",
  "files": [
    { "file": "出力\\（3）現場組織表.xls",
      "listSheet": "（3）現場組織表", "listCol": "L", "colWidth": 18,
      "printArea": "$A$1:$J$36",
      "dropdowns": [ { "sheet": "（3）現場組織表", "cells": ["B7","B11","B12"] } ] }
  ]
}
#>
param([Parameter(Mandatory = $true)][string]$Json)

$ErrorActionPreference = 'Stop'
$cfg = Get-Content -Path $Json -Raw -Encoding UTF8 | ConvertFrom-Json

$LIST_NAME = if ($cfg.listName) { [string]$cfg.listName } else { '氏名リスト' }
$HEADER    = if ($cfg.header)   { [string]$cfg.header }   else { '▼氏名リスト（印刷範囲外・行を足すとリストも増えます）' }
$BEFORE    = if ($cfg.insertBefore) { [string]$cfg.insertBefore } else { '今後追加分' }

$xlValidateList   = 3
$xlValidAlertStop = 1
$xlBetween        = 1

function Get-Sheet($bk, $s) {
    if ($s -is [int] -or ([string]$s) -match '^\d+$') { return $bk.Worksheets.Item([int]$s) }
    try { return $bk.Worksheets.Item([string]$s) } catch {}
    foreach ($w in $bk.Worksheets) { if ($w.Name.Trim() -eq ([string]$s).Trim()) { return $w } }
    throw ("シートが見つかりません: '" + $s + "'")
}

# @($null).Count は 1 になるので、空判定はこの関数を通す。
# 戻り値が1要素だと配列がスカラーに戻るため、呼び出し側を必ず @() で囲む
function Get-NonEmpty($v) {
    if ($null -eq $v) { return @() }
    return @($v | Where-Object { $null -ne $_ -and -not [string]::IsNullOrWhiteSpace([string]$_) })
}

function Read-List($ws, [string]$col) {
    $v = @()
    for ($r = 2; $r -le 500; $r++) {
        $t = [string]$ws.Range($col + $r).Value2
        if ([string]::IsNullOrWhiteSpace($t)) { break }
        $v += $t
    }
    return $v
}

function Write-List($ws, [string]$col, [string[]]$names, [double]$width) {
    $ws.Range($col + '1').Value2 = $HEADER
    $ws.Range($col + '1').Font.Bold = $true
    for ($i = 0; $i -lt $names.Count; $i++) { $ws.Range($col + ($i + 2)).Value2 = $names[$i] }
    # 作り直しで短くなった場合に前の行が残らないよう、下を少し消す
    for ($i = $names.Count; $i -lt $names.Count + 20; $i++) {
        $ws.Range($col + ($i + 2)).ClearContents() | Out-Null
    }
    if ($width -gt 0) { $ws.Columns($col).ColumnWidth = $width }
}

function Add-NameRef($bk, $ws, [string]$col) {
    $q = "'" + $ws.Name + "'"
    $f = '=OFFSET(' + $q + '!$' + $col + '$2,0,0,COUNTA(' + $q + '!$' + $col + ':$' + $col + ')-1,1)'
    try { $bk.Names.Item($LIST_NAME).Delete() } catch {}
    $bk.Names.Add($LIST_NAME, $f) | Out-Null
}

function Set-Dropdown($ws, [string[]]$cells) {
    foreach ($a in $cells) {
        $rg = $ws.Range($a)
        try { $rg.Validation.Delete() } catch {}
        $rg.Validation.Add($xlValidateList, $xlValidAlertStop, $xlBetween, '=' + $LIST_NAME) | Out-Null
        $rg.Validation.IgnoreBlank    = $true
        $rg.Validation.InCellDropdown = $true
        $rg.Validation.ShowInput      = $false
        $rg.Validation.ShowError      = $false     # 一覧に無い氏名も手入力できるようにする
    }
}

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
try {
    foreach ($f in @($cfg.files)) {
        if ($null -eq $f) { continue }
        $bk = $xl.Workbooks.Open((Resolve-Path $f.file).Path)
        $lws = Get-Sheet $bk $f.listSheet
        $col = [string]$f.listCol
        $width = if ($null -ne $f.colWidth) { [double]$f.colWidth } else { 18 }

        # ---- 元リスト
        $newNames = @(Get-NonEmpty $cfg.names)
        if ($newNames.Count) {
            Write-List $lws $col $newNames $width
            $cur = $newNames
        }
        else {
            $cur = @(Read-List $lws $col)
            foreach ($n in @(Get-NonEmpty $cfg.appendNames)) {
                if (-not $n) { continue }
                if ($cur -contains $n) { Write-Output ("  すでに登録済み: " + $n); continue }
                $i = [array]::IndexOf($cur, $BEFORE)
                if ($i -lt 0)      { $cur = @($cur) + $n }
                elseif ($i -eq 0)  { $cur = @($n) + @($cur) }
                else               { $cur = @($cur[0..($i - 1)]) + $n + @($cur[$i..($cur.Count - 1)]) }
                Write-Output ("  追加: " + $n)
            }
            Write-List $lws $col @($cur) $width
        }
        if (@($cur).Count -lt 1) {
            throw ('元リストが空です（names も appendNames も無く、' + $col + '2 以下にも氏名がない）。' +
                   'リストが空のまま名前定義を作ると入力規則の設定に失敗する')
        }
        Add-NameRef $bk $lws $col

        # ---- 入力規則
        $cnt = 0
        foreach ($d in @($f.dropdowns)) {
            if ($null -eq $d) { continue }
            $ws = Get-Sheet $bk $d.sheet
            Set-Dropdown $ws @($d.cells)
            $cnt += @($d.cells).Count
        }

        # ---- 印刷範囲（元リストが印刷されないことを確かめる）
        if ($f.printArea) { $lws.PageSetup.PrintArea = [string]$f.printArea }
        $pa = [string]$lws.PageSetup.PrintArea
        if ([string]::IsNullOrWhiteSpace($pa)) {
            Write-Warning ((Split-Path -Leaf $f.file) + " / " + $lws.Name + " : 印刷範囲が未設定。元リスト列($col)が印刷されるおそれがある")
        }
        else {
            try {
                $rg = $lws.Range($pa)
                $lastCol = $rg.Column + $rg.Columns.Count - 1
                if ($lws.Range($col + '1').Column -le $lastCol) {
                    Write-Warning ((Split-Path -Leaf $f.file) + " / " + $lws.Name + " : 元リスト列($col)が印刷範囲 $pa の中にある")
                }
            } catch {}
        }

        $bk.Worksheets.Item(1).Activate() | Out-Null
        $bk.Save(); $bk.Close($false)
        Write-Output ((Split-Path -Leaf $f.file) + " : リスト " + $cur.Count + "件 / ドロップダウン " + $cnt + "箇所")
    }
}
finally {
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}
