<#
雛形のコピーと、JSON で指定したセル編集を .xls / .xlsx に適用する。
書式（罫線・結合・印刷設定）には触らず、値だけを差し替える。

  powershell -ExecutionPolicy Bypass -File apply_edits.ps1 -Json 編集.json

JSON（詳細は scripts/README.md）:
{
  "copies": [ { "src": "雛形\\（3）現場組織表.xls", "dst": "出力\\（3）現場組織表.xls" } ],
  "edits": [
    { "file":  "出力\\（3）現場組織表.xls",
      "sheet": "（3）現場組織表",          // 数値ならシート番号。前後の空白は無視して照合する
      "sheets": [ { "copyFrom":"...", "name":"...", "after":"..." } ],
      "unmerge": ["A4:O30"],
      "clear":   ["A4:O30"],
      "merge":   ["A4:A6", "F4:F6"],
      "text":    ["B5:K5"],                 // 先に書式を文字列にする（'2026年10月' の日付化を防ぐ）
      "cells":   { "B7": "山田　太郎", "A22": null, "F13": "090-0000-0000" },
      "font":    [ { "range":"B4:C40", "color":0, "bold":false } ],
      "rowHeight": { "2": 42 }
    }
  ]
}
#>
param([Parameter(Mandatory = $true)][string]$Json)

$ErrorActionPreference = 'Stop'
$cfg = Get-Content -Path $Json -Raw -Encoding UTF8 | ConvertFrom-Json

function Get-Sheet($bk, $s) {
    if ($s -is [int] -or ([string]$s) -match '^\d+$') { return $bk.Worksheets.Item([int]$s) }
    try { return $bk.Worksheets.Item([string]$s) } catch {}
    # 雛形はシート名の前後に空白が入っていることがあるため trim で照合する
    foreach ($w in $bk.Worksheets) { if ($w.Name.Trim() -eq ([string]$s).Trim()) { return $w } }
    throw ("シートが見つかりません: '" + $s + "'")
}

# PowerShell は COM プロパティの型を最初の代入で覚えてしまい、次に別の型を入れると
# 「型 'System.Int32' のオブジェクトを型 'System.String' にキャストできません」で落ちる。
# 遅延バインドで呼べば型を問わず入れられる。
function Set-Value($rg, $v) {
    [void]$rg.GetType().InvokeMember('Value2', 'SetProperty', $null, $rg, @($v))
}

function Clear-Any($ws, [string]$addr) {
    try { $ws.Range($addr).ClearContents() | Out-Null }
    catch {
        # 結合セルを含むと ClearContents が失敗するのでセル単位に落とす
        foreach ($c in $ws.Range($addr).Cells) {
            try { $c.MergeArea.ClearContents() | Out-Null } catch {}
        }
    }
}

# @($null).Count は 1 になるので、null を除いた配列にしてから扱う。
# 戻り値が1要素だと配列がスカラーに戻るため、呼び出し側を必ず @() で囲む
function Get-NonNull($v) {
    if ($null -eq $v) { return @() }
    return @($v | Where-Object { $null -ne $_ })
}

# ---------------------------------------------------------------- 雛形のコピー
foreach ($c in @(Get-NonNull $cfg.copies)) {
    if ($null -eq $c) { continue }
    if (-not (Test-Path $c.src)) { throw ("雛形が見つかりません: " + $c.src) }
    $d = Split-Path -Parent $c.dst
    if ($d -and -not (Test-Path $d)) { New-Item -ItemType Directory -Force $d | Out-Null }
    Copy-Item -Path $c.src -Destination $c.dst -Force
    Write-Output ("copy : " + (Split-Path -Leaf $c.dst))
}

# ---------------------------------------------------------------- セル編集
$editList = @(Get-NonNull $cfg.edits)
if (-not $editList.Count) { return }

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.ScreenUpdating = $false
try {
    $open = @{}
    foreach ($e in $editList) {
        if (-not $open.ContainsKey($e.file)) {
            $open[$e.file] = $xl.Workbooks.Open((Resolve-Path $e.file).Path)
        }
        $bk = $open[$e.file]

        # シートの追加・複製・削除・改名
        foreach ($op in @($e.sheets)) {
            if ($null -eq $op) { continue }
            if ($op.delete) { (Get-Sheet $bk $op.delete).Delete(); continue }
            if ($op.rename) { (Get-Sheet $bk $op.rename.from).Name = $op.rename.to; continue }
            if ($op.copyFrom) {
                $src = Get-Sheet $bk $op.copyFrom
                $after = if ($op.after) { Get-Sheet $bk $op.after } else { $src }
                $src.Copy([System.Reflection.Missing]::Value, $after) | Out-Null
                $new = $bk.Worksheets.Item($after.Index + 1)
                if ($op.name) { $new.Name = $op.name }
                Write-Output ("sheet: + " + $new.Name)
            }
        }

        if (-not $e.sheet) { continue }
        $ws = Get-Sheet $bk $e.sheet

        foreach ($a in @($e.unmerge)) { if ($a) { $ws.Range($a).UnMerge() } }
        foreach ($a in @($e.clear))   { if ($a) { Clear-Any $ws $a } }
        foreach ($a in @($e.merge))   { if ($a) { $ws.Range($a).Merge() | Out-Null } }
        foreach ($a in @($e.text))    { if ($a) { $ws.Range($a).NumberFormat = '@' } }

        if ($e.cells) {
            foreach ($p in $e.cells.PSObject.Properties) {
                $cell = $ws.Range($p.Name)
                if ($null -eq $p.Value) { $cell.MergeArea.ClearContents() | Out-Null }
                else { Set-Value $cell $p.Value }     # 数値は数値、文字列は文字列のまま入れる
            }
        }

        foreach ($f in @($e.font)) {
            if ($null -eq $f) { continue }
            $rg = $ws.Range($f.range)
            if ($null -ne $f.color) { $rg.Font.Color = [int]$f.color }
            if ($null -ne $f.bold)  { $rg.Font.Bold = [bool]$f.bold }
            if ($null -ne $f.size)  { $rg.Font.Size = [double]$f.size }
        }

        if ($e.rowHeight) {
            foreach ($p in $e.rowHeight.PSObject.Properties) {
                $h = [double]$p.Value
                if ($h -gt 409) { $h = 409 }          # Excel の行高上限
                $ws.Rows([int]$p.Name).RowHeight = $h
            }
        }

        Write-Output ("edit : " + (Split-Path -Leaf $e.file) + " / " + $ws.Name)
    }

    foreach ($k in @($open.Keys)) {
        $open[$k].Worksheets.Item(1).Activate() | Out-Null
        $open[$k].Save()
        $open[$k].Close($false)
    }
}
finally {
    $xl.ScreenUpdating = $true
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}
