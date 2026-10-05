<#
検証用に、シート（またはブック全体）をPDFに書き出す。
書き出したPDFは pdf_png.py で画像にして必ず目視する（文字切れ・図のはみ出し・ページ数は
数値の検算では見つからない）。

  # 1シートだけ
  powershell -ExecutionPolicy Bypass -File to_pdf.ps1 -File "…\（3）現場組織表.xls" -Sheet "（3）現場組織表" -Out <出力先>

  # ブック全体
  powershell -ExecutionPolicy Bypass -File to_pdf.ps1 -File "…\（8）施工管理計画.xls" -Out <出力先>

  # まとめて
  powershell -ExecutionPolicy Bypass -File to_pdf.ps1 -Json 検証.json

JSON:
{
  "out": "…\\scratchpad\\pdf",
  "jobs": [
    { "file": "…\\（3）現場組織表.xls", "sheet": "（3）現場組織表", "name": "p3.pdf" },
    { "file": "…\\（7）施工方法.xls",  "sheet": "*",             "name": "p7" }
  ]
}
#>
param(
    [string]$Json,
    [string]$File,
    [string]$Sheet,
    [string]$Out
)

$ErrorActionPreference = 'Stop'

if ($Json) {
    $cfg = Get-Content -Path $Json -Raw -Encoding UTF8 | ConvertFrom-Json
    $outDir = if ($Out) { $Out } else { [string]$cfg.out }
    $jobs = @($cfg.jobs)
}
else {
    if (-not $File) { throw '-Json または -File を指定してください' }
    $outDir = if ($Out) { $Out } else { Split-Path -Parent (Resolve-Path $File).Path }
    $jobs = @(@{ file = $File; sheet = $Sheet; name = ([System.IO.Path]::GetFileNameWithoutExtension($File)) })
}

if (-not (Test-Path $outDir)) { New-Item -ItemType Directory -Force $outDir | Out-Null }
$outDir = (Resolve-Path $outDir).Path

function Get-Sheet($bk, $s) {
    if ($s -is [int] -or ([string]$s) -match '^\d+$') { return $bk.Worksheets.Item([int]$s) }
    try { return $bk.Worksheets.Item([string]$s) } catch {}
    foreach ($w in $bk.Worksheets) { if ($w.Name.Trim() -eq ([string]$s).Trim()) { return $w } }
    throw ("シートが見つかりません: '" + $s + "'")
}

function Safe-Name([string]$s) {
    foreach ($c in [System.IO.Path]::GetInvalidFileNameChars()) { $s = $s.Replace([string]$c, '_') }
    return $s
}

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
try {
    foreach ($j in $jobs) {
        if ($null -eq $j) { continue }
        $bk = $xl.Workbooks.Open((Resolve-Path $j.file).Path)
        $base = if ($j.name) { [string]$j.name } else { [System.IO.Path]::GetFileNameWithoutExtension([string]$j.file) }
        $base = $base -replace '\.pdf$', ''

        if (-not $j.sheet) {
            # ブック全体（非表示シートは出ない）
            $dst = Join-Path $outDir ($base + '.pdf')
            $bk.ExportAsFixedFormat(0, $dst)      # 0 = xlTypePDF
            Write-Output ((Split-Path -Leaf $dst) + "  <- ブック全体")
        }
        elseif ([string]$j.sheet -eq '*') {
            foreach ($ws in $bk.Worksheets) {
                if ($ws.Visible -ne -1) { continue }
                if ($ws.UsedRange.Cells.Count -le 1 -and [string]::IsNullOrWhiteSpace([string]$ws.Range('A1').Value2)) { continue }
                $dst = Join-Path $outDir ($base + '_' + (Safe-Name $ws.Name) + '.pdf')
                $ws.ExportAsFixedFormat(0, $dst)
                Write-Output ((Split-Path -Leaf $dst) + "  <- " + $ws.Name)
            }
        }
        else {
            $ws = Get-Sheet $bk $j.sheet
            $dst = Join-Path $outDir ($base + '.pdf')
            $ws.ExportAsFixedFormat(0, $dst)
            Write-Output ((Split-Path -Leaf $dst) + "  <- " + $ws.Name)
        }
        $bk.Close($false)
    }
    Write-Output ("出力先: " + $outDir)
}
finally {
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}
