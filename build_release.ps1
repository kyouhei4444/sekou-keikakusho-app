<#
配布用の zip を作る（exe のビルド → 必要なフォルダーを並べる → zip）。
GitHub Actions（.github/workflows/release.yml）もこのスクリプトを使う。

  powershell -ExecutionPolicy Bypass -File build_release.ps1 -Version v2026.10.05

  # 別の雛形を同梱する（既定はリポジトリの「施工計画書雛形（山梨県）」）
  powershell -ExecutionPolicy Bypass -File build_release.ps1 -Version v2026.10.05 -Templates "C:\…\施工計画書雛形"

出力：release\sekou-keikakusho-app-<Version>.zip
  施工計画書アプリ\
    施工計画書アプリ.exe
    engine\  data\  web\  projects\  README.md
    施工計画書雛形\        … 既定はリポジトリの「施工計画書雛形（山梨県）」（個人情報を抜いた公開用）

※ このファイルは UTF-8（BOM 付き）で保存すること。BOM が無いと Windows PowerShell 5.1 が日本語を読み違える。
#>
param(
    [string]$Version = 'dev',
    [string]$Templates = '',
    [string]$Python = 'python'
)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $root

$rel = Join-Path $root 'release'
$stage = Join-Path $rel '施工計画書アプリ'
if (Test-Path $rel) { Remove-Item -Recurse -Force $rel }
New-Item -ItemType Directory -Force $stage | Out-Null

Write-Output '■ exe をビルドしています（1分ほど）'
& $Python -m PyInstaller --noconfirm --log-level WARN `
    --distpath (Join-Path $rel 'build') --workpath (Join-Path $rel 'work') `
    sekou-keikakusho-app.spec
if ($LASTEXITCODE -ne 0) { throw "PyInstaller が失敗しました（終了コード $LASTEXITCODE）" }

Write-Output '■ 配布フォルダーを組み立てています'
Copy-Item (Join-Path $rel 'build\sekou-keikakusho-app.exe') (Join-Path $stage '施工計画書アプリ.exe')
foreach ($d in 'engine', 'web') {
    Copy-Item -Recurse (Join-Path $root $d) (Join-Path $stage $d)
}
# data は見本だけ入れる（本物の名簿・会社マスタは各PCで作る）
$dataOut = Join-Path $stage 'data'
New-Item -ItemType Directory -Force $dataOut | Out-Null
Get-ChildItem (Join-Path $root 'data') -File |
    Where-Object { $_.Name -notin @('staff.json', 'company.json') } |
    ForEach-Object { Copy-Item $_.FullName $dataOut }
Copy-Item -Recurse (Join-Path $root 'data\templates') (Join-Path $dataOut 'templates')
New-Item -ItemType Directory -Force (Join-Path $stage 'projects') | Out-Null
Copy-Item (Join-Path $root 'README.md') $stage

# exe の隣に「施工計画書雛形」として置くと、アプリが自動で雛形フォルダーに選ぶ
if (-not $Templates) { $Templates = Join-Path $root '施工計画書雛形（山梨県）' }
if (-not (Test-Path $Templates)) { throw "雛形フォルダーが見つかりません: $Templates" }
Copy-Item -Recurse $Templates (Join-Path $stage '施工計画書雛形')
Write-Output "  雛形を同梱しました: $Templates"

# zip は Python で作る（PowerShell 5.1 の Compress-Archive は日本語のファイル名が化けることがある）
$zip = Join-Path $rel ("sekou-keikakusho-app-$Version.zip")
$py = @"
import os, sys, zipfile
stage, out = sys.argv[1], sys.argv[2]
base = os.path.dirname(stage)
with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
    for r, ds, fs in os.walk(stage):
        for f in fs:
            p = os.path.join(r, f)
            z.write(p, os.path.relpath(p, base))
"@
& $Python -c $py $stage $zip
if ($LASTEXITCODE -ne 0) { throw 'zip の作成に失敗しました' }

$mb = [Math]::Round((Get-Item $zip).Length / 1MB, 1)
Write-Output "■ できました: $zip（$mb MB）"
