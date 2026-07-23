# Release 用の zip を 3 つ作る。
#
# 利用者に渡すのは 1 つの zip だけで完結するようにしている。
# DLL と .bas を別々に配ると必ず組み合わせを間違えるため。
#
# 使い方: pwsh -File scripts\package_release.ps1 -Version 0.1.0

param(
    [Parameter(Mandatory = $true)][string]$Version,
    [string]$Config = 'release'
)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$bin = Join-Path $root "target\$Config"
$out = Join-Path $root "dist"

$tiers = @(
    @{ Id = 'r';     Dll = 'dackdb_r.dll';     Bas = 'DackDbR.bas';     Name = '読み取り専用'; Can = 'SELECT のみ' }
    @{ Id = 'rw';    Dll = 'dackdb_rw.dll';    Bas = 'DackDbRW.bas';    Name = '読み書き可';   Can = 'SELECT / INSERT / UPDATE / DELETE' }
    @{ Id = 'admin'; Dll = 'dackdb_admin.dll'; Bas = 'DackDbAdmin.bas'; Name = '管理者';       Can = 'すべて（DB・テーブルの作成/削除を含む）' }
)

if (Test-Path $out) { Remove-Item $out -Recurse -Force }
New-Item -ItemType Directory -Path $out | Out-Null

$cp932 = [System.Text.Encoding]::GetEncoding(932)

foreach ($t in $tiers) {
    $dllPath = Join-Path $bin $t.Dll
    if (-not (Test-Path $dllPath)) {
        throw "$($t.Dll) がありません。先に cargo build --$Config を実行してください。"
    }

    $stage = Join-Path $out "stage-$($t.Id)"
    New-Item -ItemType Directory -Path (Join-Path $stage 'dll') -Force | Out-Null
    New-Item -ItemType Directory -Path (Join-Path $stage 'vba') -Force | Out-Null

    Copy-Item $dllPath (Join-Path $stage 'dll')
    Copy-Item (Join-Path $root "vba\$($t.Bas)") (Join-Path $stage 'vba')
    Copy-Item (Join-Path $root 'docs\manual.html') (Join-Path $stage '使い方マニュアル.html')
    Copy-Item (Join-Path $root 'LICENSE') $stage

    # 同梱の手順書も VBE 利用者に合わせて CP932 で書く
    $readme = @"
dackdb v$Version  —  $($t.Name)版

このフォルダの中身
------------------
  dll\$($t.Dll)          Excel から呼び出す本体
  vba\$($t.Bas)          VBA モジュール
  使い方マニュアル.html   ブラウザで開いてください
  LICENSE                 MIT License

この版でできること
------------------
  $($t.Can)

導入手順
--------
  1. dll フォルダを、Excel ブックと同じ場所にコピーします。
  2. Excel で Alt + F11 を押して VBE を開き、
     「ファイル」→「ファイルのインポート」から vba\$($t.Bas) を選びます。
  3. VBE の左側で ThisWorkbook をダブルクリックし、次を貼り付けます。

       Private Sub Workbook_Open()
           DackDbInit ThisWorkbook.Path & "\dll"
       End Sub

  4. ブックを保存して開き直せば準備完了です。

  ※ 64bit 版 Excel 専用です。
     「ファイル」→「アカウント」→「Excel のバージョン情報」で確認できます。

詳しい使い方は「使い方マニュアル.html」をご覧ください。
オンライン版: https://km190112.github.io/duckdb-vba-dll/
"@
    [System.IO.File]::WriteAllText(
        (Join-Path $stage 'はじめにお読みください.txt'), $readme, $cp932)

    $zip = Join-Path $out "dackdb-$($t.Id)-v$Version-x64.zip"
    Compress-Archive -Path "$stage\*" -DestinationPath $zip -Force
    Remove-Item $stage -Recurse -Force

    $mb = [math]::Round((Get-Item $zip).Length / 1MB, 1)
    Write-Host ("作成: {0}  ({1} MB)" -f (Split-Path $zip -Leaf), $mb) -ForegroundColor Green
}

Write-Host "`n出力先: $out" -ForegroundColor Cyan
