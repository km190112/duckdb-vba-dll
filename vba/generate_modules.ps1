# DackDb.template.bas から 3 つの VBA モジュールを生成する。
#
# 3 つのモジュールは Lib の DLL 名と説明文だけが違う。手で 3 つ保守すると必ず
# ズレるので 1 つのテンプレートから生成する。
#
# 【文字コードが重要】
# VBE の「ファイルのインポート」は .bas を Shift-JIS (CP932) として読む。
# UTF-8 のまま渡すと日本語コメントがすべて文字化けする。
# テンプレートは UTF-8、生成物は CP932。テンプレートと生成物を分けているのは、
# 生成物を次回の入力にすると再実行で壊れるため。
#
# 使い方: pwsh -File vba\generate_modules.ps1

$ErrorActionPreference = 'Stop'
$here = $PSScriptRoot
$templatePath = Join-Path $here 'DackDb.template.bas'

$template = [System.IO.File]::ReadAllText($templatePath, [System.Text.Encoding]::UTF8)
$cp932 = [System.Text.Encoding]::GetEncoding(932)

$tiers = @(
    @{
        MODULE     = 'DackDbR'
        DLL        = 'dackdb_r.dll'
        TITLE      = '【① 読み取り専用】'
        CANDO      = 'SELECT のみ'
        CANNOTDO   = 'INSERT/UPDATE/DELETE、CREATE/DROP/ALTER、DB 作成'
        WRITE_WARN = "' この DLL（読み取り専用）で呼ぶとエラーになります。`r`n' 使うには DackDbRW モジュール + dackdb_rw.dll に切り替えてください。"
        ADMIN_WARN = "' この DLL（読み取り専用）で呼ぶとエラーになります。`r`n' 使うには DackDbAdmin モジュール + dackdb_admin.dll に切り替えてください。"
    }
    @{
        MODULE     = 'DackDbRW'
        DLL        = 'dackdb_rw.dll'
        TITLE      = '【② 読み書き可】'
        CANDO      = 'SELECT / INSERT / UPDATE / DELETE / トランザクション'
        CANNOTDO   = 'CREATE/DROP/ALTER などのスキーマ変更、DB 作成'
        WRITE_WARN = "' この DLL で使えます。"
        ADMIN_WARN = "' この DLL（読み書き可）で呼ぶとエラーになります。`r`n' 使うには DackDbAdmin モジュール + dackdb_admin.dll に切り替えてください。"
    }
    @{
        MODULE     = 'DackDbAdmin'
        DLL        = 'dackdb_admin.dll'
        TITLE      = '【③ 管理者】'
        CANDO      = 'すべて（DB 作成、テーブル作成/削除、キー設定、スキーマ出力を含む）'
        CANNOTDO   = '（制限なし）'
        WRITE_WARN = "' この DLL で使えます。"
        ADMIN_WARN = "' この DLL で使えます。"
    }
)

$keys = @('MODULE', 'DLL', 'TITLE', 'CANDO', 'CANNOTDO', 'WRITE_WARN', 'ADMIN_WARN')

foreach ($t in $tiers) {
    $out = $template
    foreach ($k in $keys) {
        $out = $out.Replace("{{$k}}", $t[$k])
    }

    if ($out -match '\{\{') {
        throw "未置換のプレースホルダが残っています: $($t.MODULE)"
    }

    $path = Join-Path $here "$($t.MODULE).bas"
    [System.IO.File]::WriteAllText($path, $out, $cp932)
    Write-Host "生成: $($t.MODULE).bas -> $($t.DLL)" -ForegroundColor Green
}

# --- 検証：DLL の取り違えが無いこと -----------------------------------------
$errors = 0
foreach ($t in $tiers) {
    $path = Join-Path $here "$($t.MODULE).bas"
    $text = [System.IO.File]::ReadAllText($path, $cp932)

    # Declare はすべて自分の DLL を指していること
    $expectedDeclares = 18
    $mine = ([regex]::Matches($text, 'Lib "' + [regex]::Escape($t.DLL) + '"')).Count
    if ($mine -ne $expectedDeclares) {
        Write-Host "  [NG] $($t.MODULE): Lib 宣言が $expectedDeclares 個でない (実際 $mine)" -ForegroundColor Red
        $errors++
    }

    # 他 DLL 名は「切り替えてください」の案内文にしか出てこないこと
    foreach ($other in $tiers) {
        if ($other.DLL -eq $t.DLL) { continue }
        $stray = $text -split "`r`n" | Where-Object {
            $_ -like "*$($other.DLL)*" -and $_ -notlike "*切り替えてください*"
        }
        if ($stray) {
            Write-Host "  [NG] $($t.MODULE): $($other.DLL) が混入" -ForegroundColor Red
            $stray | ForEach-Object { Write-Host "       $_" -ForegroundColor Red }
            $errors++
        }
    }

    if ($text -notmatch [regex]::Escape("Attribute VB_Name = ""$($t.MODULE)""")) {
        Write-Host "  [NG] $($t.MODULE): モジュール名が不正" -ForegroundColor Red
        $errors++
    }

    # CP932 で往復できること（VBE が読めるかの確認）
    $bytes = [System.IO.File]::ReadAllBytes($path)
    if ($bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB) {
        Write-Host "  [NG] $($t.MODULE): UTF-8 BOM が付いている（VBE が読めない）" -ForegroundColor Red
        $errors++
    }
    if ($text -notlike '*読み取り*' -and $text -notlike '*使えます*') {
        Write-Host "  [NG] $($t.MODULE): 日本語が壊れている" -ForegroundColor Red
        $errors++
    }
}

if ($errors -eq 0) {
    Write-Host "`n検証 OK: 3 モジュールとも正しい DLL を参照し、CP932 で出力されています" -ForegroundColor Green
    exit 0
} else {
    Write-Host "`n検証 NG: $errors 件" -ForegroundColor Red
    exit 1
}
