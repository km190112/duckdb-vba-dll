# 社内固有情報がリポジトリに混入していないか検査する。
#
# 背景: 2026-08-08 に別リポジトリで、社内サーバ名と機密区分の表記を含む
# UNC パスが約 3.5 か月 public に出ていた事故があった。人の目視に頼らず
# 機械的に止めるためのもの。
#
# 検出したい語そのものは、このファイルのどこにも (コメントを含めて) 書かない。
# 書けば自分自身が一致する。
#
# 設計上の約束:
#
#  1. パターンに実際のサーバ名や個人名を書かない。
#     書いた時点でこのファイルがその名前を持つことになり、検出しようと
#     しているものを自分で作り出す。秘密ではなく「構造的な特徴」で拾う。
#
#  2. 日本語パターンは文字コードから組み立てる。
#     リテラルで書くとこのファイル自身が一致し、自分を走査対象から外す
#     必要が出る。除外すればそこが恒久的な検査の穴になる。
#     コードから組み立てれば自分自身も検査できる。
#
#  3. 読めなかったファイルは黙って飛ばさない。失敗として扱う。
#     「読めないので検査していない」を「問題なし」と報告してはいけない。
#
#  4. 対象は git の追跡ファイル全件。拡張子の allowlist にしない。
#     未知の拡張子に置かれた漏洩を見逃すため。バイナリは明示的に除外し、
#     除外したことを必ず出力に残す。

[CmdletBinding()]
param(
    # バイナリとして扱い、内容検査の対象外にする拡張子
    [string[]] $BinaryExtensions = @(
        '.zip', '.dll', '.exe', '.pdb', '.lib', '.png', '.jpg', '.jpeg',
        '.gif', '.ico', '.pdf', '.xlsx', '.xlsm', '.docx', '.7z', '.ttf', '.otf'
    )
)

$ErrorActionPreference = 'Stop'

function ConvertFrom-HexCode {
    # 'XXXX,YYYY' 形式の UTF-16 コードポイント列から文字列を作る
    param([string] $Codes)
    -join ($Codes -split ',' | ForEach-Object { [char][Convert]::ToInt32($_.Trim(), 16) })
}

# 検出したい語は「コメントにも」書かない。コメントも本文の一部として
# 走査されるため、書いた瞬間にこのファイル自身が一致する。
# 何の語かは下の Name で分かるようにしてある。
$markJa1 = ConvertFrom-HexCode '793E,5916,79D8'
$markJa2 = ConvertFrom-HexCode '90E8,5916,79D8'
$markEn = ConvertFrom-HexCode '43,4F,4E,46,49,44,45,4E,54,49,41,4C'
$sharedFolder = ConvertFrom-HexCode '500B,4EBA,30D5,30A9,30EB,30C0'

$patterns = @(
    @{ Name = 'ConfidentialMarkJa1'; Pattern = [regex]::Escape($markJa1) },
    @{ Name = 'ConfidentialMarkJa2'; Pattern = [regex]::Escape($markJa2) },
    @{ Name = 'ConfidentialMarkEn'; Pattern = [regex]::Escape($markEn) },
    @{ Name = 'SharedPersonalFolder'; Pattern = [regex]::Escape($sharedFolder) },
    # ホスト名が3桁数字で終わる UNC。社内サーバの命名によくある形
    @{ Name = 'UncHostWithDigits'; Pattern = '\\\\[A-Za-z0-9_-]*\d{3}\\' },
    # 開発機の絶対パス
    @{ Name = 'LocalUserPath'; Pattern = '[A-Za-z]:\\Users\\[A-Za-z0-9_]+' }
)

# 追跡ファイルのみを対象にする。未追跡はそもそも push されない。
$tracked = & git ls-files
if ($LASTEXITCODE -ne 0) { throw 'git ls-files に失敗しました' }

$utf8 = [System.Text.Encoding]::UTF8
try { $cp932 = [System.Text.Encoding]::GetEncoding(932) } catch { $cp932 = $null }

$scanned = 0
$skipped = @()
$findings = @()
$unreadable = @()

foreach ($rel in $tracked) {
    if (-not (Test-Path -LiteralPath $rel)) { continue }
    $ext = [System.IO.Path]::GetExtension($rel).ToLowerInvariant()
    if ($BinaryExtensions -contains $ext) {
        $skipped += $rel
        continue
    }

    try {
        $bytes = [System.IO.File]::ReadAllBytes((Resolve-Path -LiteralPath $rel))
    } catch {
        # 黙って飛ばさない。読めないなら検査できていないので失敗にする。
        $unreadable += "$rel : $($_.Exception.Message)"
        continue
    }

    # 保存時の文字コードが分からないので複数の解釈で照合する。
    # UTF-8 だけで読むと CP932 で保存された日本語を取りこぼす。
    $views = @($utf8.GetString($bytes))
    if ($null -ne $cp932) { $views += $cp932.GetString($bytes) }

    foreach ($p in $patterns) {
        foreach ($text in $views) {
            $m = [regex]::Match($text, $p.Pattern)
            if ($m.Success) {
                $line = ($text.Substring(0, $m.Index) -split "`n").Count
                $findings += [pscustomobject]@{
                    File = $rel; Line = $line; Rule = $p.Name; Match = $m.Value
                }
                break
            }
        }
    }
    $scanned++
}

Write-Output "検査したファイル: $scanned 件"
if ($skipped.Count -gt 0) {
    # 何を見ていないかを必ず出す。黙って減らすと「全部見た」ように読める。
    Write-Output "バイナリとして内容検査を省いたファイル: $($skipped.Count) 件"
    $skipped | ForEach-Object { Write-Output "    $_" }
}

$failed = $false

if ($unreadable.Count -gt 0) {
    Write-Output "::error::読み取れなかったファイルがあります。検査できていないため失敗として扱います"
    $unreadable | ForEach-Object { Write-Output "::error::$_" }
    $failed = $true
}

if ($findings.Count -gt 0) {
    foreach ($f in $findings) {
        Write-Output "::error file=$($f.File),line=$($f.Line)::社内固有情報の疑い [$($f.Rule)] 該当: $($f.Match)"
    }
    Write-Output "検出: $($findings.Count) 件"
    $failed = $true
}

if ($failed) { exit 1 }

Write-Output "検出なし"
exit 0
