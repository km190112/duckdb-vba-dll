# VBA 境界の自動テスト。
#
# .NET の `[MarshalAs(UnmanagedType.Struct)] ref object` は VARIANT* として
# マーシャリングされる。これは VBA の `ByRef result As Variant` と**まったく同じ
# ABI** なので、Excel を起動せずに VBA 境界そのものを検証できる。
# 入力文字列も Marshal.StringToHGlobalUni（NUL 終端 UTF-16）＝ VBA の StrPtr() と同じ。
#
# 使い方:  pwsh -File tests\vba_boundary_test.ps1
#          pwsh -File tests\vba_boundary_test.ps1 -Config debug

param(
    [ValidateSet('debug', 'release')]
    [string]$Config = 'release'
)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$binDir = Join-Path $root "target\$Config"

$script:passed = 0
$script:failed = 0

function Assert-That {
    param([bool]$Condition, [string]$Name, [string]$Detail = '')
    if ($Condition) {
        $script:passed++
        Write-Host "  [OK]   $Name" -ForegroundColor Green
    } else {
        $script:failed++
        Write-Host "  [FAIL] $Name" -ForegroundColor Red
        if ($Detail) { Write-Host "         $Detail" -ForegroundColor Red }
    }
}

function Assert-Equal {
    param($Expected, $Actual, [string]$Name)
    Assert-That ($Expected -eq $Actual) $Name "期待値: [$Expected]  実際: [$Actual]"
}

# --- 3 つの DLL それぞれに P/Invoke クラスを定義する -------------------------
# 同じ関数名を持つ 3 つの DLL を同一プロセスで扱うため、クラスを分ける。

$tiers = @(
    @{ Name = 'R';     Dll = 'dackdb_r.dll';     Class = 'DackR' }
    @{ Name = 'RW';    Dll = 'dackdb_rw.dll';    Class = 'DackRW' }
    @{ Name = 'ADMIN'; Dll = 'dackdb_admin.dll'; Class = 'DackAdmin' }
)

foreach ($t in $tiers) {
    $path = (Join-Path $binDir $t.Dll) -replace '\\', '\\'
    $src = @"
using System;
using System.Runtime.InteropServices;
public static class $($t.Class) {
    const string L = "$path";
    // VBA:  Declare PtrSafe Function X Lib "..." (..., ByRef r As Variant) As Long
    [DllImport(L)] public static extern int DackVersion([MarshalAs(UnmanagedType.Struct)] ref object r);
    [DllImport(L)] public static extern int DackCapabilities([MarshalAs(UnmanagedType.Struct)] ref object r);
    [DllImport(L)] public static extern int DackOpen(IntPtr path, [MarshalAs(UnmanagedType.Struct)] ref object r);
    [DllImport(L)] public static extern int DackClose(long h, [MarshalAs(UnmanagedType.Struct)] ref object r);
    [DllImport(L)] public static extern int DackQuery(long h, IntPtr sql, [MarshalAs(UnmanagedType.Struct)] ref object r);
    [DllImport(L)] public static extern int DackListTables(long h, [MarshalAs(UnmanagedType.Struct)] ref object r);
    [DllImport(L)] public static extern int DackDescribe(long h, IntPtr t, [MarshalAs(UnmanagedType.Struct)] ref object r);
    [DllImport(L)] public static extern int DackExecute(long h, IntPtr sql, [MarshalAs(UnmanagedType.Struct)] ref object r);
    // VBA: ByRef params As Variant / ByRef data As Variant も VARIANT* として渡る
    [DllImport(L)] public static extern int DackQueryParams(long h, IntPtr sql, [MarshalAs(UnmanagedType.Struct)] ref object p, [MarshalAs(UnmanagedType.Struct)] ref object r);
    [DllImport(L)] public static extern int DackExecuteParams(long h, IntPtr sql, [MarshalAs(UnmanagedType.Struct)] ref object p, [MarshalAs(UnmanagedType.Struct)] ref object r);
    [DllImport(L)] public static extern int DackAppendArray(long h, IntPtr t, [MarshalAs(UnmanagedType.Struct)] ref object d, [MarshalAs(UnmanagedType.Struct)] ref object r);
    [DllImport(L)] public static extern int DackBegin(long h, [MarshalAs(UnmanagedType.Struct)] ref object r);
    [DllImport(L)] public static extern int DackCommit(long h, [MarshalAs(UnmanagedType.Struct)] ref object r);
    [DllImport(L)] public static extern int DackRollback(long h, [MarshalAs(UnmanagedType.Struct)] ref object r);
    [DllImport(L)] public static extern int DackCreateDatabase(IntPtr path, [MarshalAs(UnmanagedType.Struct)] ref object r);
    [DllImport(L)] public static extern int DackExecuteDDL(long h, IntPtr sql, [MarshalAs(UnmanagedType.Struct)] ref object r);
    [DllImport(L)] public static extern int DackExportSchema(long h, IntPtr f, [MarshalAs(UnmanagedType.Struct)] ref object r);
    [DllImport(L)] public static extern int DackCheckpoint(long h, [MarshalAs(UnmanagedType.Struct)] ref object r);
}
"@
    Add-Type -TypeDefinition $src
}

# VBA の StrPtr() 相当（NUL 終端 UTF-16 へのポインタ）
function Str-Ptr([string]$s) { [Runtime.InteropServices.Marshal]::StringToHGlobalUni($s) }
function Free-Ptr([IntPtr]$p) { [Runtime.InteropServices.Marshal]::FreeHGlobal($p) }

$dbPath = Join-Path $env:TEMP "dackdb-vba-test.db"
if (Test-Path $dbPath) { Remove-Item $dbPath -Force }

$JP = "テスト漢字" + [char]::ConvertFromUtf32(0x20BB7) + [char]::ConvertFromUtf32(0x1F600)

Write-Host "`n=== 1. バージョンと権限レベルの申告 ===" -ForegroundColor Cyan
$r = $null
Assert-Equal 0 ([DackR]::DackVersion([ref]$r)) "R: DackVersion が成功"
Assert-That ($r -like '*READ_ONLY*') "R: 権限レベルを READ_ONLY と申告" "実際: $r"
Write-Host "         $r" -ForegroundColor DarkGray

Assert-Equal 0 ([DackRW]::DackVersion([ref]$r)) "RW: DackVersion が成功"
Assert-That ($r -like '*READ_WRITE*') "RW: 権限レベルを READ_WRITE と申告" "実際: $r"
Write-Host "         $r" -ForegroundColor DarkGray

Assert-Equal 0 ([DackAdmin]::DackVersion([ref]$r)) "ADMIN: DackVersion が成功"
Assert-That ($r -like '*ADMIN*') "ADMIN: 権限レベルを ADMIN と申告" "実際: $r"
Write-Host "         $r" -ForegroundColor DarkGray

Assert-Equal 0 ([DackR]::DackCapabilities([ref]$r)) "DackCapabilities が成功"
Assert-That ($r -notlike '*DackExecuteDDL*') "読み取り DLL の一覧に DDL が無い" "実際: $r"
Assert-Equal 0 ([DackAdmin]::DackCapabilities([ref]$r)) "ADMIN: DackCapabilities"
Assert-That ($r -like '*DackExecuteDDL*') "管理者 DLL の一覧に DDL がある" "実際: $r"

Write-Host "`n=== 2. 管理者 DLL で DB とテーブルを作成（日本語識別子）===" -ForegroundColor Cyan
$r = $null
$p = Str-Ptr $dbPath
$rc = [DackAdmin]::DackCreateDatabase($p, [ref]$r)
Free-Ptr $p
Assert-Equal 0 $rc "DackCreateDatabase が成功"
Assert-That ($r -is [long]) "ハンドルが LongLong (VT_I8)" "実際: $($r.GetType())"
$hAdmin = [long]$r

$ddl = "CREATE TABLE 売上 (id INTEGER PRIMARY KEY, 部署 VARCHAR, 金額 BIGINT, 単価 DECIMAL(18,2), 日付 DATE, 有効 BOOLEAN)"
$p = Str-Ptr $ddl
$rc = [DackAdmin]::DackExecuteDDL($hAdmin, $p, [ref]$r)
Free-Ptr $p
Assert-Equal 0 $rc "DackExecuteDDL で CREATE TABLE"

$ins = "INSERT INTO 売上 VALUES (1,'$JP',1000,1234.56,DATE '2024-01-15',true),(2,NULL,NULL,NULL,NULL,NULL)"
$p = Str-Ptr $ins
$rc = [DackAdmin]::DackExecute($hAdmin, $p, [ref]$r)
Free-Ptr $p
Assert-Equal 0 $rc "DackExecute で INSERT"
Assert-Equal 2 $r "影響行数が 2"

Write-Host "`n=== 3. 2 次元 Variant 配列（VBA の Range.Value と同じ形）===" -ForegroundColor Cyan
$p = Str-Ptr "SELECT * FROM 売上 ORDER BY id"
$rc = [DackAdmin]::DackQuery($hAdmin, $p, [ref]$r)
Free-Ptr $p
Assert-Equal 0 $rc "DackQuery が成功"
Assert-That ($r -is [object[,]]) "2 次元配列が返る" "実際: $($r.GetType())"

$arr = [object[,]]$r
Assert-Equal 1 $arr.GetLowerBound(0) "行の下限が 1（VBA の既定と一致）"
Assert-Equal 1 $arr.GetLowerBound(1) "列の下限が 1"
Assert-Equal 3 $arr.GetUpperBound(0) "ヘッダ 1 行 + データ 2 行"
Assert-Equal 6 $arr.GetUpperBound(1) "6 列"

Assert-Equal "id"   $arr[1,1] "ヘッダ 1 列目"
Assert-Equal "部署" $arr[1,2] "ヘッダに日本語列名"
Assert-Equal "金額" $arr[1,3] "ヘッダに日本語列名 (2)"

Write-Host "`n=== 4. 日本語・サロゲートペア・絵文字の往復（ODBC 経路が壊す部分）===" -ForegroundColor Cyan
Assert-Equal $JP $arr[2,2] "漢字＋サロゲートペア＋絵文字が完全一致"
Write-Host "         期待: $JP" -ForegroundColor DarkGray
Write-Host "         実際: $($arr[2,2])" -ForegroundColor DarkGray

Write-Host "`n=== 5. 型マッピング ===" -ForegroundColor Cyan
Assert-That ($arr[2,1] -is [int])     "INTEGER が Long (VT_I4)"    "実際: $($arr[2,1].GetType())"
Assert-That ($arr[2,3] -is [long])    "BIGINT が LongLong (VT_I8)" "実際: $($arr[2,3].GetType())"
Assert-That ($arr[2,4] -is [double])  "DECIMAL(18,2) が Double"    "実際: $($arr[2,4].GetType())"
Assert-Equal 1234.56 $arr[2,4] "DECIMAL の値"
Assert-That ($arr[2,5] -is [datetime]) "DATE が Date (VT_DATE)"    "実際: $($arr[2,5].GetType())"
Assert-Equal '2024-01-15' $arr[2,5].ToString('yyyy-MM-dd') "DATE の値"
Assert-That ($arr[2,6] -is [bool])    "BOOLEAN が Boolean"         "実際: $($arr[2,6].GetType())"
Assert-Equal $true $arr[2,6] "BOOLEAN の値"

Write-Host "`n=== 6. NULL の扱い（VBA の IsNull / 空セル）===" -ForegroundColor Cyan
Assert-That ($arr[3,2] -is [DBNull]) "VARCHAR の NULL が DBNull (VT_NULL)" "実際: $($arr[3,2])"
Assert-That ($arr[3,3] -is [DBNull]) "BIGINT の NULL が DBNull"
Assert-That ($arr[3,5] -is [DBNull]) "DATE の NULL が DBNull"

Write-Host "`n=== 7. 権限階層 ===" -ForegroundColor Cyan
$rc = [DackAdmin]::DackClose($hAdmin, [ref]$r)
Assert-Equal 0 $rc "管理者接続を閉じる"

# ① 読み取り専用
$p = Str-Ptr $dbPath
$rc = [DackR]::DackOpen($p, [ref]$r)
Free-Ptr $p
Assert-Equal 0 $rc "読み取り DLL で DackOpen"
$hR = [long]$r

$p = Str-Ptr "SELECT count(*) FROM 売上"
$rc = [DackR]::DackQuery($hR, $p, [ref]$r)
Free-Ptr $p
Assert-Equal 0 $rc "読み取り DLL で SELECT は成功"

$p = Str-Ptr "INSERT INTO 売上 VALUES (9,'x',1,1,NULL,true)"
$rc = [DackR]::DackExecute($hR, $p, [ref]$r)
Free-Ptr $p
Assert-Equal -403 $rc "読み取り DLL の DackExecute が DACK_E_FORBIDDEN"
Assert-That ($r -like '*dackdb_rw.dll*') "上位 DLL への案内が入っている" "実際: $r"

$p = Str-Ptr "DROP TABLE 売上"
$rc = [DackR]::DackExecuteDDL($hR, $p, [ref]$r)
Free-Ptr $p
Assert-Equal -403 $rc "読み取り DLL の DackExecuteDDL が拒否される"

# SQL 経由のバイパス試験
$p = Str-Ptr "SELECT 1; DROP TABLE 売上;"
$rc = [DackR]::DackQuery($hR, $p, [ref]$r)
Free-Ptr $p
Assert-That ($rc -ne 0) "複数文に紛れた DROP が拒否される" "rc=$rc"

[void][DackR]::DackClose($hR, [ref]$r)

# ② 読み書き可
$p = Str-Ptr $dbPath
$rc = [DackRW]::DackOpen($p, [ref]$r)
Free-Ptr $p
Assert-Equal 0 $rc "読み書き DLL で DackOpen"
$hRW = [long]$r

$p = Str-Ptr "INSERT INTO 売上 VALUES (3,'経理部',3000,1,DATE '2024-03-01',false)"
$rc = [DackRW]::DackExecute($hRW, $p, [ref]$r)
Free-Ptr $p
Assert-Equal 0 $rc "読み書き DLL で INSERT が成功"
Assert-Equal 1 $r "1 行挿入された"

$p = Str-Ptr "CREATE TABLE t2 (a INTEGER)"
$rc = [DackRW]::DackExecuteDDL($hRW, $p, [ref]$r)
Free-Ptr $p
Assert-Equal -403 $rc "読み書き DLL の DackExecuteDDL が拒否される"
Assert-That ($r -like '*dackdb_admin.dll*') "管理者 DLL への案内が入っている" "実際: $r"

[void][DackRW]::DackClose($hRW, [ref]$r)

Write-Host "`n=== 8. 異常系（Excel を落とさないこと）===" -ForegroundColor Cyan
foreach ($bad in @(0L, -1L, [long]::MaxValue, 999999L)) {
    $rc = [DackR]::DackQuery($bad, (Str-Ptr "SELECT 1"), [ref]$r)
    Assert-That ($rc -ne 0) "不正ハンドル $bad がエラーになる（クラッシュしない）"
}
$rc = [DackR]::DackQuery(1L, [IntPtr]::Zero, [ref]$r)
Assert-That ($rc -ne 0) "null の SQL ポインタがエラーになる"

Write-Host "`n=== 9. 性能（10 万行 x 6 列）===" -ForegroundColor Cyan
$p = Str-Ptr $dbPath
[void][DackAdmin]::DackOpen($p, [ref]$r); Free-Ptr $p
$hPerf = [long]$r
$sql = "SELECT i AS id, '部署' || (i%10) AS 部署, i*100 AS 金額, i*1.5 AS 単価, DATE '2024-01-01' + INTERVAL (i%365) DAY AS 日付, (i%2=0) AS 有効 FROM range(100000) t(i)"
$sw = [Diagnostics.Stopwatch]::StartNew()
$p = Str-Ptr $sql
$rc = [DackAdmin]::DackQuery($hPerf, $p, [ref]$r)
Free-Ptr $p
$sw.Stop()
Assert-Equal 0 $rc "10 万行の SELECT が成功"
$arr2 = [object[,]]$r
Assert-Equal 100001 $arr2.GetUpperBound(0) "10 万行 + ヘッダ"
Assert-That ($sw.ElapsedMilliseconds -lt 5000) "5 秒未満で完了" "実際: $($sw.ElapsedMilliseconds) ms"
Write-Host "         $($sw.ElapsedMilliseconds) ms" -ForegroundColor DarkGray

Write-Host "`n=== 10. メモリリーク（1000 回クエリ）===" -ForegroundColor Cyan
[GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect()
$before = [Diagnostics.Process]::GetCurrentProcess().PrivateMemorySize64
for ($i = 0; $i -lt 1000; $i++) {
    $p = Str-Ptr "SELECT * FROM 売上"
    [void][DackAdmin]::DackQuery($hPerf, $p, [ref]$r)
    Free-Ptr $p
    $r = $null
}
[GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect()
$after = [Diagnostics.Process]::GetCurrentProcess().PrivateMemorySize64
$growthMB = [math]::Round(($after - $before) / 1MB, 1)
Assert-That ($growthMB -lt 50) "1000 回クエリでメモリ増加が 50MB 未満" "実際: +$growthMB MB"
Write-Host "         +$growthMB MB" -ForegroundColor DarkGray

Write-Host "`n=== 11. パラメータバインド（SQL インジェクション対策）===" -ForegroundColor Cyan
# VBA の Array(...) 相当（1 次元 Variant 配列、下限 0）
$prm = [object[]]@($JP)
$p = Str-Ptr "SELECT id FROM 売上 WHERE 部署 = ?"
$rc = [DackAdmin]::DackQueryParams($hPerf, $p, [ref]$prm, [ref]$r)
Free-Ptr $p
Assert-Equal 0 $rc "DackQueryParams が成功"
$pa = [object[,]]$r
Assert-Equal 2 $pa.GetUpperBound(0) "ヘッダ + 1 件ヒット"
Assert-Equal 1 $pa[2,1] "日本語パラメータでヒットした"

# インジェクションを試みるパラメータ
$prm = [object[]]@("' OR 1=1; DROP TABLE 売上; --")
$p = Str-Ptr "SELECT id FROM 売上 WHERE 部署 = ?"
$rc = [DackAdmin]::DackQueryParams($hPerf, $p, [ref]$prm, [ref]$r)
Free-Ptr $p
Assert-Equal 0 $rc "インジェクション文字列でもクエリ自体は成功"
Assert-Equal 1 ([object[,]]$r).GetUpperBound(0) "ヒット 0 件（値として扱われた）"

$p = Str-Ptr "SELECT count(*) FROM 売上"
$rc = [DackAdmin]::DackQuery($hPerf, $p, [ref]$r)
Free-Ptr $p
Assert-Equal 0 $rc "テーブルが DROP されていない"

# 個数不一致
$prm = [object[]]@(1, 2)
$p = Str-Ptr "SELECT * FROM 売上 WHERE id = ?"
$rc = [DackAdmin]::DackQueryParams($hPerf, $p, [ref]$prm, [ref]$r)
Free-Ptr $p
Assert-That ($rc -ne 0) "パラメータ個数の不一致がエラーになる"

# 読み取り DLL で ExecuteParams。
# 権限チェックはハンドルを引く前に行われるので、無効ハンドルでも -403 になる。
# （DuckDB は同一プロセスで書き込み中のファイルを別接続から開けないため、
#   ここで読み取り接続を開くことはできない。この制約は README に記載。）
$prm = [object[]]@(1)
$p = Str-Ptr "DELETE FROM 売上 WHERE id = ?"
$rc = [DackR]::DackExecuteParams(0L, $p, [ref]$prm, [ref]$r)
Free-Ptr $p
Assert-Equal -403 $rc "読み取り DLL の DackExecuteParams が拒否される"
Assert-That ($r -like '*dackdb_rw.dll*') "上位 DLL への案内が入っている" "実際: $r"

Write-Host "`n=== 12. 一括投入（Appender）===" -ForegroundColor Cyan
$p = Str-Ptr "CREATE TABLE 投入先 (id INTEGER, 名前 VARCHAR, 金額 BIGINT)"
[void][DackAdmin]::DackExecuteDDL($hPerf, $p, [ref]$r); Free-Ptr $p

# Excel の Range.Value 相当（2 次元、下限 1）
$rows = 20000
$data = [object[,]]::new($rows, 3)
# .NET の既定は下限 0。Excel と同じ下限 1 の配列を作って渡す
$data = [Array]::CreateInstance([object], @($rows, 3), @(1, 1))
for ($i = 1; $i -le $rows; $i++) {
    $data.SetValue([int]$i, $i, 1)
    $data.SetValue("部署$JP", $i, 2)
    $data.SetValue([long]($i * 100), $i, 3)
}
$boxed = [object]$data
$sw = [Diagnostics.Stopwatch]::StartNew()
$p = Str-Ptr "投入先"
$rc = [DackAdmin]::DackAppendArray($hPerf, $p, [ref]$boxed, [ref]$r)
Free-Ptr $p
$sw.Stop()
Assert-Equal 0 $rc "DackAppendArray が成功"
Assert-Equal $rows $r "$rows 行投入された"
Assert-That ($sw.ElapsedMilliseconds -lt 5000) "$rows 行の投入が 5 秒未満" "実際: $($sw.ElapsedMilliseconds) ms"
Write-Host "         $($sw.ElapsedMilliseconds) ms" -ForegroundColor DarkGray

$p = Str-Ptr "SELECT count(*) AS n, max(名前) AS nm FROM 投入先"
$rc = [DackAdmin]::DackQuery($hPerf, $p, [ref]$r)
Free-Ptr $p
$chk = [object[,]]$r
Assert-Equal $rows $chk[2,1] "テーブル上の行数が一致"
Assert-Equal "部署$JP" $chk[2,2] "日本語が壊れずに投入された"

# 列数不一致
$bad = [Array]::CreateInstance([object], @(1, 2), @(1, 1))
$bad.SetValue([int]1, 1, 1); $bad.SetValue("x", 1, 2)
$boxedBad = [object]$bad
$p = Str-Ptr "投入先"
$rc = [DackAdmin]::DackAppendArray($hPerf, $p, [ref]$boxedBad, [ref]$r)
Free-Ptr $p
Assert-That ($rc -ne 0) "列数不一致がエラーになる"
Assert-That ($r -like '*3 列*') "テーブルの列数が案内される" "実際: $r"

# 読み取り DLL では拒否（権限チェックはハンドルより先。上のコメント参照）
$p = Str-Ptr "投入先"
$rc = [DackR]::DackAppendArray(0L, $p, [ref]$boxed, [ref]$r)
Free-Ptr $p
Assert-Equal -403 $rc "読み取り DLL の DackAppendArray が拒否される"
Assert-That ($r -like '*dackdb_rw.dll*') "上位 DLL への案内が入っている" "実際: $r"

Write-Host "`n=== 13. スキーマ出力（管理者のみ）===" -ForegroundColor Cyan
$p = Str-Ptr "ddl"
$rc = [DackAdmin]::DackExportSchema($hPerf, $p, [ref]$r)
Free-Ptr $p
Assert-Equal 0 $rc "DackExportSchema(ddl) が成功"
Assert-That ($r -like '*CREATE TABLE*') "CREATE 文が返る"

$p = Str-Ptr "table"
$rc = [DackAdmin]::DackExportSchema($hPerf, $p, [ref]$r)
Free-Ptr $p
Assert-Equal 0 $rc "DackExportSchema(table) が成功"
$schemaArr = [object[,]]$r
Assert-Equal "キー" $schemaArr[1,8] "PK 列のヘッダ"
Assert-Equal "PK" $schemaArr[2,8] "主キーが PK と印されている"

[void][DackAdmin]::DackClose($hPerf, [ref]$r)

# --- 結果 -------------------------------------------------------------------
Write-Host "`n============================================" -ForegroundColor Cyan
if ($script:failed -eq 0) {
    Write-Host "全 $($script:passed) 項目 成功" -ForegroundColor Green
    exit 0
} else {
    Write-Host "成功 $($script:passed) / 失敗 $($script:failed)" -ForegroundColor Red
    exit 1
}
