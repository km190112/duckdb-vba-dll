# duckdb-vba-dll

**Excel VBA から DuckDB を操作するための、権限階層付き Rust 製 DLL。**

日本語が文字化けせず、10 万行の SELECT が 0.2 秒、2 万行の一括投入が 0.04 秒。
配布する DLL を変えるだけで、利用者ごとに操作範囲を制限できます。

| DLL | 権限 | できること |
|---|---|---|
| `dackdb_r.dll` | ① 読み取り専用 | SELECT のみ |
| `dackdb_rw.dll` | ② 読み書き可 | ① + INSERT / UPDATE / DELETE / トランザクション |
| `dackdb_admin.dll` | ③ 管理者 | ② + DB 作成、CREATE/DROP/ALTER、キー設定、スキーマ出力 |

**64bit Excel 専用**です。DLL 1 ファイルで完結します（DuckDB を静的リンク済み・約 25MB）。

📖 **[利用者向け使い方マニュアル](docs/manual.html)** — VBA を書く人はこちらから

<details>
<summary><b>English summary</b></summary>

A Rust DLL that lets Excel VBA talk to DuckDB, in three permission tiers
(read-only / read-write / admin) — you restrict what a user can do simply by
choosing which DLL you hand them.

The main reason this exists: the official DuckDB ODBC driver returns strings as
`SQL_C_CHAR` (UTF-8 bytes), which the OLE DB provider then reinterprets in the
OS ANSI codepage. On a Japanese Windows install that is CP932, so **all non-ASCII
text comes back mojibake**. This DLL only ever crosses the boundary as UTF-16
`BSTR`, so that class of bug cannot occur — surrogate pairs and emoji round-trip
intact, and there is a test that proves it.

Permissions are enforced in two layers: DuckDB's own `access_mode=READ_ONLY` plus
`lock_configuration` (unbreakable from SQL), and an allowlist over statement types
determined by DuckDB's own parser via `duckdb_extract_statements` +
`duckdb_prepared_statement_type` (so `SELECT 1; DROP TABLE t;` and
`WITH x AS (...) INSERT ...` are both caught).

Windows x64 only. Documentation is in Japanese.
</details>

---

## なぜ公式 ODBC ドライバではなくこれを作ったのか

DuckDB には[公式 ODBC ドライバ](https://duckdb.org/docs/lts/clients/odbc/windows)があり、
`ADODB.Connection` で VBA から使えます。しかし**日本語が構造的に文字化けします**。

ドライバは `SQL_C_CHAR`（UTF-8 バイト列）で文字列を返すのに、OLE DB Provider が
それを OS の ANSI コードページ（日本語環境では CP932）として解釈するためです
（[詳細](https://redraiment.medium.com/solving-the-character-encoding-issue-when-reading-duckdb-via-odbc-in-excel-vba-68fbaffb299d)）。
回避するには毎クエリで `encode()` してバイト列で受け取り、`ADODB.Stream` で
UTF-8 デコードする必要があります。

この DLL は **UTF-16 (BSTR) でしかやり取りしない**ため、原理的に文字化けしません。
サロゲートペア（𠮷）や絵文字も含めて完全に往復することをテストで保証しています。

また、既存の Excel 向け DuckDB 連携（[xlDuckDb](https://github.com/RusselWebber/xlDuckDb)、
[excel-duckdb](https://github.com/bill-ash/excel-duckdb)）には権限階層の仕組みがありません。

---

## クイックスタート

```
配布物の構成:
  dll\dackdb_r.dll        ← 利用者に応じて 1 つだけ配る
  vba\DackDbR.bas         ← 対応するモジュールをインポート
```

1. VBE で「ファイル > ファイルのインポート」から `DackDbR.bas` を読み込む
2. 起動時に一度だけ初期化する

```vba
Private Sub Workbook_Open()
    DackDbInit ThisWorkbook.Path & "\dll"
End Sub
```

3. 使う

```vba
Sub 売上を取得()
    Dim h As LongLong
    h = OpenDb("C:\data\dack.db")

    ' 結果をシートに一括貼り付け（10 万行でも 1 回の代入）
    Dim n As Long
    n = QueryToSheet(h, "SELECT * FROM 売上 WHERE 部署 = '営業部'", Sheet1.Range("A1"))
    MsgBox n & " 件取得しました"

    CloseDb h
End Sub
```

配列として受け取りたい場合：

```vba
Dim arr As Variant
arr = Query(h, "SELECT * FROM 売上")
' arr(1, c) = 列名（ヘッダ行）
' arr(2, c) 以降がデータ
' NULL は IsNull(arr(r, c)) で判定できる
```

### 利用者の入力を条件に使うとき

**文字列連結ではなく `QueryP` / `ExecP` を使ってください。** `?` に値が
バインドされるので、`' OR 1=1; DROP TABLE 売上; --` のような入力でも
ただの文字列として扱われます。

```vba
arr = QueryP(h, "SELECT * FROM 売上 WHERE 部署 = ? AND 金額 > ?", _
             Array(Range("B1").Value, 1000))
```

パラメータは `Array(...)` でも `Range("A1:C1").Value` でも、縦横どちらでも渡せます。

### シートのデータを DB に入れる

```vba
Dim n As LongLong
n = AppendRange(h, "売上", Sheet1.Range("A2:D10000").Value)
```

DuckDB の Appender API を使うため、INSERT 文の反復より 1〜2 桁高速です（実測 2 万行 38ms）。

- **見出し行は含めないこと。** データ行だけの範囲を渡します
- 列はテーブルの列定義の順に対応します（位置指定）
- **途中で失敗した場合は 1 行も入りません**（全体がロールバックされます）
- セルが `#N/A` などのエラー値だと、位置を示してエラーになります
  （黙って NULL にするとデータが静かに壊れるため）

---

## API

すべての関数はエラー時に VBA のエラー（`Err.Raise`）として日本語メッセージを送出します。

| VBA 関数 | ① | ② | ③ | 説明 |
|---|:-:|:-:|:-:|---|
| `DackDbInit(dllFolder)` | ✓ | ✓ | ✓ | DLL を読み込む（起動時に一度） |
| `DackVersion()` | ✓ | ✓ | ✓ | DLL / DuckDB の版数と権限レベル |
| `DackCapabilities()` | ✓ | ✓ | ✓ | この DLL で使える関数の一覧 |
| `OpenDb(path)` | ✓ | ✓ | ✓ | 接続を開きハンドルを返す |
| `CloseDb(h)` | ✓ | ✓ | ✓ | 接続を閉じる |
| `Query(h, sql)` | ✓ | ✓ | ✓ | SELECT → 2 次元 Variant 配列 |
| `QueryP(h, sql, params)` | ✓ | ✓ | ✓ | `?` バインド付き SELECT |
| `QueryToSheet(h, sql, cell)` | ✓ | ✓ | ✓ | SELECT → シートに一括貼り付け |
| `ListTables(h)` | ✓ | ✓ | ✓ | テーブル一覧 |
| `Describe(h, table)` | ✓ | ✓ | ✓ | 列定義 |
| `Exec(h, sql)` | ✗ | ✓ | ✓ | DML → 影響行数 |
| `ExecP(h, sql, params)` | ✗ | ✓ | ✓ | `?` バインド付き DML |
| `AppendRange(h, table, data)` | ✗ | ✓ | ✓ | シート範囲を一括投入（高速） |
| `BeginTx/CommitTx/RollbackTx(h)` | ✗ | ✓ | ✓ | トランザクション |
| `CreateDb(path)` | ✗ | ✗ | ✓ | 新しい `dack.db` を作成 |
| `ExecDDL(h, sql)` | ✗ | ✗ | ✓ | CREATE / DROP / ALTER など |
| `ExportSchema(h, format)` | ✗ | ✗ | ✓ | `"table"` = 配列 / `"ddl"` = CREATE 文 |
| `Checkpoint(h)` | ✗ | ✗ | ✓ | 変更をディスクに確定 |

3 つの DLL は**同じ関数をすべてエクスポート**します。権限外の関数を呼ぶと
`DACK_E_FORBIDDEN` と上位 DLL への案内が返るため、階層② 向けに書いたコードは
階層③ でそのまま動きます。

### 型マッピング（DuckDB → Excel）

| DuckDB | Excel / VBA | 備考 |
|---|---|---|
| NULL | `Null` (VT_NULL) | `IsNull()` で判定、セルは空欄 |
| BOOLEAN | `Boolean` | |
| TINYINT〜INTEGER | `Long` | |
| BIGINT | `LongLong` | |
| UBIGINT (i64 超) / HUGEINT / UUID | `String` | 桁落ちを避けるため厳密な文字列 |
| FLOAT / DOUBLE | `Double` | |
| DECIMAL(≤18, s) | `Double` | `SUM()` が効くよう数値で返す |
| DECIMAL(≥19, s) | `String` | f64 では表せないため丸めずに文字列 |
| VARCHAR | `String` | UTF-16。**日本語が化けない** |
| DATE / TIME / TIMESTAMP | `Date` | Excel が表示できない範囲は ISO 文字列 |
| BLOB | `Byte()` | |
| INTERVAL / ENUM | `String` | |
| LIST / STRUCT / MAP / ARRAY など | — | **クエリごとエラー**。SQL 側で `::VARCHAR` にキャストしてください |

コンテナ型でセルに謎の文字列を入れるのではなく、列名と型名を示してエラーにします。
壊れたデータを黙って貼るより、直し方が分かるほうが良いという判断です。

---

## 権限モデル

防御は 2 層です。

### 層1：DuckDB エンジンレベル（SQL からは突破不可能）

接続時の `duckdb_config` で権限を決めます。

| 設定 | ① | ② | ③ |
|---|---|---|---|
| `access_mode` | `READ_ONLY` | `READ_WRITE` | `READ_WRITE` |
| `enable_external_access` | false | false | 選択可 |
| `allow_unsigned_extensions` | false | false | false |
| `autoload_known_extensions` | false | false | true |
| `lock_configuration` | **true** | **true** | false |

`lock_configuration=true` を**最後に**設定することで、利用者が
`SET access_mode='READ_WRITE'` で昇格することを防いでいます。
階層① はこれだけでほぼ完結し、DuckDB 自身が
`Cannot execute statement of type 'CREATE' on database which is attached in read-only mode!`
を返して全書き込みを拒否します。

### 層2：DuckDB のパーサによる文種別の許可リスト

DuckDB には「DML は可、DDL は不可」というモードが無いため、階層② はここで実装しています。

`duckdb_extract_statements` で文を分割し、`duckdb_prepared_statement_type` で
種別を判定します。**DuckDB 自身のパーサ**を使うので、下記のような入力も確実に検出できます
（正規表現方式では破綻します）。

- `SELECT 1; DROP TABLE t;` — 複数文の後半に紛れた DDL
- `WITH x AS (SELECT 1) INSERT INTO t ...` — 先頭は WITH だが実体は INSERT
- `/* SELECT */ DROP TABLE t` — コメントによる偽装

**1 文でも許可されない文が含まれていたら SQL 全体を拒否**します（部分実行しない）。
許可リスト方式なので、DuckDB に新しい文種別が増えても既定で拒否されます。

### ⚠ セキュリティ境界ではありません

**DLL の分割は運用上の事故防止であって、セキュリティ境界ではありません。**
管理者 DLL のファイルを入手した利用者は何でもできます。

実効的な制限が必要な場合は、`dack.db` 自体に Windows のファイル ACL で
読み取り専用権限を設定してください。層1 と組み合わせると突破できなくなります。

---

## 開発

### ビルド

必要なもの: Rust (x86_64-pc-windows-msvc)、Visual Studio C++ Build Tools、Windows SDK。
clang は不要です（`libduckdb-sys` の生成済みバインディングを使用）。

```bash
cargo build --workspace --release
```

初回は DuckDB のソースをコンパイルするため 3 分ほどかかります。
成果物は `target/release/dackdb_{r,rw,admin}.dll`。

### テスト

```bash
cargo test --workspace
```

```bash
pwsh -File tests/vba_boundary_test.ps1
```

`tests/vba_boundary_test.ps1` が本プロジェクトの要です。
.NET の `[MarshalAs(UnmanagedType.Struct)] ref object` は VARIANT* としてマーシャリング
されるため、**VBA の `ByRef result As Variant` とまったく同じ ABI** を通ります。
Excel を起動せずに VBA 境界そのものを検証できます。検証内容:

- 日本語・サロゲートペア・絵文字の完全往復
- 下限 1 の 2 次元配列（`Range.Value` と同じ形）
- 型マッピングと NULL の扱い
- 権限階層と SQL 経由のバイパス試験
- パラメータバインドと SQL インジェクション耐性
- 一括投入（2 万行）と列数不一致の検出
- 不正ハンドル・null ポインタでクラッシュしないこと
- 10 万行 × 6 列の性能
- 1000 回クエリでのメモリリーク

### VBA モジュールの再生成

`vba/*.bas` は自動生成物です。直接編集せず、テンプレートを直してから生成してください。

```bash
pwsh -File vba/generate_modules.ps1
```

**生成物は Shift-JIS (CP932) で出力されます。** VBE の「ファイルのインポート」は
`.bas` を CP932 として読むため、UTF-8 のままでは日本語コメントが全て文字化けします。

### 構成

```
crates/dackdb-core/src/
  oleaut.rs    VARIANT / SAFEARRAY の手書き Win32 バインディング
  variant.rs   出力: 2 次元 SAFEARRAY の組み立て（列優先の並べ替え）
  value.rs     出力: DuckDB の値 → VARIANT
  inbound.rs   入力: VBA の配列読み取り、VARIANT → duckdb_value
  append.rs    Appender API による一括投入
  raw.rs       libduckdb-sys の RAII ラッパ
  conn.rs      接続生成（層1）とハンドルレジストリ
  classify.rs  文種別判定と権限ゲート（層2）
  query.rs     クエリ実行
  schema.rs    スキーマ出力
  api.rs       公開 API の実装本体
  ffi.rs       export_dackdb_ffi! マクロ（catch_unwind を含む）
crates/dackdb-{r,rw,admin}/src/lib.rs   マクロを 1 行呼ぶだけ
```

#### 設計上の注意点

- **`panic = "abort"` を設定しないこと。** FFI 境界は `catch_unwind` で panic を
  捕まえて VBA にエラーを返す設計です。abort にすると Excel がプロセスごと落ちます。
- **権限を Cargo の feature で切り替えないこと。** ワークスペース一括ビルド時に
  feature 統合が起きて 3 つの DLL すべてが管理者権限になります。
  レベルは `export_dackdb_ffi!` のマクロ引数として渡しています。
- **`windows` クレートは使っていません。** `VARIANT` に `Drop` があると、VBA 所有の
  メモリに書き込む本用途では二重解放の危険があるためです。詳細は `oleaut.rs` 冒頭。
- **`vba/*.bas` は Shift-JIS (CP932) です。** VBE のインポートがそう読むためで、
  GitHub の Web 画面では日本語コメントが文字化けして見えます。編集するのは
  UTF-8 の `vba/DackDb.template.bas` のほうです。

---

## ライセンス

MIT License. 詳細は [LICENSE](LICENSE) を参照してください。

DuckDB 本体も MIT License です。

---

## 制約事項

### 同じファイルを同時に開けない

DuckDB は**同一プロセス内で 1 つのファイルを複数回開けません**。
Excel の 1 セッション内で

```vba
hAdmin = CreateDb("C:\data\dack.db")   ' 書き込みで開いている
hRead  = OpenDb("C:\data\dack.db")     ' ← IO Error になる
```

のような使い方はできません。先に `CloseDb hAdmin` してください。
別プロセス（別の Excel）から読み取り専用で開くことは可能です。

### `?` バインドは 1 文のみ

`QueryP` / `ExecP` は SQL を 1 文だけ受け付けます。複数文を許すと
どのパラメータがどの文に対応するかが曖昧になり、検査した内容と実行内容が
ずれる余地が生まれるためです。

### コンテナ型はキャストが必要

LIST / STRUCT / MAP / ARRAY などを含む列は、SQL 側で `::VARCHAR` に
キャストしてください。詳細は上の型マッピング表を参照。
