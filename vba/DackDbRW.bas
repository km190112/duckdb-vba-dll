Attribute VB_Name = "DackDbRW"
'==============================================================================
' DackDbRW - DuckDB アクセスモジュール【② 読み書き可】
'
' 対象 DLL   : dackdb_rw.dll
' できること : SELECT / INSERT / UPDATE / DELETE / トランザクション
' できないこと: CREATE/DROP/ALTER などのスキーマ変更、DB 作成
'
' 【重要】このモジュールは 64bit Excel 専用です。
'
' 【このファイルは自動生成されています】
'   直接編集せず vba\DackDb.template.bas を直してから
'   pwsh -File vba\generate_modules.ps1 を実行してください。
'
' 【使い方】
'   1. このファイルを VBE で「ファイル > ファイルのインポート」する
'   2. dackdb_rw.dll を任意のフォルダに置く
'   3. 起動時に一度 DackDbInit "C:\path\to\dll_folder" を呼ぶ
'   4. h = OpenDb("C:\data\dack.db")
'      QueryToSheet h, "SELECT * FROM 売上", Sheet1.Range("A1")
'      CloseDb h
'
' 【設計上の約束】
'   ・入力文字列は StrPtr() で生の UTF-16 ポインタとして渡す。
'     VBA の Declare は String を必ず ANSI（日本語環境では CP932）に変換して
'     しまうため、String のまま渡すと日本語が文字化けする。
'   ・出力は必ず ByRef ... As Variant で受け取る。成功時は値、失敗時は
'     エラーメッセージ文字列が同じ引数に入る。
'   ・戻り値 0 が成功、負値がエラー。
'
' 【3 つのモジュールの関係】
'   DackDbR / DackDbRW / DackDbAdmin は Lib の DLL 名だけが異なります。
'   下位向けに書いたコードは上位モジュールでもそのまま動きます。
'==============================================================================
Option Explicit

Private Const DLL_NAME As String = "dackdb_rw.dll"
Private Const ERR_BASE As Long = vbObjectError + 1000

' 戻り値コード（Rust 側 api.rs と対応）
Private Const DACK_OK As Long = 0
Private Const DACK_E_FORBIDDEN As Long = -403
Private Const DACK_E_PANIC As Long = -999

'------------------------------------------------------------------ 宣言部 ----
' 文字列引数は StrPtr() で渡す UTF-16 文字列ポインタ。
' 最後の result は必ず ByRef Variant。

Private Declare PtrSafe Function LoadLibraryW Lib "kernel32" ( _
    ByVal lpFileName As LongPtr) As LongPtr

Private Declare PtrSafe Function DackVersionRaw Lib "dackdb_rw.dll" Alias "DackVersion" ( _
    ByRef result As Variant) As Long

Private Declare PtrSafe Function DackCapabilitiesRaw Lib "dackdb_rw.dll" Alias "DackCapabilities" ( _
    ByRef result As Variant) As Long

Private Declare PtrSafe Function DackOpen Lib "dackdb_rw.dll" ( _
    ByVal pathPtr As LongPtr, ByRef result As Variant) As Long

Private Declare PtrSafe Function DackClose Lib "dackdb_rw.dll" ( _
    ByVal handle As LongLong, ByRef result As Variant) As Long

Private Declare PtrSafe Function DackQuery Lib "dackdb_rw.dll" ( _
    ByVal handle As LongLong, ByVal sqlPtr As LongPtr, ByRef result As Variant) As Long

Private Declare PtrSafe Function DackQueryParams Lib "dackdb_rw.dll" ( _
    ByVal handle As LongLong, ByVal sqlPtr As LongPtr, ByRef params As Variant, _
    ByRef result As Variant) As Long

Private Declare PtrSafe Function DackListTables Lib "dackdb_rw.dll" ( _
    ByVal handle As LongLong, ByRef result As Variant) As Long

Private Declare PtrSafe Function DackDescribe Lib "dackdb_rw.dll" ( _
    ByVal handle As LongLong, ByVal tablePtr As LongPtr, ByRef result As Variant) As Long

Private Declare PtrSafe Function DackExecute Lib "dackdb_rw.dll" ( _
    ByVal handle As LongLong, ByVal sqlPtr As LongPtr, ByRef result As Variant) As Long

Private Declare PtrSafe Function DackExecuteParams Lib "dackdb_rw.dll" ( _
    ByVal handle As LongLong, ByVal sqlPtr As LongPtr, ByRef params As Variant, _
    ByRef result As Variant) As Long

Private Declare PtrSafe Function DackAppendArray Lib "dackdb_rw.dll" ( _
    ByVal handle As LongLong, ByVal tablePtr As LongPtr, ByRef data As Variant, _
    ByRef result As Variant) As Long

Private Declare PtrSafe Function DackBegin Lib "dackdb_rw.dll" ( _
    ByVal handle As LongLong, ByRef result As Variant) As Long

Private Declare PtrSafe Function DackCommit Lib "dackdb_rw.dll" ( _
    ByVal handle As LongLong, ByRef result As Variant) As Long

Private Declare PtrSafe Function DackRollback Lib "dackdb_rw.dll" ( _
    ByVal handle As LongLong, ByRef result As Variant) As Long

Private Declare PtrSafe Function DackCreateDatabase Lib "dackdb_rw.dll" ( _
    ByVal pathPtr As LongPtr, ByRef result As Variant) As Long

Private Declare PtrSafe Function DackExecuteDDL Lib "dackdb_rw.dll" ( _
    ByVal handle As LongLong, ByVal sqlPtr As LongPtr, ByRef result As Variant) As Long

Private Declare PtrSafe Function DackExportSchema Lib "dackdb_rw.dll" ( _
    ByVal handle As LongLong, ByVal formatPtr As LongPtr, ByRef result As Variant) As Long

Private Declare PtrSafe Function DackCheckpoint Lib "dackdb_rw.dll" ( _
    ByVal handle As LongLong, ByRef result As Variant) As Long

'------------------------------------------------------------------ 初期化 ----

' DLL を明示的に読み込む。Declare の Lib にフルパスを直書きしなくて済むよう、
' 起動時に一度だけ呼ぶ。以降の Declare は読み込み済みモジュールに解決される。
'
'   例: DackDbInit ThisWorkbook.Path & "\dll"
Public Sub DackDbInit(ByVal dllFolder As String)
    If Not IsWin64() Then
        Err.Raise ERR_BASE, "DackDbRW", _
            "このモジュールは 64bit 版 Excel 専用です。" & vbCrLf & _
            "「ファイル > アカウント > Excel のバージョン情報」で確認してください。"
    End If

    Dim fullPath As String
    fullPath = RTrim$(dllFolder)
    If Right$(fullPath, 1) = "\" Then fullPath = Left$(fullPath, Len(fullPath) - 1)
    fullPath = fullPath & "\" & DLL_NAME

    If Dir$(fullPath) = "" Then
        Err.Raise ERR_BASE, "DackDbRW", "DLL が見つかりません: " & fullPath
    End If

    If LoadLibraryW(StrPtr(fullPath)) = 0 Then
        Err.Raise ERR_BASE, "DackDbRW", _
            "DLL の読み込みに失敗しました: " & fullPath & vbCrLf & _
            "32bit 版 Excel を使っていないか確認してください。"
    End If
End Sub

Private Function IsWin64() As Boolean
    #If Win64 Then
        IsWin64 = True
    #Else
        IsWin64 = False
    #End If
End Function

'------------------------------------------------------- 情報取得（全階層）----

' DLL の版数と権限レベルを返す。
Public Function DackVersion() As String
    Dim v As Variant
    Check DackVersionRaw(v), v, "DackVersion"
    DackVersion = CStr(v)
End Function

' この DLL で使える関数の一覧をカンマ区切りで返す。
Public Function DackCapabilities() As String
    Dim v As Variant
    Check DackCapabilitiesRaw(v), v, "DackCapabilities"
    DackCapabilities = CStr(v)
End Function

'--------------------------------------------------------------- 接続管理 ----

' データベースを開いて接続ハンドルを返す。
Public Function OpenDb(ByVal dbPath As String) As LongLong
    Dim v As Variant
    Check DackOpen(StrPtr(dbPath), v), v, "OpenDb"
    OpenDb = CLngLng(v)
End Function

' 接続を閉じる。閉じ忘れるとファイルがロックされたままになる。
Public Sub CloseDb(ByVal handle As LongLong)
    Dim v As Variant
    Check DackClose(handle, v), v, "CloseDb"
End Sub

'----------------------------------------------------------------- 読み取り ----

' SELECT を実行して 2 次元 Variant 配列を返す。
'   arr(1, c) = 列名（ヘッダ行）
'   arr(2, c) 以降がデータ
Public Function Query(ByVal handle As LongLong, ByVal sql As String) As Variant
    Dim v As Variant
    Check DackQuery(handle, StrPtr(sql), v), v, "Query"
    Query = v
End Function

' SELECT の結果をシートに貼り付け、データ行数を返す（ヘッダ行は含まない）。
' 10 万行でも 1 回の代入で済むので高速。
'
'   例: n = QueryToSheet(h, "SELECT * FROM 売上", Sheet1.Range("A1"))
Public Function QueryToSheet(ByVal handle As LongLong, ByVal sql As String, _
                             ByVal targetCell As Range) As Long
    Dim arr As Variant
    arr = Query(handle, sql)

    Dim nRows As Long, nCols As Long
    nRows = UBound(arr, 1)
    nCols = UBound(arr, 2)

    targetCell.Resize(nRows, nCols).Value = arr
    QueryToSheet = nRows - 1   ' ヘッダ行を除いたデータ行数
End Function

' パラメータ付きで SELECT を実行する。SQL 中の ? に値が順番に入る。
'
' 利用者の入力を SQL に埋め込むときは必ずこちらを使うこと。文字列連結だと
' 「' OR 1=1 --」のような入力で意図しない SQL になる（SQL インジェクション）。
'
'   例: arr = QueryP(h, "SELECT * FROM 売上 WHERE 部署 = ? AND 金額 > ?", _
'                    Array(Range("B1").Value, 1000))
'
' params は Array(...) でも Range("A1:C1").Value でも、縦横どちらでも受け付ける。
Public Function QueryP(ByVal handle As LongLong, ByVal sql As String, _
                       ByRef params As Variant) As Variant
    Dim v As Variant
    Check DackQueryParams(handle, StrPtr(sql), params, v), v, "QueryP"
    QueryP = v
End Function

' テーブル一覧を 2 次元配列で返す。
Public Function ListTables(ByVal handle As LongLong) As Variant
    Dim v As Variant
    Check DackListTables(handle, v), v, "ListTables"
    ListTables = v
End Function

' 指定テーブルの列定義（列名・型・NULL 可否・既定値）を 2 次元配列で返す。
Public Function Describe(ByVal handle As LongLong, ByVal tableName As String) As Variant
    Dim v As Variant
    Check DackDescribe(handle, StrPtr(tableName), v), v, "Describe"
    Describe = v
End Function

'--------------------------------------------------- 書き込み（階層② 以上）----
' この DLL で使えます。

' INSERT / UPDATE / DELETE を実行し、影響行数を返す。
Public Function Exec(ByVal handle As LongLong, ByVal sql As String) As LongLong
    Dim v As Variant
    Check DackExecute(handle, StrPtr(sql), v), v, "Exec"
    Exec = CLngLng(v)
End Function

' パラメータ付きで DML を実行し、影響行数を返す。
'
'   例: n = ExecP(h, "UPDATE 売上 SET 金額 = ? WHERE id = ?", Array(5000, 3))
Public Function ExecP(ByVal handle As LongLong, ByVal sql As String, _
                      ByRef params As Variant) As LongLong
    Dim v As Variant
    Check DackExecuteParams(handle, StrPtr(sql), params, v), v, "ExecP"
    ExecP = CLngLng(v)
End Function

' シートの範囲をテーブルへ一括投入し、投入行数を返す。
' INSERT 文を 1 行ずつ実行するより 1?2 桁速い。Excel → DB の主力経路。
'
' 【重要】
'   ・見出し行は含めないこと（データ行だけの範囲を渡す）
'   ・列はテーブルの列定義の順に対応する（位置指定）
'   ・途中で失敗した場合は 1 行も入らない（全体がロールバックされる）
'
'   例: n = AppendRange(h, "売上", Sheet1.Range("A2:D1000").Value)
Public Function AppendRange(ByVal handle As LongLong, ByVal tableName As String, _
                            ByRef data As Variant) As LongLong
    Dim v As Variant
    Check DackAppendArray(handle, StrPtr(tableName), data, v), v, "AppendRange"
    AppendRange = CLngLng(v)
End Function

Public Sub BeginTx(ByVal handle As LongLong)
    Dim v As Variant
    Check DackBegin(handle, v), v, "BeginTx"
End Sub

Public Sub CommitTx(ByVal handle As LongLong)
    Dim v As Variant
    Check DackCommit(handle, v), v, "CommitTx"
End Sub

Public Sub RollbackTx(ByVal handle As LongLong)
    Dim v As Variant
    Check DackRollback(handle, v), v, "RollbackTx"
End Sub

'------------------------------------------------------- 管理（階層③ のみ）----
' この DLL（読み書き可）で呼ぶとエラーになります。
' 使うには DackDbAdmin モジュール + dackdb_admin.dll に切り替えてください。

' 新しい dack.db を作成して開く。既存ファイルがある場合はエラーになる。
Public Function CreateDb(ByVal dbPath As String) As LongLong
    Dim v As Variant
    Check DackCreateDatabase(StrPtr(dbPath), v), v, "CreateDb"
    CreateDb = CLngLng(v)
End Function

' CREATE / DROP / ALTER / ATTACH などのスキーマ変更を実行する。
Public Function ExecDDL(ByVal handle As LongLong, ByVal sql As String) As LongLong
    Dim v As Variant
    Check DackExecuteDDL(handle, StrPtr(sql), v), v, "ExecDDL"
    ExecDDL = CLngLng(v)
End Function

' スキーマ情報を出力する。
'   format = "table" … 2 次元配列（そのままシートに貼れる）
'   format = "ddl"   … CREATE 文を連結した文字列
Public Function ExportSchema(ByVal handle As LongLong, _
                             Optional ByVal format As String = "table") As Variant
    Dim v As Variant
    Check DackExportSchema(handle, StrPtr(format), v), v, "ExportSchema"
    ExportSchema = v
End Function

' 変更をディスクに確定させる。
Public Sub Checkpoint(ByVal handle As LongLong)
    Dim v As Variant
    Check DackCheckpoint(handle, v), v, "Checkpoint"
End Sub

'--------------------------------------------------------------- 内部処理 ----

' DLL の戻り値を検査し、エラーなら VBA のエラーとして送出する。
' 失敗時は result に日本語のエラーメッセージが入っている。
Private Sub Check(ByVal rc As Long, ByRef result As Variant, ByVal funcName As String)
    If rc = DACK_OK Then Exit Sub

    Dim msg As String
    If VarType(result) = vbString Then
        msg = CStr(result)
    Else
        msg = "不明なエラー (コード " & rc & ")"
    End If

    Select Case rc
        Case DACK_E_FORBIDDEN
            msg = "【権限エラー】" & vbCrLf & msg
        Case DACK_E_PANIC
            msg = "【内部エラー】" & vbCrLf & msg
    End Select

    Err.Raise ERR_BASE + Abs(rc), "DackDbRW." & funcName, msg
End Sub
