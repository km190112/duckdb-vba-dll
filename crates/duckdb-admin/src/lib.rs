//! ③ 管理者 DLL（`duckdb_admin.dll`）。DB 作成・DDL・スキーマ出力まで全機能。
//!
//! 実体はすべて `duckdb-core` にある。このクレートは権限レベルを 1 つ指定するだけ。

duckdb_core::export_duckdb_ffi!(duckdb_core::Level::Admin);
