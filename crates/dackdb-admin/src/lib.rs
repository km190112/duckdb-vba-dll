//! ③ 管理者 DLL（`dackdb_admin.dll`）。DB 作成・DDL・スキーマ出力まで全機能。
//!
//! 実体はすべて `dackdb-core` にある。このクレートは権限レベルを 1 つ指定するだけ。

dackdb_core::export_dackdb_ffi!(dackdb_core::Level::Admin);
