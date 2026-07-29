//! ② 読み書き可 DLL（`duckdb_rw.dll`）。SELECT と DML は可、DDL は不可。
//!
//! 実体はすべて `duckdb-core` にある。このクレートは権限レベルを 1 つ指定するだけ。

duckdb_core::export_duckdb_ffi!(duckdb_core::Level::ReadWrite);
