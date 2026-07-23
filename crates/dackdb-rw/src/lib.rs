//! ② 読み書き可 DLL（`dackdb_rw.dll`）。SELECT と DML は可、DDL は不可。
//!
//! 実体はすべて `dackdb-core` にある。このクレートは権限レベルを 1 つ指定するだけ。

dackdb_core::export_dackdb_ffi!(dackdb_core::Level::ReadWrite);
