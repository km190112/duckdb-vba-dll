//! ① 読み取り専用 DLL（`dackdb_r.dll`）。SELECT のみ。
//!
//! 実体はすべて `dackdb-core` にある。このクレートは権限レベルを 1 つ指定するだけ。
//! 書き込み系・DDL 系の関数もエクスポートされるが、呼ぶと `DACK_E_FORBIDDEN` と
//! 上位 DLL への案内を返す。

dackdb_core::export_dackdb_ffi!(dackdb_core::Level::Read);
