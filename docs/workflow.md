# 処理フロー整理メモ

## 全体方針
このプロジェクトは、VBA・Power Query・Python in Excelが連動するExcel業務ツールとして管理する。

## 読ませる順番
1. `/docs/project_overview.md`
2. `/workbooks_summary/*_sheets.md`
3. `/workbooks_summary/*_tables.md`
4. `/powerquery/*_connections.md`
5. `/python/*.py` と `*_py_formulas.md`
6. `/vba/*.bas`（GitHub上に別途追加済みのVBAコード）

## 変更時の注意
- VBAだけを修正すると、Power QueryやPY関数との列名・シート名の不整合が起きる可能性がある。
- Power Query側のクエリ名、テーブル名、出力シート名はVBAから参照される可能性があるため、名称変更は慎重に行う。
- Python in Excelは実行環境や認証状態に依存するため、Codex/Claude上では「コードレビュー」は可能でもExcel上の再計算結果までは保証できない。
