# Codex / Claude Code 用 指示文

このリポジトリはExcel業務ツールです。VBA、Power Query、Python in Excelが連動しています。

まず以下を全部読んでください。
- `README - 仕様目的.md`
- `/docs/workflow.md`
- `/workbooks_summary/`
- `/powerquery/`
- `/python/`
- `/vba/` の `.bas` ファイル

最初にやること：
1. 処理全体の流れを説明してください。
2. VBA、Power Query、PY関数の依存関係を整理してください。
3. 壊れやすい箇所、列名・シート名変更で影響が出そうな箇所を指摘してください。
4. 修正はまだしないでください。

修正を依頼された場合：
- 変更範囲を最小限にしてください。
- 既存のファイル作成処理、Power Queryの出力、PY関数の出力列を壊さないでください。
- 変更前に影響範囲を説明してください。
- 変更後は、どのファイルのどこを変えたかを要約してください。
