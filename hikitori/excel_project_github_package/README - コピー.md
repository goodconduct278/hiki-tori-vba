# Excel VBA / Power Query / Python in Excel 管理パッケージ

このフォルダは、アップロードされた3つのExcelブックをGitHub/Codex/Claude Codeで読みやすくするために分解したものです。

## 重要
- VBA本体は、GitHub上では `.bas` / `.cls` / `.frm` として `/vba/` に置いてください。
- このZIP内の `/vba/_raw_vbaProject_bin/` はExcel内のVBAプロジェクトのバイナリ退避です。Codexが読むための主ファイルではありません。
- Power Queryは `/powerquery/`、ExcelのPython関連は `/python/`、説明資料は `/docs/` に置いています。
-エクセル本体を入れていますがファイルの破損や問題が起こった際の参照用のデータなので必要でなければ読み込む必要はありません

## 推奨配置
```text
/vba/
  一括入力.bas
  フォルダ作成VBA.bas
  個別引き取りVBA.bas
/powerquery/
/python/
/docs/
/workbooks_summary/
```

## 対象ブック
- ファイル作成マクロ.xlsm：シート5、テーブル0、数式3、接続0、VBA=あり、Python in Excel=なし
- 班員個人引き取り予定(1).xlsm：シート5、テーブル3、数式1、接続0、VBA=あり、Python in Excel=なし
- 班員合計引き取り予定(3).xlsm：シート14、テーブル8、数式592、接続6、VBA=あり、Python in Excel=あり
