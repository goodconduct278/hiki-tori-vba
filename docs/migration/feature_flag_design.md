# 段階移行フラグ設計: USE_PYTHON_PATH

## 目的

移行期間中に「旧Python経路」と「新VBA経路」を切り替え可能にする。
トラブル時に即座にロールバックできる安全装置。

---

## フラグの格納場所

**ブック**: `班員合計引き取り予定(3).xlsm`

**シート**: `プログラム設定`（既存の非表示シート）

| セル | 内容 | 既存利用 |
|---|---|---|
| A1 | 自身のフォルダパス（マクロが自動書込） | 使用中 |
| A2 | 保存先フォルダパス | 使用中 |
| **A3** | **`USE_PYTHON_PATH` の値（TRUE/FALSE）** | **新規** |
| A4 | 保存ファイル名（四工品） | 使用中 |
| A5 | 保存ファイル名（仕入品） | 使用中 |

A3セルに `FALSE`（または空白）を設定することで、VBA経路が有効になる。

---

## フラグの意味

| 値 | 動作 |
|---|---|
| `FALSE`（または空白） | VBA経路を使用（`変換候補一覧を生成する()` が実行される）**← デフォルト** |
| `TRUE` | VBA経路をスキップ。「Python経路を使用中」メッセージを表示し終了 |

---

## VBA側の実装（Module4内）

```vba
Private Const SETTING_SHEET   As String = "プログラム設定"
Private Const USE_PY_FLAG_ROW As Long = 3
Private Const USE_PY_FLAG_COL As Long = 1

' フラグチェック
Dim usePyPath As Boolean
Dim wsSet As Worksheet
On Error Resume Next
Set wsSet = ThisWorkbook.Sheets(SETTING_SHEET)
On Error GoTo 0

If Not wsSet Is Nothing Then
    Dim flagVal As Variant
    flagVal = wsSet.Cells(USE_PY_FLAG_ROW, USE_PY_FLAG_COL).Value
    usePyPath = (UCase(Trim(CStr(flagVal))) = "TRUE")
End If

If usePyPath Then
    MsgBox "プログラム設定の USE_PYTHON_PATH が TRUE のため" & vbCrLf & _
           "このマクロはスキップします。" & vbCrLf & vbCrLf & _
           "VBA経路に切り替えるには A3 セルを FALSE または空白にしてください。", _
           vbInformation, "Python経路が有効"
    Exit Sub
End If
```

---

## Excel本体での設定手順

### VBA経路へ切り替え（移行）

1. `プログラム設定` シートを右クリック → 再表示
2. A3セルに `FALSE` を入力（または空白にする）
3. シートを再非表示
4. `変換候補一覧` シートのA5セルからPY()数式を削除
5. `変換候補を生成` ボタンのマクロ割り当てを `変換候補一覧を生成する` に変更

### Python経路へロールバック

1. `プログラム設定` シートを再表示
2. A3セルに `TRUE` を入力
3. `変換候補一覧` シートのA5セルに以下の数式を再入力:
   ```
   =_xlfn._xlws.PY(0,1,未変換一覧[],未一致一覧,テーブル1[])
   ```
4. ボタンのマクロ割り当てを旧マクロに戻す

---

## 移行スケジュール例

```
フェーズ1（並走期間）:
  USE_PYTHON_PATH = TRUE  → Python経路が主
  VBAマクロを手動実行して出力を比較（テスト目的のみ）

フェーズ2（VBA経路切替）:
  USE_PYTHON_PATH = FALSE → VBA経路が主
  PY()セル削除
  1〜2週間の本番運用で問題がないか確認

フェーズ3（Python撤去）:
  A3セルを削除
  python/ ディレクトリを dev-only として .gitignore または archive/移動
  initialization.py の Excel依存ライブラリをコメントアウト
```

---

## リスク管理

| リスク | 対処 |
|---|---|
| ロールバック時にPY()数式を忘れた場合 | PY()数式を `docs/migration/` に記録済み（本ファイルに掲載）|
| A3セルが誤って削除された場合 | 空白 = FALSE として扱うため、VBA経路が自動選択される（安全側） |
| 複数人が同時に設定を変更する場合 | ファイルサーバ上での排他制御に依存（Excelの通常運用と同様） |
