# 設計書妥当性レビュー

## 重大な見落とし（要修正）

### fuzzy matching は Power Query で再現不可能

設計書タスクBは「Python処理をPower Queryへ移植」と記載しているが、**SequenceMatcher.ratio() による類似度計算はM言語に存在しない**。

Power Queryは行列変換・結合・フィルタが主機能であり、文字列アルゴリズム的処理は対象外。

**正しいアーキテクチャ**:

```
Python script_1.py の処理内訳
├── データ収集・結合・未変換抽出 → 既にPQで完結（変更不要）
└── fuzzy matching（変換候補一覧の生成）→ VBAで代替（新規実装）
```

### 実態：PQはすでに95%完成している

現行の `_Formulas_Section1.m` には以下が実装済み:
- `★班員データ` - 班員ファイルフォルダ収集
- `★変換済み結合` - 変換リストとの正規化マッチング結合
- `★変換済みデータ` - 入力フォーム表示名・変換状態の付与
- `★未変換一覧` - 未変換行の抽出

つまりタスクBの成果物「powerquery/*_Formulas_Section1.m への変更案」は**変更なし**が正解。

---

## その他のギャップ

| 項目 | 問題 | 対処 |
|---|---|---|
| `USE_PYTHON_PATH` 格納先 | 設計書が「設定シートまたは名前定義」と曖昧 | `プログラム設定` シート A3 に固定（A1:A5のうちA3が空き） |
| `変換候補一覧` 列構造未定義 | VBAが読む列番号が文書化されていない | A:元材料名, B:区分, C:候補1, D:スコア1, E:候補2, F:スコア2, G:候補3, H:スコア3 |
| `initialization.py` の扱い | Excelのboilerplate。移行対象ではない | Python退役後はリポジトリに残置（実行されない） |
| 並走検証の前提 | PY()セルはPython環境が必要。環境がなければ旧経路を実行できない | 開発者PCで旧経路を事前実行し結果をスナップショット保存する方式に変更 |
| `★変換済みデータ` テーブル列数 | tables.md抽出時は16列だが、PQ更新後は19列（区分・取込UR・変換状態追加） | PQ更新後の列定義をModule2の定数（CONV_COL_*）で信頼 |

---

## 変更対象ファイル一覧

### 新規作成
| ファイル | 内容 |
|---|---|
| `docs/migration/design_review.md` | 本ファイル |
| `docs/migration/python_logic_inventory.md` | Python処理の機能分解 |
| `docs/migration/input_output_contract.md` | 入出力仕様 |
| `docs/migration/python_to_powerquery_mapping.md` | 処理対応表（PQ変更なしの根拠含む） |
| `docs/migration/feature_flag_design.md` | USE_PYTHON_PATH フラグ設計 |
| `docs/migration/parallel_run_test_plan.md` | 並走検証手順 |
| `docs/migration/diff_check_template.md` | 差分チェックテンプレート |
| `vba/_raw_vbaProject_bis/標準モジュール/班員合計引き取り予定(3)Module4_候補生成.bas` | fuzzy matching VBA（PY()代替） |

### 変更なし
| ファイル | 理由 |
|---|---|
| `powerquery/班員合計引き取り予定(3)_Formulas_Section1.m` | 既に完成済み、変更不要 |
| `vba/_raw_vbaProject_bis/標準モジュール/班員合計引き取り予定(3)Module2.bas` | `担当者データ抽出`・`LoadCandidateMap` は変更不要 |
| `vba/_raw_vbaProject_bis/標準モジュール/班員合計引き取り予定(3)Module1.bas` | 変更不要 |

### Excel本体での手動作業（リポジトリ外）
1. `変換候補一覧` シートのA5セルからPY()数式を削除
2. `プログラム設定` シートのA3セルに `USE_PYTHON_PATH` ラベル（A列）と値 `FALSE`（B列等）を追加
3. 「変換候補を生成」ボタンを`変換候補一覧を生成する`マクロに割り当て変更
