# Python → Power Query / VBA 処理対応表

## 結論

**Power Queryの変更は不要。** 既存M式がすでに必要な変換を実装済み。
Python処理のうち移行が必要なのは「変換候補一覧の生成（fuzzy matching）」のみであり、これはVBAで代替する。

---

## 処理対応マトリクス

| Python処理 | 現行実装 | 移行先 | 変更の有無 |
|---|---|---|---|
| 班員データフォルダ収集 | `★班員データ`（PQ） | 変更なし | なし |
| 変換リスト読み込み | `変換リスト`（PQ） | 変更なし | なし |
| 正規化キーでの突合 | `★変換済み結合`（PQ） | 変更なし | なし |
| 変換状態・表示名の付与 | `★変換済みデータ`（PQ） | 変更なし | なし |
| 未変換行の抽出 | `★未変換一覧`（PQ） | 変更なし | なし |
| **fuzzy matching（候補生成）** | `変換候補一覧!A5` の PY()セル | **VBA Module4** | **新規実装** |

---

## Power Queryの既存実装との対応詳細

### Python `正規化()` ↔ PQ `fnNormalize()`

| 処理 | Python | PQ M言語 | 差異 |
|---|---|---|---|
| 全角→半角 | `unicodedata.normalize("NFKC")` | `StrConv`相当なし、個別Replace | 実用上ほぼ同等 |
| 大文字/小文字 | `upper()` | `Text.Lower()` | Python=大文字、PQ=小文字（内部比較のみなので問題なし） |
| スペース除去 | `re.sub(r"\s+")` | `Text.Remove(narrow, {" ", "　", ...})` | 同等 |
| ハイフン統一 | 複数パターン→`-` | 複数Replace | 同等 |
| スラッシュ統一 | `／`→`/` | `Text.Replace(...)` | 同等 |
| 丸括弧統一 | `（）`→`()` | `Text.Replace(...)` | 同等 |
| ml正規化 | `ｍｌ`→`ml`, `?`→`ml` | `ml1`, `ml2` ステップ | 同等 |
| 単位変換 | `㎡→M2`, `㎜→MM`, `メートル→M` | **未実装** | PQには単位変換なし（材料名突合への影響は限定的） |

**単位変換の差異について**: PQは材料名を変換リストに突合するのみ（完全一致・正規化後一致）。
fuzzy matchingはVBAが担当するため、単位変換はVBA `NormalizeForMatch()` で実装済み。PQへの追加は不要。

### Python の UR判定 ↔ PQ の UR整備ステップ

```python
# Python
def UR案件か(value):
    return "UR" in text.upper()

def UR品か(value):
    return text in ["UR", "○", "〇", "1", "TRUE", "YES", "対象"]
```

```m
// PQ (★班員データ)
UR整備 = Table.TransformColumns(...,
    each if Text.Contains(Text.Upper(Text.From(_)), "UR") then "UR" else null
),
区分追加 = Table.AddColumn(...,
    each if [取込UR] = "UR" then "UR" else "通常"
)
```

PQの`区分`列が`★変換済みデータ`に引き継がれるため、VBAはこの列を読むだけでよい。

---

## fuzzy matching がPQで実現できない理由

Power Query M言語には以下が存在しない:
- `Text.Similarity()` 関数（存在しない）
- `SequenceMatcher` 相当のアルゴリズム
- 動的スコアリングと上位N件ソート

M言語は宣言的変換言語であり、反復アルゴリズム（ループ内ループでのスコア集積）には不適。
`List.Generate` を使えば原理的には可能だが、数百行×数百行の総当たりはExcel上で非実用的。

→ **VBAが唯一の現実的な代替手段**。

---

## Power Query ファイルへの変更

`powerquery/班員合計引き取り予定(3)_Formulas_Section1.m` : **変更なし**

現状のM式で要件を満たしている。
