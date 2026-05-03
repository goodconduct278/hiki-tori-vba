# Python処理 機能分解インベントリ

対象ファイル: `python/班員合計引き取り予定(3)_python_script_1.py`

---

## 機能分解

### 関数一覧

| 関数名 | 役割 | 移行先 |
|---|---|---|
| `正規化(text)` | NFKC正規化・大文字化・記号除去・単位変換 | VBA: `NormalizeForMatch()` |
| `類似度(a, b)` | 文字列類似度スコア（0〜100）を返す | VBA: `SimilarityScore()` |
| `UR案件か(value)` | 区分文字列に"UR"が含まれるか | VBA: `IsUrCase()` |
| `UR品か(value)` | 製品がUR対象品かどうか | VBA: `IsUrProduct()` |

### メイン処理フロー

```
1. 入力読み込み
   ├── 未変換一覧（P2 or P3 fallback）→ 7列: Name, 担当者, 区分, 材料, 数量, 単位, 納品日
   └── 変換リスト（P4）→ 4列: 変換前（材料名）, 変換後（製品名）, UR, メーカー

2. 前処理
   ├── 未変換: iloc[:, :7] で7列に限定、列名付与
   ├── 変換リスト: iloc[:, :4] で4列に限定
   ├── 未変換: 材料列がnullの行を除去
   └── 変換リスト: 変換前・変換後どちらも空の行を除去

3. ユニーク抽出
   └── (区分, 材料) のユニークペアを取得

4. スコア計算（各ユニークペアに対して）
   ├── 変換リストの全行とのペア比較
   ├── score = max(類似度(材料, 変換前), 類似度(材料, 変換後))
   ├── UR案件 かつ UR品 → +10ボーナス（上限100）
   └── スコア降順で上位3件を選択

5. 出力
   └── DataFrame: 元の材料名, 区分, 候補1, スコア1, 候補2, スコア2, 候補3, スコア3
```

### `正規化()` 詳細

```python
unicodedata.normalize("NFKC", text)  # 全角→半角相当
text.upper()                          # 大文字化
re.sub(r"\s+", "")                    # 空白除去
# 記号除去: - / ( ) [ ] { } 【】「」『』. , _
# 単位変換:
#   ㎡, 平米, ｍ２, Ｍ２, M² → M2
#   ㎜, ミリ, ｍｍ, ＭＭ     → MM
#   メートル, ｍ, Ｍ         → M（単位変換は部分置換のため注意）
```

### `類似度()` 詳細

```python
if 正規化(a) == 正規化(b): return 100       # 完全一致
if a in b or b in a:       return 92        # 部分一致
return SequenceMatcher(None, a, b).ratio() * 100  # 類似度
```

`SequenceMatcher.ratio()` = `2 * M / T`
- M: マッチングブロックの文字数合計
- T: 両文字列の長さの合計

VBA実装ではbigram重複率で近似（相対順位の一致を優先）。

---

## 移行対象外

### `python/班員合計引き取り予定(3)_initialization.py`

```python
import numpy, pandas, matplotlib, seaborn, statsmodels, excel
warnings.simplefilter('ignore')
excel.set_xl_scalar_conversion(...)
excel.set_xl_array_conversion(...)
```

Excel Python環境の初期化ボイラープレート。業務ロジックなし。
移行対象外（Python退役後はリポジトリに残置のみ）。

---

## 既存VBAとの正規化比較

| 項目 | Python `正規化()` | VBA `Normalize()`（Module2） |
|---|---|---|
| 全角→半角 | `unicodedata.normalize("NFKC")` | `StrConv(s, vbNarrow)` |
| 大文字化 | `upper()` | `LCase()`（小文字） |
| 空白除去 | `re.sub(r"\s+")` | `Replace(s, " ")` など個別 |
| 単位変換 | ㎡→M2、㎜→MM など | mlのみ（不完全） |

VBAの大文字・小文字は逆だが、正規化後の比較のみに使用するため実用上問題なし。
Module4では Python に合わせて大文字（`UCase`）を採用する。
