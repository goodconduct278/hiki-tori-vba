import pandas as pd
from difflib import SequenceMatcher
import unicodedata
import re

def 正規化(text):
    if pd.isna(text):
        return ""

    text = str(text)
    text = unicodedata.normalize("NFKC", text)
    text = text.upper()
    text = re.sub(r"\s+", "", text)

    remove_chars = [
        " ", "　", "-", "－", "ー", "ｰ", "―", "‐",
        "/", "／", "･", "・", "(", ")", "（", "）",
        "[", "]", "［", "］", "{", "}", "｛", "｝",
        "【", "】", "「", "」", "『", "』",
        ".", "．", ",", "，", "_", "＿"
    ]

    for ch in remove_chars:
        text = text.replace(ch, "")

    replacements = {
        "㎡": "M2",
        "平米": "M2",
        "ｍ２": "M2",
        "Ｍ２": "M2",
        "M²": "M2",
        "㎜": "MM",
        "ミリ": "MM",
        "ｍｍ": "MM",
        "ＭＭ": "MM",
        "メートル": "M",
        "ｍ": "M",
        "Ｍ": "M"
    }

    for before, after in replacements.items():
        text = text.replace(before, after)

    return text.strip()

def 類似度(a, b):
    a = 正規化(a)
    b = 正規化(b)

    if a == "" or b == "":
        return 0

    if a == b:
        return 100

    if a in b or b in a:
        return 92

    return SequenceMatcher(None, a, b).ratio() * 100

def UR案件か(value):
    if pd.isna(value):
        return False
    text = unicodedata.normalize("NFKC", str(value).strip().upper())
    return "UR" in text

def UR品か(value):
    if pd.isna(value):
        return False
    text = unicodedata.normalize("NFKC", str(value).strip().upper())
    return text in ["UR", "○", "〇", "1", "TRUE", "YES", "対象"]

try:
    未変換 = xl(%P2%, headers=False)
except:
    未変換 = xl(%P3%, headers=False)

未変換 = 未変換.iloc[:, :7]
未変換.columns = ["Name", "担当者", "区分", "材料", "数量", "単位", "納品日"]

変換リスト = xl(%P4%, headers=False)
変換リスト = 変換リスト.iloc[:, :4]
変換リスト.columns = ["変換前（材料名）", "変換後（製品名）", "UR", "メーカー"]

未変換 = 未変換.dropna(subset=["材料"])
変換リスト = 変換リスト.dropna(
    subset=["変換前（材料名）", "変換後（製品名）"],
    how="all"
)

未変換材料リスト = (
    未変換[["区分", "材料"]]
    .dropna(subset=["材料"])
    .copy()
)

未変換材料リスト["材料"] = 未変換材料リスト["材料"].astype(str).str.strip()
未変換材料リスト = 未変換材料リスト[
    未変換材料リスト["材料"] != ""
].drop_duplicates()

結果 = []

for _, target in 未変換材料リスト.iterrows():
    区分 = target["区分"]
    材料 = target["材料"]
    is_ur_case = UR案件か(区分)

    scores = []

    for _, row in 変換リスト.iterrows():
        変換前 = row["変換前（材料名）"]
        変換後 = row["変換後（製品名）"]
        UR = row["UR"]
        メーカー = row["メーカー"]

        score_変換前 = 類似度(材料, 変換前)
        score_変換後 = 類似度(材料, 変換後)
        base_score = max(score_変換前, score_変換後)

        ur_bonus = 10 if is_ur_case and UR品か(UR) else 0
        score = min(base_score + ur_bonus, 100)

        scores.append({
            "候補製品名": 変換後,
            "スコア": score
        })

    上位候補 = sorted(scores, key=lambda x: x["スコア"], reverse=True)[:3]

    row = {
        "元の材料名": 材料,
        "区分": 区分,
        "候補1": "",
        "スコア1": "",
        "候補2": "",
        "スコア2": "",
        "候補3": "",
        "スコア3": ""
    }

    for idx, item in enumerate(上位候補, start=1):
        row[f"候補{idx}"] = item["候補製品名"]
        row[f"スコア{idx}"] = round(item["スコア"], 1)

    結果.append(row)

pd.DataFrame(結果)
