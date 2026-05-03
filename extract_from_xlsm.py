"""
xlsm から VBA / Power Query / Python を抽出してリポジトリのファイルを上書きするスクリプト。
使い方: python extract_from_xlsm.py
実行場所: excel_project_github_package/ 直下
"""

import base64
import re
import struct
import subprocess
import sys
import zipfile
import io
from pathlib import Path

XLSM_DIR = Path(__file__).parent / "エクセル見本"
VBA_DIR = Path(__file__).parent / "vba" / "_raw_vbaProject_bis"
PQ_DIR = Path(__file__).parent / "powerquery"
PY_DIR = Path(__file__).parent / "python"

XLSM_FILES = {
    "班員合計引き取り予定(3)": XLSM_DIR / "班員合計引き取り予定.xlsm",
    "ファイル作成マクロ":       XLSM_DIR / "ファイル作成マクロ.xlsm",
    "班員個人引き取り予定(1)":  XLSM_DIR / "班員個人引き取り予定.xlsm",
}

VBA_TYPE_DIR = {
    ".cls": {
        "ThisWorkbook": "ThisWorkbookモジュール",
        "Sheet":        "シートモジュール",
        "frm":          "フォームモジュール",
    },
    ".bas": "標準モジュール",
    ".frm": "フォームモジュール",
}


def vba_subdir(module_name: str, ext: str) -> str:
    if ext == ".bas":
        return "標準モジュール"
    if ext == ".frm":
        return "フォームモジュール"
    if module_name.startswith("ThisWorkbook"):
        return "ThisWorkbookモジュール"
    if re.match(r"^Sheet\d+$", module_name) or re.match(r"^frm", module_name):
        return "シートモジュール"
    return "クラスモジュール"


def extract_vba(xlsm_path: Path, prefix: str):
    result = subprocess.run(
        ["olevba", "--reveal", "--decode", str(xlsm_path)],
        capture_output=True, text=True, errors="replace"
    )
    output = result.stdout

    # 構造: 区切り線 → "VBA MACRO <file>" → "in file:..." → "- - - -" → コード → 区切り線
    primary_sep = re.compile(r"^-{10,}$", re.MULTILINE)
    secondary_sep = re.compile(r"^- - - - -", re.MULTILINE)

    # 主区切りで分割
    segments = primary_sep.split(output)

    for seg in segments:
        header_match = re.search(r"^VBA MACRO (\S+)", seg, re.MULTILINE)
        if not header_match:
            continue

        filename = header_match.group(1)

        # 副区切り以降がコード
        sec = secondary_sep.split(seg, maxsplit=1)
        code = sec[1].strip() if len(sec) > 1 else ""

        name = Path(filename).stem
        ext = Path(filename).suffix
        subdir = vba_subdir(name, ext)

        out_dir = VBA_DIR / subdir
        out_dir.mkdir(parents=True, exist_ok=True)
        out_path = out_dir / f"{prefix}{filename}"
        out_path.write_text(code, encoding="utf-8")
        print(f"  VBA: {out_path.relative_to(Path(__file__).parent)}")


def extract_powerquery(xlsm_path: Path, prefix: str):
    with zipfile.ZipFile(xlsm_path) as zf:
        names = zf.namelist()
        if "customXml/item1.xml" not in names:
            return

        raw_xml = zf.read("customXml/item1.xml")

    try:
        text = raw_xml.decode("utf-16")
    except Exception:
        text = raw_xml.decode("utf-8", errors="replace")

    match = re.search(r"<DataMashup[^>]*>(.*?)</DataMashup>", text, re.DOTALL)
    if not match:
        return

    raw = base64.b64decode(match.group(1).strip())
    zip_size = struct.unpack_from("<I", raw, 4)[0]
    zip_data = raw[8:8 + zip_size]

    try:
        zf2 = zipfile.ZipFile(io.BytesIO(zip_data))
    except Exception:
        return

    for name in zf2.namelist():
        content = zf2.read(name).decode("utf-8", errors="replace")
        safe_name = name.replace("/", "_").replace("\\", "_")
        out_path = PQ_DIR / f"{prefix}_{safe_name}"
        out_path.parent.mkdir(parents=True, exist_ok=True)
        out_path.write_text(content, encoding="utf-8")
        print(f"  PQ:  {out_path.relative_to(Path(__file__).parent)}")

    # connections.xml / queryTables もコピー
    with zipfile.ZipFile(xlsm_path) as zf:
        for entry in zf.namelist():
            if entry in ("xl/connections.xml",) or entry.startswith("xl/queryTables/"):
                content = zf.read(entry).decode("utf-8", errors="replace")
                safe = entry.replace("xl/", "").replace("/", "_")
                out_path = PQ_DIR / f"{prefix}__{safe}"
                out_path.write_text(content, encoding="utf-8")
                print(f"  PQ:  {out_path.relative_to(Path(__file__).parent)}")


def extract_python(xlsm_path: Path, prefix: str):
    with zipfile.ZipFile(xlsm_path) as zf:
        if "xl/python.xml" not in zf.namelist():
            return
        content = zf.read("xl/python.xml").decode("utf-8")

    # initialization
    init_match = re.search(
        r"<initialization[^>]*><code xml:space=\"preserve\">(.*?)</code></initialization>",
        content, re.DOTALL
    )
    if init_match:
        code = init_match.group(1)
        out_path = PY_DIR / f"{prefix}_initialization.py"
        out_path.write_text(code, encoding="utf-8")
        print(f"  PY:  {out_path.relative_to(Path(__file__).parent)}")

    # pythonScripts
    scripts = re.findall(
        r"<pythonScript><code xml:space=\"preserve\">(.*?)</code></pythonScript>",
        content, re.DOTALL
    )
    for i, code in enumerate(scripts, start=1):
        out_path = PY_DIR / f"{prefix}_python_script_{i}.py"
        out_path.write_text(code, encoding="utf-8")
        print(f"  PY:  {out_path.relative_to(Path(__file__).parent)}")


def main():
    PY_DIR.mkdir(parents=True, exist_ok=True)
    PQ_DIR.mkdir(parents=True, exist_ok=True)
    VBA_DIR.mkdir(parents=True, exist_ok=True)

    for prefix, xlsm_path in XLSM_FILES.items():
        if not xlsm_path.exists():
            print(f"[SKIP] {xlsm_path.name} が見つかりません")
            continue
        print(f"\n=== {xlsm_path.name} ===")
        extract_vba(xlsm_path, prefix)
        extract_powerquery(xlsm_path, prefix)
        extract_python(xlsm_path, prefix)

    print("\n完了")


if __name__ == "__main__":
    main()
