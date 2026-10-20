# -*- coding: utf-8 -*-
"""AI AGENT 資料夾與專案設定的共用路徑規則。

程式可以放在兩種位置：
  1. AI AGENT 資料夾內的 01_AI運行本體/程式（原本的獨立使用方式）
  2. Revit 外掛解壓的位置 %LOCALAPPDATA%/DT_AIAgent/python/<版本>/（隨外掛更新）
第 2 種情況由外掛設定環境變數 DT_AIAGENT_ROOT 指向使用者的 AI AGENT 資料夾。

專案專屬設定寫在 %APPDATA%/DT_AIAgent/seek.json（不放在程式裡）：
  {
    "pdf_dir": "圖說 PDF 資料夾（省略時依索引自動尋找 03_資料庫/圖說 底下的資料夾）",
    "index": "圖說查找索引.db 路徑（省略時用 03_資料庫/圖說/圖說查找索引.db）",
    "rfi_db": "釋疑問題庫 xlsx 路徑（省略時在 03_資料庫/釋疑問題庫 找 *釋疑問題庫.xlsx）",
    "building_by_grid": [[7, 17, "實驗棟"], [1, 6, "行政棟"]]
  }
"""
import glob
import json
import os

HERE = os.path.dirname(os.path.abspath(__file__))


def root():
    env = os.environ.get("DT_AIAGENT_ROOT", "").strip()
    if env:
        return os.path.normpath(env)
    return os.path.normpath(os.path.join(HERE, "..", ".."))


def settings():
    p = os.path.join(os.environ.get("APPDATA", os.path.expanduser("~")), "DT_AIAgent", "seek.json")
    try:
        with open(p, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def rfi_db_path(base=None):
    """釋疑問題庫：seek.json 的 rfi_db → 資料夾內唯一的 *釋疑問題庫.xlsx → 釋疑問題庫.xlsx"""
    st = settings()
    if st.get("rfi_db"):
        return st["rfi_db"]
    d = os.path.join(base or root(), "03_資料庫", "釋疑問題庫")
    exact = os.path.join(d, "釋疑問題庫.xlsx")
    if os.path.exists(exact):
        return exact
    found = sorted(p for p in glob.glob(os.path.join(d, "*釋疑問題庫.xlsx")) if not os.path.basename(p).startswith("~$"))
    return found[0] if found else exact


def pdf_dir(dwg_dir, sample_files):
    """圖說 PDF 資料夾：seek.json 的 pdf_dir → 03_資料庫/圖說 或其下一層中，含有索引所列檔案的資料夾"""
    st = settings()
    if st.get("pdf_dir"):
        return st["pdf_dir"]
    sample = [f for f in sample_files if f][:3]
    cands = [dwg_dir]
    try:
        cands += sorted(os.path.join(dwg_dir, d) for d in os.listdir(dwg_dir) if os.path.isdir(os.path.join(dwg_dir, d)))
    except OSError:
        pass
    for c in cands:
        if sample and all(os.path.exists(os.path.join(c, f)) for f in sample):
            return c
    return dwg_dir


def building_rules():
    """柱線號碼 → 棟別，例如 [(range(7, 18), "實驗棟")]；未設定時不推測棟別"""
    rules = []
    for r in settings().get("building_by_grid") or []:
        try:
            lo, hi, name = int(r[0]), int(r[1]), str(r[2])
            rules.append((range(lo, hi + 1), name))
        except Exception:
            continue
    return rules
