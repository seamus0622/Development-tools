# -*- coding: utf-8 -*-
"""Revit「產生簡報」按鈕呼叫的程式（在使用者電腦上執行）

  python revit_deck.py --items 收件匣資料夾1 收件匣資料夾2 ... [--title 標題] [--date 115.10.16] [--no-db]

預設：選取的擷取 → 匯入釋疑問題庫（給追蹤編號）→ 只用這幾題產生釋疑簡報
--no-db：不寫入問題庫（測試或沒有問題庫的專案），直接用擷取內容做簡報，擷取移到 收件匣/_已產出簡報/
輸出：04_成果/簡報PPT/<民國日期>_<標題>_<時分>.pptx；最後一行印出 RESULT {json} 給外掛讀取。
產出的檔案同時記到 04_成果/簡報PPT/_Canva待上傳.txt，之後由 Claude 一次同步到 Canva。
"""
import argparse
import datetime as dt
import json
import os
import re
import shutil
import sys
import traceback
from types import SimpleNamespace

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import agent_paths
ROOT = agent_paths.root()
OUT_DIR = os.path.join(ROOT, "04_成果", "簡報PPT")
QUEUE = os.path.join(OUT_DIR, "_Canva待上傳.txt")


def roc_today():
    t = dt.date.today()
    return f"{t.year - 1911}.{t.month:02d}.{t.day:02d}"


def safe(s):
    return re.sub(r'[\\/:*?"<>|\s]+', "_", s).strip("_")[:40]


def from_items(folders, inbox):
    """不經問題庫，直接把擷取轉成 gen_rfi_deck 的 items"""
    items = []
    for k, f in enumerate(folders, 1):
        path = f if os.path.isabs(f) else os.path.join(inbox, f)
        it = json.load(open(os.path.join(path, "item.json"), encoding="utf-8-sig"))
        imgs = []
        for s in (it.get("Shots") or []):
            p = os.path.join(path, s.get("File", ""))
            if os.path.isfile(p):
                imgs.append({"path": p, "caption": s.get("Caption", "")})
        items.append({
            "no": f"Q{k:03d}", "group": "", "category": (it.get("Category") or "")[:1].upper(),
            "status": it.get("Status") or "預計提出", "drawing": it.get("Drawing", ""),
            "location": " ".join(x for x in [it.get("Building", ""), it.get("Location", "")] if x),
            "question": it.get("Question", ""), "suggest": it.get("Suggest", ""), "images": imgs,
            "note": "擷取：" + f'{it.get("Author", "")} {it.get("CreatedAt", "")}'.strip(),
            "_project": it.get("Project", ""),
        })
    return items


def main():
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    ap = argparse.ArgumentParser()
    ap.add_argument("--items", nargs="+", required=True)
    ap.add_argument("--title", default="")
    ap.add_argument("--date", default="")
    ap.add_argument("--no-db", action="store_true")
    ap.add_argument("--db", default="")
    a = ap.parse_args()

    from rfi_db import DB, DB_DEFAULT, cmd_import
    from gen_rfi_deck import load_input, filter_items, generate

    db_path = a.db or DB_DEFAULT
    inbox = os.path.join(os.path.dirname(db_path), "收件匣")
    date = a.date or roc_today()
    res = {"ok": False}
    try:
        if a.no_db:
            items = from_items(a.items, inbox)
            cfg = {"project": items[0].get("_project", "") if items else "", "title": a.title or "BIM釋疑-預計提出",
                   "subtitle": "", "date": date, "company": "", "show_stats": len(items) > 1,
                   "show_index": len(items) > 1, "items": items}
            done = os.path.join(inbox, "_已產出簡報")
            os.makedirs(done, exist_ok=True)
            moved = []
            for f in a.items:
                src = f if os.path.isabs(f) else os.path.join(inbox, f)
                moved.append(src)
            res["tids"] = []
        else:
            db = DB(db_path)
            tids = cmd_import(db, SimpleNamespace(group="", json_out=None, only=a.items))
            if not tids:
                raise SystemExit("選取的擷取都已經匯入過了（或找不到）。")
            cfg = load_input(db_path)
            filter_items(cfg, only=",".join(tids))
            cfg["date"] = date
            if a.title:
                cfg["title"] = a.title
            cfg["show_stats"] = len(tids) > 1
            cfg["show_index"] = len(tids) > 1
            res["tids"] = tids
            moved = []
        os.makedirs(OUT_DIR, exist_ok=True)
        stamp = dt.datetime.now().strftime("%H%M")
        name = f'{date.replace(".", "")}_{safe(cfg.get("title") or "釋疑簡報")}_{stamp}.pptx'
        out = os.path.join(OUT_DIR, name)
        generate(cfg, out)
        if a.no_db:
            # 簡報產生成功後才移走（圖片已嵌入簡報）
            for src in moved:
                if os.path.isdir(src):
                    dst = os.path.join(inbox, "_已產出簡報", os.path.basename(src))
                    if os.path.exists(dst):
                        dst += "_" + stamp
                    shutil.move(src, dst)
        with open(QUEUE, "a", encoding="utf-8") as f:
            f.write(f"{dt.datetime.now():%Y-%m-%d %H:%M}\t{out}\n")
        res.update(ok=True, pptx=out, count=len(cfg["items"]))
    except SystemExit as e:
        res["error"] = str(e)
    except Exception as e:
        res["error"] = f"{type(e).__name__}: {e}"
        res["trace"] = traceback.format_exc()[-1500:]
    print("RESULT " + json.dumps(res, ensure_ascii=False))
    sys.exit(0 if res.get("ok") else 1)


if __name__ == "__main__":
    main()
