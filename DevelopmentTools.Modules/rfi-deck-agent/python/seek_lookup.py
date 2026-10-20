# -*- coding: utf-8 -*-
"""圖說查找：用 Revit 擷取的情境（棟別、樓層、柱線、元件）在「圖說查找索引」找對應的 PDF 圖說，並裁切到問題位置

  python seek_lookup.py --item 收件匣/xxx/item.json --out 輸出資料夾 [--top 4]
  python seek_lookup.py --building 實驗棟 --level 3F --grid "LINE(4-5, B-C)" --cats 結構構架 --out out

輸出：out/seek_result.json ＋ out/cand_N.png（裁切圖，紅框＝柱線範圍、紅圈＝物件位置；有結構柱時自動校正圖說偏移）
另外回傳 history：釋疑問題庫中同棟、同樓層、柱線 ±1 格內的舊問題。
索引與 PDF 位置：預設 03_資料庫/圖說/圖說查找索引.db，PDF 依索引在 03_資料庫/圖說（或其下一層）自動尋找；
別台電腦 PDF 放別處 → %APPDATA%/DT_AIAgent/seek.json 寫 {"pdf_dir": "...", "index": "..."}（詳見 agent_paths.py）。
只讀索引，不做 OCR，所以每台電腦結果相同（前提：PDF 版本相同，程式會比對檔案大小與特徵碼）。
"""
import argparse
import json
import os
import re
import sqlite3
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import agent_paths
ROOT = agent_paths.root()
DWG = os.path.join(ROOT, "03_資料庫", "圖說")
DEFAULT_INDEX = os.path.join(DWG, "圖說查找索引.db")

CAT_DISC = [  # Revit 類別關鍵字 → 圖號前綴（專業）
    (("結構", "基礎", "鋼筋"), ["S"]),
    (("風管", "空調", "機械設備", "風口"), ["AC"]),
    (("消防", "撒水", "灑水"), ["F"]),
    (("電纜", "電氣", "燈", "照明", "線槽", "開關", "插座", "管道"), ["E", "T"]),
    (("管", "衛浴", "給水", "排水", "衛工", "管件", "管配件"), ["P", "W"]),
    (("牆", "門", "窗", "樓板", "天花", "樓梯", "欄杆", "屋頂", "房間", "帷幕", "家具"), ["A"]),
]


def settings():
    return agent_paths.settings()


def norm_level(s):
    """Revit 樓層名 → 索引格式：B3F / 3F / RF / PRF"""
    t = (s or "").upper().replace("Ｆ", "F")
    if re.search(r"屋突|PR", t):
        return "PRF"
    m = re.search(r"B\s*(\d+)", t)
    if m:
        return f"B{int(m.group(1))}F"
    if re.search(r"\bR\d*F|RF|屋頂|屋面|\bR\d", t):
        return "RF"
    m = re.search(r"(\d+)\s*F", t) or re.search(r"(\d+)", t)
    if m:
        return f"{int(m.group(1))}F"
    m = re.search(r"地下([一二三四五])", s or "")
    if m:
        return f'B{"一二三四五".index(m.group(1)) + 1}F'
    return ""


def parse_grid(s):
    """'LINE(8-9, G'-H)' → [['8','9'], ["G'",'H']]"""
    m = re.search(r"\((.*)\)", s or "")
    body = m.group(1) if m else (s or "")
    out = []
    for part in body.split(","):
        labs = [x.strip() for x in re.split(r"[-~～]", part.strip()) if x.strip()]
        if labs:
            out.append(labs)
    return out


CAT_WORDS = [  # Revit 類別 → 圖名關鍵字加分
    (("風管", "風口"), "風管"), (("消防", "撒水", "灑水"), "撒水"), (("電纜", "線槽"), "電"),
    (("給水",), "給水"), (("排水", "衛浴"), "排水"), (("結構",), "結構"), (("天花",), "天花"),
]
CN = "一二三四五六七八九十"


def level_cn(lv):
    """B2F → 地下二層；3F → 三層"""
    m = re.match(r"^(B)?(\d+)F$", lv or "")
    if not m:
        return ""
    n = int(m.group(2))
    t = CN[n - 1] if n <= 10 else "十" + CN[n - 11]
    return ("地下" if m.group(1) else "") + t + "層"


def want_prefixes(cats):
    pre = []
    for c in cats:
        for keys, ps in CAT_DISC:
            if any(k in c for k in keys):
                pre += ps
                break
    return list(dict.fromkeys(pre))


MARK = re.compile(r"\b[A-Z]{1,4}\d{1,3}[A-Z]?\b")


BUILDING_BY_GRID = agent_paths.building_rules()   # 專案設定：seek.json 的 building_by_grid


def guess_building(gparts):
    for part in gparts:
        for lab in part:
            if lab.isdigit():
                for rg, b in BUILDING_BY_GRID:
                    if int(lab) in rg:
                        return b
    return ""


def find(con, building, level, grid, cats, types, keywords, top):
    pres = want_prefixes(cats) or ["A", "S"]
    if not building:
        building = guess_building(parse_grid(grid))
    lv = norm_level(level)
    gparts = parse_grid(grid)
    want_labels = {l for part in gparts for l in part}
    marks = sorted({m for t in types for m in MARK.findall((t or "").upper())})
    rows = con.execute("SELECT id,file,page,w,h,drawing_no,title,prefix,building,levels,kind,grid_v,grid_h FROM pages").fetchall()
    title_of = {r[0]: r[6] for r in rows}
    cands = []
    for r in rows:
        pid, f, pg, w, h, no, title, prefix, bld, levels, kind, gv, gh = r
        score, why = 0, []
        lvls = (levels or "").split(",")
        if prefix in pres:
            score += 30 if prefix == pres[0] else 22; why.append(f"專業 {prefix}")
        elif prefix in ("A",) and not cats:
            score += 10
        if lv and lv in lvls:
            score += 40; why.append(f"樓層 {lv}")
        elif lv and lvls and lvls != [""]:
            score -= 25
        if building and building in (bld or ""):
            score += 20; why.append(building)
        elif building and bld:
            score -= 15
        if kind == "平面圖":
            score += 15
        # 棟別＋樓層要連在一起（例：「實驗棟地下二層及行政棟地下一層」→ 實驗棟是 B2F）
        lc = level_cn(lv)
        if building and lc:
            if building + lc in (title or ""):
                score += 20; why.append(building + lc)
            elif building in (title or "") and lc in (title or ""):
                score -= 10
        for keys, word in CAT_WORDS:
            if any(k in c for c in cats for k in keys) and word in (title or ""):
                score += 10; why.append("圖名含「" + word + "」")
                break
        gv, gh = json.loads(gv or "[]"), json.loads(gh or "[]")
        labs = {g[0] for g in gv + gh}
        hit = want_labels & labs
        if want_labels and hit:
            score += 5 * len(hit); why.append("柱線 " + "、".join(sorted(hit)))
        cands.append([score, pid, why])
    # 構件編號（例：G3、SG1）→ 加分詳圖/大樣
    if marks:
        q = ",".join("?" * len(marks))
        for pid, t in con.execute(f"SELECT page_id, text FROM marks WHERE text IN ({q})", marks):
            for c in cands:
                if c[1] == pid:
                    c[0] += 4           # 元件類型名稱常與圖說不同，只當參考
                    if "構件 " + t not in c[2]:
                        c[2].append("構件 " + t)
    # 關鍵字全文
    for kw in keywords:
        kw = kw.strip()
        if len(kw) < 3:
            continue
        try:
            ids = {r[0] for r in con.execute("SELECT rowid FROM fts WHERE fts MATCH ?", ('"' + kw.replace('"', "") + '"',))}
        except sqlite3.OperationalError:
            ids = set()
        for c in cands:
            if c[1] in ids:
                c[0] += 8
                c[2].append("文字 " + kw)
    cands.sort(key=lambda c: -c[0])
    return [c for c in cands[:top] if c[0] > 0], gparts, marks


def grid_box(gv, gh, gparts):
    """柱線範圍 → 頁面比例框 (x0,y0,x1,y1)；找不到回傳 None"""
    vmap, hmap = {g[0]: g[1] for g in gv}, {g[0]: g[1] for g in gh}
    xs, ys = [], []
    for part in gparts:
        if all(l in vmap for l in part):
            xs = [vmap[l] for l in part]
        elif all(l in hmap for l in part):
            ys = [hmap[l] for l in part]
        else:
            # 部分找到也用
            xs = xs or [vmap[l] for l in part if l in vmap]
            ys = ys or [hmap[l] for l in part if l in hmap]
    if not xs and not ys:
        return None
    def span(vals, allv):
        if not vals:
            return None
        if len(vals) == 1:
            # 單一柱線：取相鄰柱線距離的一半當寬度
            pos = sorted(v for v in allv)
            gap = min([abs(b - a) for a, b in zip(pos, pos[1:])] or [0.05])
            return vals[0] - gap / 2, vals[0] + gap / 2
        return min(vals), max(vals)
    sx = span(xs, vmap.values()) or (0.05, 0.85)
    sy = span(ys, hmap.values()) or (0.05, 0.95)
    return sx[0], sy[0], sx[1], sy[1]


def _axis_map(gv, gh, part):
    """part 的柱線在哪個方向：回傳 (axis 'x'/'y', {label: 頁面比例})；找不到回傳 (None, {})"""
    vmap, hmap = {g[0]: g[1] for g in gv}, {g[0]: g[1] for g in gh}
    if all(l in vmap for l in part):
        return "x", vmap
    if all(l in hmap for l in part):
        return "y", hmap
    return None, {}


def grid_point(gv, gh, gparts, cellpos):
    """柱線格內相對位置 (fx,fy) → 頁面比例座標 (px,py)；任一方向算不出來就回傳 None"""
    try:
        fr = [float(v) for v in (cellpos or "").split(",")]
    except ValueError:
        return None
    pt = {}
    for k, part in enumerate(gparts[:2]):
        axis, m = _axis_map(gv, gh, part)
        if not axis:
            continue
        if len(part) == 1:
            pt[axis] = m[part[0]]
        elif k < len(fr):
            a, b = m[part[0]], m[part[-1]]
            pt[axis] = a + max(0.0, min(1.0, fr[k])) * (b - a)
    if "x" in pt and "y" in pt:
        return pt["x"], pt["y"]
    return None


def column_offset(page, gv, gh, gparts):
    """以結構柱校正：設計方 PDF 的柱線圓圈常與實際柱位錯開。
    在目標格附近的柱線交點找「實心近方形」（柱斷面），取偏移量中位數。
    回傳 (dx, dy, 命中數)（頁面比例）；找不到可信的偏移回傳 None"""
    if not gv or not gh:
        return None
    W, H = page.rect.width, page.rect.height
    xs = sorted(g[1] for g in gv)
    ys = sorted(g[1] for g in gh)
    def med(v):
        v = sorted(v) or [0.05]
        return v[len(v) // 2]
    # 用柱距中位數（輔助柱線 B' 之類距離很小，不能用最小值）
    sx = med([b - a for a, b in zip(xs, xs[1:]) if b - a > 1e-4])
    sy = med([b - a for a, b in zip(ys, ys[1:]) if b - a > 1e-4])
    # 目標附近的柱線（前後各 2 條）
    def near(glist, part):
        labs = [g[0] for g in sorted(glist, key=lambda g: g[1])]
        idx = [labs.index(l) for l in part if l in labs]
        if not idx:
            return [g[1] for g in glist]
        lo, hi = max(0, min(idx) - 2), min(len(labs) - 1, max(idx) + 2)
        pos = {g[0]: g[1] for g in glist}
        return [pos[labs[i]] for i in range(lo, hi + 1)]
    px, py = [g[1] for g in gv], [g[1] for g in gh]
    for part in gparts[:2]:
        axis, _ = _axis_map(gv, gh, part)
        if axis == "x":
            px = near(gv, part)
        elif axis == "y":
            py = near(gh, part)
    mat = page.rotation_matrix
    cols = []
    try:
        drawings = page.get_drawings()
    except Exception:
        return None
    for d in drawings:
        if d.get("fill") is None:
            continue
        r = d["rect"] * mat
        w, h = r.width / W, r.height / H
        if w <= 0 or h <= 0:
            continue
        if not (0.03 * sx <= w <= 0.35 * sx and 0.03 * sy <= h <= 0.35 * sy):
            continue
        if max(w / sx, h / sy) > 2.5 * min(w / sx, h / sy):
            continue
        cols.append(((r.x0 + r.x1) / 2 / W, (r.y0 + r.y1) / 2 / H, (w / sx) * (h / sy)))
    if len(cols) < 3:
        return None
    offs = []
    for x in px:
        for y in py:
            near_c = [(cx - x, cy - y, a) for cx, cy, a in cols if abs(cx - x) <= 0.45 * sx and abs(cy - y) <= 0.45 * sy]
            if not near_c:
                continue
            big = max(c[2] for c in near_c)
            near_c = [c for c in near_c if c[2] >= 0.5 * big]     # 主柱優先，略過小柱（C0a 之類）
            dx, dy, _ = min(near_c, key=lambda c: (c[0] / sx) ** 2 + (c[1] / sy) ** 2)
            offs.append((dx, dy))
    if len(offs) < 3:
        return None
    mdx = sorted(o[0] for o in offs)[len(offs) // 2]
    mdy = sorted(o[1] for o in offs)[len(offs) // 2]
    agree = [o for o in offs if abs(o[0] - mdx) <= 0.08 * sx and abs(o[1] - mdy) <= 0.08 * sy]
    if len(agree) < max(3, len(offs) * 0.5):
        return None
    dx = sum(o[0] for o in agree) / len(agree)
    dy = sum(o[1] for o in agree) / len(agree)
    if abs(dx) < 0.02 * sx and abs(dy) < 0.02 * sy:
        dx = dy = 0.0      # 本來就對齊
    return dx, dy, len(agree)


def render(pdf_path, page_no, box, out_png, max_px=1600, gv=None, gh=None, gparts=None, cellpos=""):
    """裁切圖說並畫紅框（柱線範圍）與紅圈（物件位置）。回傳 info dict"""
    import pymupdf
    from PIL import Image, ImageDraw
    doc = pymupdf.open(pdf_path)
    page = doc[page_no - 1]
    W, H = page.rect.width, page.rect.height
    info = {"snapped": False, "marker": None}
    mark = None
    if box and gv is not None:
        off = column_offset(page, gv, gh, gparts or [])
        info["column_check"] = [round(off[0], 4), round(off[1], 4), off[2]] if off else None
        if off and (off[0] or off[1]):
            box = (box[0] + off[0], box[1] + off[1], box[2] + off[0], box[3] + off[1])
            info["snapped"] = True
            info["snap_hits"] = off[2]
        mark = grid_point(gv, gh, gparts or [], cellpos)
        if mark and off:
            mark = (mark[0] + off[0], mark[1] + off[1])
        info["marker"] = mark
    if box:
        bw, bh = box[2] - box[0], box[3] - box[1]
        mx, my = max(bw * 0.6, 0.06), max(bh * 0.6, 0.06)
        view = [max(0, box[0] - mx), max(0, box[1] - my), min(0.88, box[2] + mx), min(1, box[3] + my)]
    else:
        view = [0, 0, 0.88, 1]          # 去掉右側圖框
    r = pymupdf.Rect(view[0] * W, view[1] * H, view[2] * W, view[3] * H)
    clip = r      # 顯示方向座標
    zoom = max_px / max(r.width, r.height)
    pix = page.get_pixmap(matrix=pymupdf.Matrix(zoom, zoom), clip=clip, alpha=False)
    img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
    if box:
        d = ImageDraw.Draw(img)
        sx, sy = img.width / (view[2] - view[0]), img.height / (view[3] - view[1])
        lw = max(3, img.width // 300)
        rect = [(box[0] - view[0]) * sx, (box[1] - view[1]) * sy, (box[2] - view[0]) * sx, (box[3] - view[1]) * sy]
        d.rectangle(rect, outline=(230, 0, 0), width=lw)
        if mark:
            cx, cy = (mark[0] - view[0]) * sx, (mark[1] - view[1]) * sy
            rad = max(18, min(rect[2] - rect[0], rect[3] - rect[1]) * 0.12)
            d.ellipse([cx - rad, cy - rad, cx + rad, cy + rad], outline=(230, 0, 0), width=lw + 1)
            d.line([cx - rad * 0.4, cy, cx + rad * 0.4, cy], fill=(230, 0, 0), width=max(2, lw - 1))
            d.line([cx, cy - rad * 0.4, cx, cy + rad * 0.4], fill=(230, 0, 0), width=max(2, lw - 1))
    img.save(out_png)
    doc.close()
    info["view"] = view
    return info


# ------------------------------------------------------------ 同位置的釋疑歷史
LETTER = re.compile(r"^([A-Z])('*)$")
GRID_TOKEN = re.compile(r"([A-Z]'*|\d{1,2})\s*[-~～]\s*([A-Z]'*|\d{1,2})(?:\s*[~～]\s*([A-Z]'*|\d{1,2})\s*[-]\s*([A-Z]'*|\d{1,2}))?")


def _lab_val(l):
    """柱線標示 → 可比較的數值：數字軸取數字；字母軸 A=1、B=2，B'=2.5"""
    l = l.strip().upper()
    if l.isdigit():
        return ("n", int(l))
    m = LETTER.match(l)
    if m:
        return ("a", ord(m.group(1)) - 64 + 0.5 * len(m.group(2)))
    return None


def _text_grid(text):
    """釋疑「位置」文字 → {'n': [數字...], 'a': [字母值...]}"""
    out = {"n": [], "a": []}
    t = (text or "").upper().replace("’", "'").replace("′", "'")
    for m in GRID_TOKEN.finditer(t):
        for g in m.groups():
            if g:
                v = _lab_val(g)
                if v:
                    out[v[0]].append(v[1])
    return out


LV_TXT = re.compile(r"(實驗棟|行政棟|歷史建築)?(地下)?([一二三四五六七八九十]+)[層樓]")


def _cn_num(t):
    """一 → 1、十 → 10、十二 → 12、二十 → 20"""
    if "十" in t:
        a, _, b = t.partition("十")
        return (CN.index(a) + 1 if a else 1) * 10 + (CN.index(b) + 1 if b else 0)
    return CN.index(t) + 1 if t in CN else 0


def _text_levels(text):
    """文字中的樓層 → {(棟別或None, 'B2F'), ...}"""
    out = set()
    t = text or ""
    for m in re.finditer(r"(?<![A-Z0-9])(B?\d{1,2})F", t.upper()):
        out.add((None, m.group(1) + "F"))
    for m in re.finditer(r"(?<!\d)(\d{1,2})樓", t):
        out.add((None, m.group(1) + "F"))
    if "筏基" in t:
        out.add((None, "筏基"))
    for m in LV_TXT.finditer(t):
        lv = ("B" if m.group(2) else "") + f"{_cn_num(m.group(3))}F"
        out.add((m.group(1), lv))
    if "實驗棟及行政棟" in t or "行政棟及實驗棟" in t:
        out = {(None, lv) for _, lv in out}
    return out


def find_history(building, level, gparts, limit=10):
    """在釋疑問題庫找同棟、同樓層、柱線相近（±1 格）的舊問題"""
    db = agent_paths.rfi_db_path(ROOT)
    if not gparts or not os.path.exists(db):
        return []
    try:
        import openpyxl
        wb = openpyxl.load_workbook(db, read_only=True, data_only=True)
    except Exception:
        return []
    ws = wb["問題清單"] if "問題清單" in wb.sheetnames else wb.worksheets[0]
    rows = ws.iter_rows(values_only=True)
    head = [str(h or "").strip() for h in next(rows, [])]
    col = {h: i for i, h in enumerate(head)}
    def g(r, k):
        i = col.get(k)
        v = r[i] if i is not None and i < len(r) else None
        return "" if v is None else str(v).strip()
    want = {"n": [], "a": []}
    for part in gparts:
        for l in part:
            v = _lab_val(l)
            if v:
                want[v[0]].append(v[1])
    if not want["n"] and not want["a"]:
        wb.close()
        return []
    lv = norm_level(level)
    out = []
    for r in rows:
        pos = g(r, "位置")
        title = g(r, "圖說編號名稱")
        q = g(r, "問題")
        if not pos and not q:
            continue
        got = _text_grid(pos)
        if not got["n"] and not got["a"]:
            continue
        ok, dist = True, 0.0
        for k in ("n", "a"):
            if want[k] and got[k]:
                dk = min(abs(a - b) for a in want[k] for b in got[k])
                dist += dk
                if dk > 1.01:
                    ok = False
            elif want[k] and not got[k] and k == "n":
                ok = False          # 沒有數字柱線就無法確認是哪一區
        if not ok:
            continue
        # 棟別
        txt = pos + " " + title
        other = [b for b in ("實驗棟", "行政棟") if b in pos and b != building]
        if building and other and building not in pos:
            continue
        # 樓層：有寫就必須相符
        lvls = _text_levels(pos) or _text_levels(title)      # 位置有寫樓層就以位置為準
        level_note = ""
        if lv and lvls:
            if not any(l == lv and (b is None or not building or b == building) for b, l in lvls):
                continue
        elif lv:
            level_note = "（未註明樓層）"
        tid = g(r, "追蹤編號") or "-".join(re.sub(r"\s+", "", x) for x in (g(r, "群組"), g(r, "編號")) if x)
        pos = pos.splitlines()[0].strip() if pos else pos
        state = g(r, "追蹤狀態") or ("已解決" if g(r, "是否解決") == "是" else (g(r, "狀態") or "未結"))
        concl = g(r, "定案結論")
        if not concl:
            hist = [x for x in g(r, "回覆歷程").splitlines() if x.strip()]
            concl = hist[-1] if hist else ""
        summ = re.sub(r"\s+", " ", q)[:40]
        out.append({"tid": tid, "state": state, "position": pos.replace("\n", " "), "drawing": title.replace("\n", " "),
                    "summary": summ + level_note, "conclusion": concl[:80],
                    "_d": dist + (0.5 if level_note else 0) + (0.3 if state == "已解決" else 0)})
    wb.close()
    out.sort(key=lambda h: h["_d"])
    for h in out:
        h.pop("_d")
    return out[:limit]


def main():
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    ap = argparse.ArgumentParser()
    ap.add_argument("--item")
    ap.add_argument("--building", default="")
    ap.add_argument("--level", default="")
    ap.add_argument("--grid", default="")
    ap.add_argument("--cellpos", default="", help="柱線格內相對位置 fx,fy（Revit 外掛自動帶入）")
    ap.add_argument("--cats", nargs="*", default=[])
    ap.add_argument("--types", nargs="*", default=[])
    ap.add_argument("--keywords", nargs="*", default=[])
    ap.add_argument("--out", required=True)
    ap.add_argument("--top", type=int, default=4)
    a = ap.parse_args()
    st = settings()
    index = st.get("index") or DEFAULT_INDEX
    pdf_dir = None
    res = {"ok": False}
    try:
        if a.item:
            it = json.load(open(a.item, encoding="utf-8-sig"))
            a.building = a.building or it.get("Building", "")
            a.level = a.level or it.get("Level", "")
            a.grid = a.grid or it.get("Grid", "")
            a.cellpos = a.cellpos or it.get("CellPos", "")
            els = it.get("Elements") or []
            a.cats = a.cats or [e.split(" ")[0] for e in els]
            a.types = a.types or els
            if not a.building:
                for b in ("實驗棟", "行政棟", "歷史建築"):
                    if b in (it.get("Location", "") + it.get("Question", "")):
                        a.building = b
        if not os.path.exists(index):
            raise SystemExit("找不到圖說查找索引：" + index)
        con = sqlite3.connect(index)
        meta = dict(con.execute("SELECT key,value FROM meta"))
        pdf_dir = agent_paths.pdf_dir(DWG, [r[0] for r in con.execute("SELECT file FROM files LIMIT 3")])
        # 檢查 PDF 版本
        warn = []
        for f, size in con.execute("SELECT file,size FROM files"):
            p = os.path.join(pdf_dir, f)
            if not os.path.exists(p):
                warn.append(f"缺少 {f}")
            elif os.path.getsize(p) != size:
                warn.append(f"{f} 與索引的版本不同")
        cands, gparts, marks = find(con, a.building, a.level, a.grid, a.cats, a.types, a.keywords, a.top)
        os.makedirs(a.out, exist_ok=True)
        out = []
        for k, (score, pid, why) in enumerate(cands, 1):
            f, pg, no, title, gv, gh, kind = con.execute(
                "SELECT file,page,drawing_no,title,grid_v,grid_h,kind FROM pages WHERE id=?", (pid,)).fetchone()
            # 只有平面圖才依柱線裁切；立面/剖面/詳圖給整張
            gvl, ghl = json.loads(gv or "[]"), json.loads(gh or "[]")
            box = grid_box(gvl, ghl, gparts) if (gparts and kind == "平面圖") else None
            png = os.path.join(a.out, f"cand_{k}.png")
            path = os.path.join(pdf_dir, f)
            info = {}
            if os.path.exists(path):
                info = render(path, pg, box, png, gv=gvl if box else None, gh=ghl, gparts=gparts, cellpos=a.cellpos)
            else:
                png = ""
            if info.get("snapped"):
                why.append("已依結構柱校正位置")
            out.append({"rank": k, "score": score, "why": why, "file": f, "page": pg, "drawing_no": no, "title": title,
                        "caption": f"{no} {title}".strip() + (f"（{a.grid}）" if box else ""),
                        "png": png, "located": bool(box), "marker": bool(info.get("marker")),
                        "snapped": bool(info.get("snapped")), "column_check": info.get("column_check")})
        history = []
        try:
            history = find_history(a.building or guess_building(gparts), a.level, gparts)
        except Exception as e:
            warn.append(f"讀取釋疑歷史失敗：{e}")
        res = {"ok": True, "query": {"building": a.building, "level": a.level, "level_norm": norm_level(a.level),
                                     "grid": a.grid, "cellpos": a.cellpos, "cats": a.cats, "marks": marks},
               "candidates": out, "history": history, "warnings": warn, "index_built": meta.get("built_at", ""),
               "index_builder": meta.get("builder", "")}
    except SystemExit as e:
        res["error"] = str(e)
    except Exception as e:
        import traceback
        res["error"] = f"{type(e).__name__}: {e}"
        res["trace"] = traceback.format_exc()[-1200:]
    os.makedirs(a.out, exist_ok=True)
    json.dump(res, open(os.path.join(a.out, "seek_result.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print("RESULT " + json.dumps(res, ensure_ascii=False))


if __name__ == "__main__":
    main()
