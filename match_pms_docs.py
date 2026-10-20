#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
match_pms_docs.py

讀一份 Excel（key 欄預設 `name`，值像 U20251203756），逐筆到 Notes `eri\\pms.nsf` 用 DOCNO 查對應文件，
抓 A17 / A21 / A25 / A26，依 A17 的「-」前綴分類（CUSILPE-B12 → CUSILPE），並附上文件的 Notes 連結
（即文件內容 → Meta → Identifier 的 notes:// URL），輸出成 xlsx：
  - 明細：原表格所有欄位 + A17/A21/A25/A26 + 機台類別 + Link（可點）
  - 分類統計：每類筆數與 amt 合計
  - 每個類別一張工作表（可關）
【唯讀，不修改任何文件】

前置需求：Windows + 已開啟並登入的 Notes client；pip install pywin32 openpyxl
執行：python match_pms_docs.py [輸入.xlsx] [--preview N]   或打包後 MatchPmsDocs.exe [輸入.xlsx]
設定：同資料夾 match_pms_docs.ini 覆蓋 DEFAULTS（見 match_pms_docs.ini.example）；命令列給的輸入檔優先。

結束代碼：0 成功 / 3 信箱開不了 / 4 Notes COM 連不上 / 6 找不到輸入 Excel /
          7 Excel 讀取失敗或找不到 key 欄 / 8 寫檔失敗
"""
import configparser
import datetime as dt
import logging
import os
import re
import sys

DEFAULTS = {
    # --- Notes ---
    "server":        "UMCD20/UMC",
    "dbfile":        r"eri\pms.nsf",
    "docno_view":    "(ByDocNo)",     # 第一欄為 DOCNO 的檢視；查不到會退回公式搜尋
    "docno_field":   "DOCNO",
    "fetch_fields":  "A17,A21,A25,A26",
    "class_field":   "A17",           # 用哪個欄位分類
    # --- Excel ---
    "input_xlsx":    "",              # 留空＝由命令列給；或填固定路徑
    "sheet":         "",              # 留空＝第一個工作表
    "key_column":    "name",          # 標題列裡的欄名（不分大小寫）
    "amt_column":    "amt",           # 分類統計要加總的欄；找不到就只算筆數
    "header_row":    "1",
    # --- 輸出 ---
    "output_xlsx":   "",              # 留空＝<輸入檔名>_notes.xlsx 放同資料夾
    "sheet_per_category": "true",
    "log_file":      "match_pms_docs_log.txt",
}


def app_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def load_config():
    cfg = dict(DEFAULTS)
    ini = os.path.join(app_dir(), "match_pms_docs.ini")
    if os.path.isfile(ini):
        raw = open(ini, "rb").read()
        text = None
        for enc in ("utf-8-sig", "big5", "cp950", "latin-1"):
            try:
                text = raw.decode(enc)
                break
            except UnicodeDecodeError:
                continue
        cp = configparser.ConfigParser(interpolation=None)
        try:
            cp.read_string(text or "")
        except configparser.Error as e:
            raise SystemExit("設定檔格式錯誤 %s：%s" % (ini, e))
        if cp.has_section("pms"):
            for k, v in cp.items("pms"):
                cfg[k.lower()] = v.strip()
    return cfg


def abs_path(base, p):
    if not p:
        return ""
    return p if os.path.isabs(p) else os.path.join(base, p)


def setup_logging(path):
    fmt = logging.Formatter("[%(asctime)s] %(levelname)s %(message)s", "%Y-%m-%d %H:%M:%S")
    root = logging.getLogger()
    root.setLevel(logging.INFO)
    sh = logging.StreamHandler(sys.stdout)
    sh.setFormatter(fmt)
    root.addHandler(sh)
    try:
        fh = logging.FileHandler(path, encoding="utf-8")
        fh.setFormatter(fmt)
        root.addHandler(fh)
    except OSError:
        pass


def cli_input_and_preview():
    """回傳 (輸入檔路徑或 "", 預覽列數或 0)。用法：[輸入.xlsx] [--preview [N]]"""
    args = sys.argv[1:]
    preview = 0
    files = []
    i = 0
    while i < len(args):
        a = args[i]
        if a == "--preview":
            if i + 1 < len(args) and args[i + 1].isdigit():
                preview = int(args[i + 1])
                i += 2
            else:
                preview = 10
                i += 1
            continue
        if not a.startswith("--"):
            files.append(a)
        i += 1
    return (files[0] if files else ""), preview


# ---------- Notes helpers ----------

def com_call(obj, name, *args):
    """Notes COM 方法：先 _FlagAsMethod 再呼叫（晚期繫結拿不到型別資訊時會把方法誤判成屬性）"""
    try:
        obj._FlagAsMethod(name)
    except Exception:
        pass
    return getattr(obj, name)(*args)


def item_text(doc, name):
    try:
        if not com_call(doc, "HasItem", name):
            return ""
        it = com_call(doc, "GetFirstItem", name)
        return (it.Text or "") if it is not None else ""
    except Exception:
        return ""


def doc_link(doc, server, replica_id):
    """文件內容 → Meta → Identifier 顯示的 notes:// URL"""
    try:
        url = doc.NotesURL
        if url:
            return url
    except Exception:
        pass
    try:
        return "notes://%s/%s/0/%s" % (server, replica_id, doc.UniversalID)
    except Exception:
        return ""


def class_of(value):
    """CUSILPE-B12 → CUSILPE；沒有 - 就整串；空值回 (空白)"""
    v = (value or "").strip()
    if not v:
        return "(空白)"
    return v.split("-", 1)[0].strip() or v


# ---------- Excel ----------

def read_input(path, sheet, header_row, key_col):
    from openpyxl import load_workbook
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb[sheet] if sheet else wb.worksheets[0]
    rows = list(ws.iter_rows(min_row=header_row, values_only=True))
    wb.close()
    if not rows:
        raise ValueError("工作表沒有資料")
    header = [("" if h is None else str(h)).strip() for h in rows[0]]
    try:
        key_idx = [h.lower() for h in header].index(key_col.strip().lower())
    except ValueError:
        raise ValueError("標題列找不到欄位「%s」。實際標題：%s" % (key_col, " | ".join(header)))
    data = []
    for r in rows[1:]:
        cells = list(r) + [None] * (len(header) - len(r))
        if all(c is None or str(c).strip() == "" for c in cells):
            continue
        data.append(cells[:len(header)])
    return header, data, key_idx


def write_output(path, header, out_rows, extra_cols, cat_idx, link_idx, amt_idx, per_cat):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill
    from openpyxl.utils import get_column_letter

    full_header = header + extra_cols
    bold = Font(bold=True)
    fill = PatternFill("solid", fgColor="F2F5F9")
    link_font = Font(color="0563C1", underline="single")

    def put_sheet(ws, rows):
        ws.append(full_header)
        for c in range(1, len(full_header) + 1):
            ws.cell(row=1, column=c).font = bold
            ws.cell(row=1, column=c).fill = fill
        for r in rows:
            ws.append(r)
            if link_idx is not None and r[link_idx]:
                cell = ws.cell(row=ws.max_row, column=link_idx + 1)
                cell.hyperlink = r[link_idx]
                cell.font = link_font
        for c, h in enumerate(full_header, 1):
            width = max([len(str(h))] + [len(str(x[c - 1])) for x in rows if x[c - 1] is not None][:200] + [4])
            ws.column_dimensions[get_column_letter(c)].width = min(max(width * 1.1, 8), 60)
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

    wb = Workbook()
    ws = wb.active
    ws.title = "明細"
    put_sheet(ws, out_rows)

    # 分類統計
    stats = {}
    for r in out_rows:
        cat = r[cat_idx]
        s = stats.setdefault(cat, {"n": 0, "amt": 0.0, "has_amt": False})
        s["n"] += 1
        if amt_idx is not None:
            try:
                s["amt"] += float(r[amt_idx] or 0)
                s["has_amt"] = True
            except (TypeError, ValueError):
                pass
    ws2 = wb.create_sheet("分類統計")
    ws2.append(["機台類別", "筆數", "amt 合計"] if amt_idx is not None else ["機台類別", "筆數"])
    for c in range(1, ws2.max_column + 1):
        ws2.cell(row=1, column=c).font = bold
        ws2.cell(row=1, column=c).fill = fill
    total_n, total_amt = 0, 0.0
    for cat in sorted(stats, key=lambda k: (-stats[k]["n"], k)):
        s = stats[cat]
        total_n += s["n"]
        total_amt += s["amt"]
        ws2.append([cat, s["n"], s["amt"]] if amt_idx is not None else [cat, s["n"]])
    ws2.append(["總計", total_n, total_amt] if amt_idx is not None else ["總計", total_n])
    for c in range(1, ws2.max_column + 1):
        ws2.cell(row=ws2.max_row, column=c).font = bold
    ws2.column_dimensions["A"].width = 18
    ws2.column_dimensions["B"].width = 8
    if amt_idx is not None:
        ws2.column_dimensions["C"].width = 14
        for row in ws2.iter_rows(min_row=2, min_col=3, max_col=3):
            row[0].number_format = "#,##0"

    if per_cat:
        for cat in sorted(stats, key=lambda k: (-stats[k]["n"], k)):
            title = re.sub(r"[\\/*?:\[\]]", "_", cat)[:31] or "_"
            if title in wb.sheetnames:
                title = (title[:28] + "_" + str(len(wb.sheetnames)))[:31]
            wsc = wb.create_sheet(title)
            put_sheet(wsc, [r for r in out_rows if r[cat_idx] == cat])

    wb.save(path)


# ---------- 主流程 ----------

def main():
    cfg = load_config()
    base = app_dir()
    cli_in, preview = cli_input_and_preview()
    in_path = abs_path(base, cli_in or cfg["input_xlsx"])
    setup_logging(abs_path(base, cfg["log_file"]))

    if not in_path or not os.path.isfile(in_path):
        logging.error("找不到輸入 Excel：%s（命令列給路徑，或在 ini 設 input_xlsx）", in_path or "(未指定)")
        return 6

    try:
        header, data, key_idx = read_input(in_path, cfg["sheet"], int(cfg["header_row"] or 1), cfg["key_column"])
    except Exception as e:
        logging.error("讀取 Excel 失敗：%s", e)
        return 7
    logging.info("輸入：%s，%d 列，key 欄「%s」", in_path, len(data), header[key_idx])

    try:
        import win32com.client
        ns = win32com.client.Dispatch("Notes.NotesSession")
    except Exception as e:
        logging.error("無法建立 Notes.NotesSession（Notes 未開/未登入？pywin32 未裝？位元數不符？）: %s", e)
        return 4

    server, dbfile = cfg["server"], cfg["dbfile"]
    try:
        db = com_call(ns, "GetDatabase", server, dbfile)
        if not db.IsOpen:
            com_call(db, "Open", server, dbfile)
        if not db.IsOpen:
            raise RuntimeError("IsOpen=False")
    except Exception as e:
        logging.error("無法開啟資料庫 [%s] %s: %s", server, dbfile, e)
        return 3
    try:
        replica_id = db.ReplicaID
    except Exception:
        replica_id = ""

    view = None
    try:
        view = com_call(db, "GetView", cfg["docno_view"])
        if view is None:
            logging.warning("找不到檢視 %s，改用公式搜尋（較慢）", cfg["docno_view"])
    except Exception as e:
        logging.warning("取得檢視失敗（%s），改用公式搜尋", e)
        view = None

    docno_field = cfg["docno_field"]
    fetch_fields = [f.strip() for f in cfg["fetch_fields"].split(",") if f.strip()]
    class_field = cfg["class_field"].strip()

    def find_doc(key):
        if view is not None:
            try:
                d = com_call(view, "GetDocumentByKey", key, True)
                if d is not None:
                    return d
            except Exception:
                pass
        formula = '%s = "%s"' % (docno_field, key.replace('"', '""'))
        try:
            dc = com_call(db, "Search", formula, None, 1)
            if dc.Count > 0:
                return com_call(dc, "GetFirstDocument")
        except Exception as e:
            logging.warning("公式搜尋失敗 %s: %s", key, e)
        return None

    extra_cols = fetch_fields + ["機台類別", "Link"]
    cat_idx = len(header) + len(fetch_fields)
    link_idx = cat_idx + 1
    amt_idx = None
    amt_name = cfg["amt_column"].strip().lower()
    if amt_name:
        for i, h in enumerate(header):
            if h.lower() == amt_name:
                amt_idx = i
                break

    out_rows, found, missing = [], 0, []
    for n, row in enumerate(data, 1):
        key = "" if row[key_idx] is None else str(row[key_idx]).strip()
        vals = [""] * len(fetch_fields)
        cat, link = "(查無)", ""
        if key:
            d = find_doc(key)
            if d is not None:
                found += 1
                vals = [item_text(d, f) for f in fetch_fields]
                cat = class_of(item_text(d, class_field))
                link = doc_link(d, server, replica_id)
            else:
                missing.append(key)
        out_rows.append(list(row) + vals + [cat, link])
        if n % 50 == 0:
            sys.stdout.write("\r  處理中… %d/%d，命中 %d" % (n, len(data), found))
            sys.stdout.flush()
    if len(data) >= 50:
        sys.stdout.write("\n")

    if preview:
        logging.info("預覽（前 %d 列）：", min(preview, len(out_rows)))
        logging.info("  %s", " | ".join(header + extra_cols))
        for r in out_rows[:preview]:
            logging.info("  %s", " | ".join("" if v is None else str(v) for v in r))

    out_path = abs_path(base, cfg["output_xlsx"]) or os.path.splitext(in_path)[0] + "_notes.xlsx"
    try:
        write_output(out_path, header, out_rows, extra_cols, cat_idx, link_idx, amt_idx,
                     cfg["sheet_per_category"].lower() in ("1", "true", "yes"))
    except Exception as e:
        logging.error("寫入 %s 失敗（檔案是否開著？）: %s", out_path, e)
        return 8

    logging.info("OK %d 列，命中 %d，查無 %d；輸出 %s", len(data), found, len(missing), out_path)
    if missing:
        logging.info("查無的 key（前 20）：%s", ", ".join(missing[:20]))
    return 0


if __name__ == "__main__":
    sys.exit(main())
