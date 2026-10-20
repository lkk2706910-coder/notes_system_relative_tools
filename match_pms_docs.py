#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
match_pms_docs.py

讀一份 Excel（key 欄預設 `name`，值像 U20251203756），逐筆到 Notes `eri\\pms.nsf` 用 DOCNO 查對應文件，
抓 A17 / A21 / A25 / A26，依 A17 的「-」前綴分類（CUSILPE-B12 → CUSILPE），並附上文件的 Notes 連結
（即文件內容 → Meta → Identifier 的 notes:// URL），輸出成 xlsx：
  - 明細：原表格所有欄位 + A17/A21/A25/A26 + 機台類別 + ePMS Link + GR Link（可點）
    （GR Link：再到 UMCD14/UMC eop\\gr.nsf 用同一個 key 對 DCNUM 找到的文件連結）
  - 分類統計：每類筆數與 amt 合計
  - 每個類別一張工作表（可關）
【唯讀，不修改任何文件】

前置需求：Windows + 已開啟並登入的 Notes client；pip install pywin32 openpyxl
執行：MatchPmsDocs.exe [輸入.xlsx]            → 開視窗：選檔、查詢、逐列確認/修改分類/勾選、再匯出
      MatchPmsDocs.exe 輸入.xlsx --cli      → 不開視窗，直接查詢並匯出（排程用）
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
    # --- 第二個資料庫（GR）：用同一個 key 對 DCNUM，只取文件連結 ---
    "gr_server":     "UMCD14/UMC",
    "gr_dbfile":     r"eop\gr.nsf",
    "gr_view":       "",              # 第一欄為 DCNUM 的檢視（可用 ListNotesViews.vbs 查）；留空＝公式搜尋（較慢）
    "gr_key_field":  "DCNUM",
    "gr_fetch_fields": "",            # 也可順便抓 GR 文件欄位，逗號分隔；留空＝只要連結
    "gr_enabled":    "true",
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
    if sys.stdout is not None:
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
    """回傳 (輸入檔路徑或 "", 預覽列數或 0, 是否 --cli)。用法：[輸入.xlsx] [--preview [N]] [--cli]"""
    args = sys.argv[1:]
    preview = 0
    cli = "--cli" in args
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
    return (files[0] if files else ""), preview, cli


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
    # 工作表常有格式/篩選範圍延伸到很右邊的空欄，只保留到最後一個有名稱的欄位
    last = max((i for i, h in enumerate(header) if h), default=-1)
    if last < 0:
        raise ValueError("標題列全部空白")
    header = header[:last + 1]
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


def write_output(path, header, out_rows, extra_cols, cat_idx, link_idx, amt_idx, per_cat, link_cols=None):
    if link_cols is None:
        link_cols = [link_idx] if link_idx is not None else []
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
            for li in link_cols:
                if li < len(r) and r[li]:
                    cell = ws.cell(row=ws.max_row, column=li + 1)
                    cell.hyperlink = r[li]
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


# ---------- 查詢引擎（CLI 與 GUI 共用） ----------

class LookupError_(Exception):
    def __init__(self, code, msg):
        super().__init__(msg)
        self.code = code


class NotesLookup:
    """連 Notes、依 DOCNO 找文件。要在使用它的那條執行緒內建立（COM 綁執行緒）。"""

    def __init__(self, server, dbfile, view_name, key_field, fetch_fields, class_field="", ns=None):
        if ns is None:
            try:
                import pythoncom
                pythoncom.CoInitialize()
            except Exception:
                pass
            try:
                import win32com.client
                ns = win32com.client.Dispatch("Notes.NotesSession")
            except Exception as e:
                raise LookupError_(4, "無法建立 Notes.NotesSession（Notes 未開/未登入？pywin32 未裝？位元數不符？）: %s" % e)
        self.ns = ns
        try:
            self.db = com_call(self.ns, "GetDatabase", server, dbfile)
            if not self.db.IsOpen:
                com_call(self.db, "Open", server, dbfile)
            if not self.db.IsOpen:
                raise RuntimeError("IsOpen=False")
        except Exception as e:
            raise LookupError_(3, "無法開啟資料庫 [%s] %s: %s" % (server, dbfile, e))
        try:
            self.replica_id = self.db.ReplicaID
        except Exception:
            self.replica_id = ""
        self.view = None
        if (view_name or "").strip():
            try:
                self.view = com_call(self.db, "GetView", view_name)
                if self.view is None:
                    logging.warning("[%s] 找不到檢視 %s，改用公式搜尋（較慢）", dbfile, view_name)
            except Exception as e:
                logging.warning("[%s] 取得檢視失敗（%s），改用公式搜尋", dbfile, e)
        else:
            logging.info("[%s] 未設定檢視，用公式搜尋 %s", dbfile, key_field)
        self.server = server
        self.fields = [f.strip() for f in (fetch_fields or "").split(",") if f.strip()]
        self.class_field = (class_field or "").strip()
        self.docno_field = key_field

    def find(self, key):
        if self.view is not None:
            try:
                d = com_call(self.view, "GetDocumentByKey", key, True)
                if d is not None:
                    return d
            except Exception:
                pass
        formula = '%s = "%s"' % (self.docno_field, key.replace('"', '""'))
        try:
            dc = com_call(self.db, "Search", formula, None, 1)
            if dc.Count > 0:
                return com_call(dc, "GetFirstDocument")
        except Exception as e:
            logging.warning("公式搜尋失敗 %s: %s", key, e)
        return None

    def lookup_row(self, key):
        """回傳 (vals, cat, link, found)"""
        if not key:
            return [""] * len(self.fields), "(查無)", "", False
        d = self.find(key)
        if d is None:
            return [""] * len(self.fields), "(查無)", "", False
        vals = [item_text(d, f) for f in self.fields]
        cat = class_of(item_text(d, self.class_field)) if self.class_field else ""
        return vals, cat, doc_link(d, self.server, self.replica_id), True


def run_lookup(cfg, header, data, key_idx, progress=None):
    """對每列查 Notes。回傳 dict(out_rows, found, missing, extra_cols, cat_idx, link_idx, amt_idx)。
    progress(n, total, found) 可選。"""
    lk = NotesLookup(cfg["server"], cfg["dbfile"], cfg["docno_view"], cfg["docno_field"],
                     cfg["fetch_fields"], cfg["class_field"])
    gr = None
    if cfg["gr_enabled"].lower() in ("1", "true", "yes"):
        gr = NotesLookup(cfg["gr_server"], cfg["gr_dbfile"], cfg["gr_view"], cfg["gr_key_field"],
                         cfg["gr_fetch_fields"], "", ns=lk.ns)
    extra_cols = lk.fields + ["機台類別", "ePMS Link"]
    if gr is not None:
        extra_cols += ["GR " + f for f in gr.fields] + ["GR Link"]
    cat_idx = len(header) + len(lk.fields)
    link_idx = cat_idx + 1
    gr_found = 0
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
        vals, cat, link, ok = lk.lookup_row(key)
        if ok:
            found += 1
        elif key:
            missing.append(key)
        extra = vals + [cat, link]
        if gr is not None:
            gvals, _, glink, gok = gr.lookup_row(key)
            if gok:
                gr_found += 1
            extra += gvals + [glink]
        out_rows.append(list(row) + extra)
        if progress and (n % 10 == 0 or n == len(data)):
            progress(n, len(data), found)
    if gr is not None:
        logging.info("GR（%s）命中 %d / %d", cfg["gr_dbfile"], gr_found, len(data))
    link_cols = [i for i, c in enumerate(header + extra_cols) if c.endswith("Link")]
    return {"out_rows": out_rows, "found": found, "missing": missing, "extra_cols": extra_cols,
            "cat_idx": cat_idx, "link_idx": link_idx, "amt_idx": amt_idx, "link_cols": link_cols,
            "gr_found": gr_found}


# ---------- CLI ----------

def run_cli(cfg, base, in_path, preview):
    if not in_path or not os.path.isfile(in_path):
        logging.error("找不到輸入 Excel：%s（命令列給路徑，或在 ini 設 input_xlsx）", in_path or "(未指定)")
        return 6
    try:
        header, data, key_idx = read_input(in_path, cfg["sheet"], int(cfg["header_row"] or 1), cfg["key_column"])
    except Exception as e:
        logging.error("讀取 Excel 失敗：%s", e)
        return 7
    logging.info("輸入：%s，%d 列，key 欄「%s」", in_path, len(data), header[key_idx])

    def prog(n, total, found):
        if sys.stdout is not None:
            sys.stdout.write("\r  處理中… %d/%d，命中 %d" % (n, total, found))
            sys.stdout.flush()
    try:
        res = run_lookup(cfg, header, data, key_idx, prog)
    except LookupError_ as e:
        logging.error("%s", e)
        return e.code
    if sys.stdout is not None and data:
        sys.stdout.write("\n")

    out_rows = res["out_rows"]
    if preview:
        logging.info("預覽（前 %d 列）：", min(preview, len(out_rows)))
        logging.info("  %s", " | ".join(header + res["extra_cols"]))
        for r in out_rows[:preview]:
            logging.info("  %s", " | ".join("" if v is None else str(v) for v in r))

    out_path = abs_path(base, cfg["output_xlsx"]) or os.path.splitext(in_path)[0] + "_notes.xlsx"
    try:
        write_output(out_path, header, out_rows, res["extra_cols"], res["cat_idx"], res["link_idx"], res["amt_idx"],
                     cfg["sheet_per_category"].lower() in ("1", "true", "yes"), res["link_cols"])
    except Exception as e:
        logging.error("寫入 %s 失敗（檔案是否開著？）: %s", out_path, e)
        return 8
    logging.info("OK %d 列，ePMS 命中 %d，查無 %d，GR 命中 %d；輸出 %s", len(data), res["found"], len(res["missing"]), res["gr_found"], out_path)
    if res["missing"]:
        logging.info("查無的 key（前 20）：%s", ", ".join(res["missing"][:20]))
    return 0


# ---------- GUI ----------

def run_gui(cfg, base, in_path):
    import threading
    import queue
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox

    root = tk.Tk()
    root.title("PMS 文件比對")
    root.geometry("1280x760")

    def _tk_exc(exc, val, tb):
        import traceback
        logging.error("介面錯誤：%s", "".join(traceback.format_exception(exc, val, tb)))
        try:
            messagebox.showerror("介面錯誤", "%s\n（詳見 log）" % val)
        except Exception:
            pass
    root.report_callback_exception = _tk_exc

    state = {"header": [], "res": None, "include": [], "cat_filter": "（全部）", "out_path": ""}
    q = queue.Queue()

    # ---- 上方：檔案 / 查詢 ----
    top = ttk.Frame(root, padding=8)
    top.pack(fill="x")
    ttk.Label(top, text="輸入 Excel：").pack(side="left")
    path_var = tk.StringVar(value=in_path or "")
    ent = ttk.Entry(top, textvariable=path_var, width=70)
    ent.pack(side="left", padx=4)

    def browse():
        p = filedialog.askopenfilename(title="選擇 Excel", filetypes=[("Excel", "*.xlsx *.xlsm"), ("所有檔案", "*.*")])
        if p:
            path_var.set(p)
    ttk.Button(top, text="瀏覽…", command=browse).pack(side="left")
    btn_run = ttk.Button(top, text="查詢 Notes")
    btn_run.pack(side="left", padx=(12, 4))
    btn_export = ttk.Button(top, text="匯出 Excel…", state="disabled")
    btn_export.pack(side="left")
    btn_open = ttk.Button(top, text="開啟輸出檔", state="disabled")
    btn_open.pack(side="left", padx=4)

    bar = ttk.Progressbar(root, mode="determinate")
    bar.pack(fill="x", padx=8)
    status = ttk.Label(root, text="請選擇 Excel 後按「查詢 Notes」", padding=(8, 2))
    status.pack(fill="x")

    # ---- 中間：左表格 / 右統計 ----
    mid = ttk.PanedWindow(root, orient="horizontal")
    mid.pack(fill="both", expand=True, padx=8, pady=4)

    left = ttk.Frame(mid)
    mid.add(left, weight=4)
    filt_bar = ttk.Frame(left)
    filt_bar.pack(fill="x")
    ttk.Label(filt_bar, text="只看類別：").pack(side="left")
    cat_cb = ttk.Combobox(filt_bar, state="readonly", width=20, values=["（全部）"])
    cat_cb.current(0)
    cat_cb.pack(side="left")
    ttk.Label(filt_bar, text="   ✓ 欄點一下可切換是否匯出；雙擊「機台類別」可修改；雙擊 Link 開啟 Notes 文件",
              foreground="#667").pack(side="left")

    tv_frame = ttk.Frame(left)
    tv_frame.pack(fill="both", expand=True)
    tv = ttk.Treeview(tv_frame, show="headings", selectmode="browse")
    vsb = ttk.Scrollbar(tv_frame, orient="vertical", command=tv.yview)
    hsb = ttk.Scrollbar(tv_frame, orient="horizontal", command=tv.xview)
    tv.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
    tv.grid(row=0, column=0, sticky="nsew")
    vsb.grid(row=0, column=1, sticky="ns")
    hsb.grid(row=1, column=0, sticky="ew")
    tv_frame.rowconfigure(0, weight=1)
    tv_frame.columnconfigure(0, weight=1)
    tv.tag_configure("miss", foreground="#b91c1c")
    tv.tag_configure("off", foreground="#9ca3af")

    right = ttk.Frame(mid)
    mid.add(right, weight=1)
    ttk.Label(right, text="分類統計（勾選中的列）", font=("", 10, "bold")).pack(anchor="w")
    stat_tv = ttk.Treeview(right, columns=("cat", "n", "amt"), show="headings", height=20)
    for c, w, t in (("cat", 140, "機台類別"), ("n", 60, "筆數"), ("amt", 110, "amt 合計")):
        stat_tv.heading(c, text=t)
        stat_tv.column(c, width=w, anchor="e" if c != "cat" else "w")
    stat_tv.pack(fill="both", expand=True)

    # ---- 下方：log ----
    log_box = tk.Text(root, height=6, state="disabled", font=("Consolas", 9))
    log_box.pack(fill="x", padx=8, pady=(0, 8))

    class TkLogHandler(logging.Handler):
        def emit(self, record):
            q.put(("log", self.format(record)))
    h = TkLogHandler()
    h.setFormatter(logging.Formatter("[%(asctime)s] %(message)s", "%H:%M:%S"))
    logging.getLogger().addHandler(h)

    def log_line(text):
        log_box.configure(state="normal")
        log_box.insert("end", text + "\n")
        log_box.see("end")
        log_box.configure(state="disabled")

    # ---- 表格填充 / 統計 ----
    def all_cols():
        return ["✓"] + state["header"] + state["res"]["extra_cols"]

    def col_ids(cols):
        return ["c%d" % i for i in range(len(cols))]

    def refresh_stats():
        stat_tv.delete(*stat_tv.get_children())
        res = state["res"]
        if not res:
            return
        stats = {}
        for r, inc in zip(res["out_rows"], state["include"]):
            if not inc:
                continue
            s = stats.setdefault(r[res["cat_idx"]], [0, 0.0])
            s[0] += 1
            if res["amt_idx"] is not None:
                try:
                    s[1] += float(r[res["amt_idx"]] or 0)
                except (TypeError, ValueError):
                    pass
        tn, ta = 0, 0.0
        for cat in sorted(stats, key=lambda k: (-stats[k][0], k)):
            n, a = stats[cat]
            tn += n
            ta += a
            stat_tv.insert("", "end", values=(cat, n, "{:,.0f}".format(a)))
        stat_tv.insert("", "end", values=("總計", tn, "{:,.0f}".format(ta)))
        cats = ["（全部）"] + sorted(stats)
        cur = cat_cb.get()
        cat_cb["values"] = cats
        cat_cb.set(cur if cur in cats else "（全部）")

    def row_values(r, inc):
        return ["✓" if inc else ""] + ["" if v is None else str(v) for v in r]

    def fill_table():
        try:
            tv.delete(*tv.get_children())
            res = state["res"]
            if not res:
                return
            cols = all_cols()
            ids = col_ids(cols)
            tv["columns"] = ids
            for cid, name in zip(ids, cols):
                tv.heading(cid, text=name)
                tv.column(cid, width=40 if name == "✓" else (260 if name.endswith("Link") else 110),
                          anchor="center" if name == "✓" else "w", stretch=False)
            f = cat_cb.get()
            n = 0
            for idx, (r, inc) in enumerate(zip(res["out_rows"], state["include"])):
                cat = r[res["cat_idx"]]
                if f != "（全部）" and cat != f:
                    continue
                tags = []
                if cat == "(查無)":
                    tags.append("miss")
                if not inc:
                    tags.append("off")
                tv.insert("", "end", iid=str(idx), values=row_values(r, inc), tags=tags)
                n += 1
                if n % 200 == 0:
                    root.update_idletasks()
            logging.info("表格已填入 %d 列", n)
            refresh_stats()
        except Exception as e:
            import traceback
            logging.error("填表失敗：%s\n%s", e, traceback.format_exc())
            messagebox.showerror("填表失敗", "%s\n（詳見 log）" % e)

    cat_cb.bind("<<ComboboxSelected>>", lambda e: fill_table())

    def on_click(event):
        if tv.identify("region", event.x, event.y) != "cell":
            return
        col = tv.identify_column(event.x)
        iid = tv.identify_row(event.y)
        if not iid or col != "#1":
            return
        idx = int(iid)
        state["include"][idx] = not state["include"][idx]
        r = state["res"]["out_rows"][idx]
        tags = ["miss"] if r[state["res"]["cat_idx"]] == "(查無)" else []
        if not state["include"][idx]:
            tags.append("off")
        tv.item(iid, values=row_values(r, state["include"][idx]), tags=tags)
        refresh_stats()
    tv.bind("<Button-1>", on_click)

    def on_dblclick(event):
        if tv.identify("region", event.x, event.y) != "cell":
            return
        col = tv.identify_column(event.x)
        iid = tv.identify_row(event.y)
        if not iid:
            return
        ci = int(col[1:]) - 1
        cols = all_cols()
        res = state["res"]
        idx = int(iid)
        if cols[ci].endswith("Link"):
            link = res["out_rows"][idx][ci - 1]
            if link:
                try:
                    os.startfile(link)
                except Exception as e:
                    messagebox.showwarning("開啟失敗", str(e))
            return
        if cols[ci] == "機台類別":
            x, y, w, hgt = tv.bbox(iid, col)
            e = ttk.Entry(tv)
            e.place(x=x, y=y, width=w, height=hgt)
            e.insert(0, res["out_rows"][idx][res["cat_idx"]])
            e.focus_set()
            e.select_range(0, "end")

            def commit(_=None):
                res["out_rows"][idx][res["cat_idx"]] = e.get().strip() or "(空白)"
                e.destroy()
                fill_table()
            e.bind("<Return>", commit)
            e.bind("<FocusOut>", commit)
            e.bind("<Escape>", lambda _: e.destroy())
    tv.bind("<Double-1>", on_dblclick)

    # ---- 查詢（背景執行緒） ----
    def do_run():
        p = path_var.get().strip()
        if not p or not os.path.isfile(p):
            messagebox.showerror("找不到檔案", "請先選擇存在的 Excel 檔")
            return
        try:
            header, data, key_idx = read_input(p, cfg["sheet"], int(cfg["header_row"] or 1), cfg["key_column"])
        except Exception as e:
            messagebox.showerror("讀取 Excel 失敗", str(e))
            return
        state["header"] = header
        btn_run.configure(state="disabled")
        btn_export.configure(state="disabled")
        bar["maximum"] = max(len(data), 1)
        bar["value"] = 0
        status.configure(text="查詢中… 0/%d" % len(data))
        logging.info("輸入：%s，%d 列，key 欄「%s」", p, len(data), header[key_idx])

        def worker():
            try:
                res = run_lookup(cfg, header, data, key_idx, lambda n, t, f: q.put(("prog", n, t, f)))
                q.put(("done", res))
            except LookupError_ as e:
                q.put(("err", str(e)))
            except Exception as e:
                q.put(("err", "查詢發生錯誤：%s" % e))
        threading.Thread(target=worker, daemon=True).start()

    btn_run.configure(command=do_run)

    def poll():
        try:
            _poll_once()
        except Exception as e:
            import traceback
            logging.error("介面更新錯誤：%s\n%s", e, traceback.format_exc())
        root.after(100, poll)

    def _poll_once():
        try:
            while True:
                msg = q.get_nowait()
                kind = msg[0]
                if kind == "log":
                    log_line(msg[1])
                elif kind == "prog":
                    _, n, t, f = msg
                    bar["value"] = n
                    status.configure(text="查詢中… %d/%d，命中 %d" % (n, t, f))
                elif kind == "done":
                    res = msg[1]
                    state["res"] = res
                    state["include"] = [True] * len(res["out_rows"])
                    bar["value"] = bar["maximum"]
                    status.configure(text="完成：%d 列，命中 %d，查無 %d。請逐列確認後按「匯出 Excel…」"
                                     % (len(res["out_rows"]), res["found"], len(res["missing"])))
                    logging.info("查詢完成：%d 列，ePMS 命中 %d，查無 %d，GR 命中 %d", len(res["out_rows"]), res["found"], len(res["missing"]), res["gr_found"])
                    if res["missing"]:
                        logging.info("查無的 key（前 20）：%s", ", ".join(res["missing"][:20]))
                    fill_table()
                    btn_run.configure(state="normal")
                    btn_export.configure(state="normal")
                elif kind == "err":
                    status.configure(text="失敗：" + msg[1])
                    logging.error("%s", msg[1])
                    messagebox.showerror("查詢失敗", msg[1])
                    btn_run.configure(state="normal")
        except queue.Empty:
            pass
    root.after(100, poll)

    # ---- 匯出 ----
    def do_export():
        res = state["res"]
        if not res:
            return
        rows = [r for r, inc in zip(res["out_rows"], state["include"]) if inc]
        if not rows:
            messagebox.showwarning("沒有資料", "沒有勾選任何列")
            return
        default = abs_path(base, cfg["output_xlsx"]) or os.path.splitext(path_var.get().strip())[0] + "_notes.xlsx"
        p = filedialog.asksaveasfilename(title="匯出 Excel", defaultextension=".xlsx",
                                         initialfile=os.path.basename(default), initialdir=os.path.dirname(default),
                                         filetypes=[("Excel", "*.xlsx")])
        if not p:
            return
        try:
            write_output(p, state["header"], rows, res["extra_cols"], res["cat_idx"], res["link_idx"], res["amt_idx"],
                         cfg["sheet_per_category"].lower() in ("1", "true", "yes"), res["link_cols"])
        except Exception as e:
            messagebox.showerror("匯出失敗", "寫入 %s 失敗（檔案是否開著？）\n%s" % (p, e))
            return
        state["out_path"] = p
        btn_open.configure(state="normal")
        status.configure(text="已匯出 %d 列到 %s" % (len(rows), p))
        logging.info("已匯出 %d 列（共 %d 列，排除 %d）到 %s", len(rows), len(res["out_rows"]), len(res["out_rows"]) - len(rows), p)
        if messagebox.askyesno("匯出完成", "已匯出 %d 列。\n要現在開啟嗎？" % len(rows)):
            try:
                os.startfile(p)
            except Exception as e:
                messagebox.showwarning("開啟失敗", str(e))

    btn_export.configure(command=do_export)
    btn_open.configure(command=lambda: os.startfile(state["out_path"]) if state["out_path"] else None)

    if in_path and os.path.isfile(in_path):
        root.after(300, do_run)
    root.mainloop()
    return 0


# ---------- 主流程 ----------

def main():
    cfg = load_config()
    base = app_dir()
    cli_in, preview, cli = cli_input_and_preview()
    in_path = abs_path(base, cli_in or cfg["input_xlsx"])
    setup_logging(abs_path(base, cfg["log_file"]))
    if cli:
        return run_cli(cfg, base, in_path, preview)
    return run_gui(cfg, base, in_path)


if __name__ == "__main__":
    sys.exit(main())
