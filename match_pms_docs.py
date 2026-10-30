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
    "docno_view":    "(ByDocNo)",     # 第一欄為 DOCNO 的檢視；查不到會退回公式搜尋（只對 DOCNO 搜尋有用）
    "docno_field":   "DOCNO",
    # 搜尋依據：Excel key 欄要對 ePMS 的哪個欄位。可選 DOCNO / A21(material) / A22(原廠料號)
    "epms_key_field": "DOCNO",
    "epms_key_options": "DOCNO,A21,A22",
    # A21 / A22 / VendorNameList 一定會帶出（不在這裡也會自動補上）
    "fetch_fields":  "A17,A21,A22,A25,A26,VendorNameList",
    "class_field":   "A17",           # 用哪個欄位分類
    "epms_enabled":  "true",
    # 備援欄位：有些表單版本把資料存在 C 欄；主欄位空時改讀備援欄，用 A21/A22 搜尋時也一併搜備援欄。
    # 格式 主欄:備援欄,...  （對應關係請用「檢查欄位」核對後調整）
    "field_fallbacks": "A21:C11,A22:C12,A26:C15",
    # 用 A21/A22 搜尋時一個 key 可能對到多筆：all＝每筆各一列（Excel 原欄位重複）；first＝只取第一筆
    "multi_match":   "all",
    # --- 第二個資料庫（GR）：用同一個 key 對 DCNUM，只取文件連結 ---
    "gr_server":     "UMCD14/UMC",
    "gr_dbfile":     r"eop\gr.nsf",
    # 候選檢視（分號分隔）：啟動時用前幾個 key 逐一驗證「查到的文件 DCNUM == key」，第一個通過的採用；都不行退回公式搜尋
    "gr_view":       "(ByDocNo);(AllByDocNoOnly);(DocNo);(AllByGRNo);(AllByGRNoV2)",
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
    """回傳 (輸入檔路徑或 "", 預覽列數或 0, 是否 --cli, opts)。
    用法：[輸入.xlsx] [--preview [N]] [--cli] [--key DOCNO|A21|A22] [--no-epms] [--no-gr]"""
    args = sys.argv[1:]
    preview = 0
    cli = "--cli" in args
    opts = {"key": None, "epms": None, "gr": None, "dump": None, "keys": None}
    files = []
    i = 0
    while i < len(args):
        a = args[i]
        if a == "--key" and i + 1 < len(args):
            opts["key"] = args[i + 1]
            i += 2
            continue
        if a == "--keys" and i + 1 < len(args):
            opts["keys"] = args[i + 1]
            i += 2
            continue
        if a == "--dump" and i + 1 < len(args):
            opts["dump"] = args[i + 1]
            i += 2
            continue
        if a == "--no-epms":
            opts["epms"] = False
            i += 1
            continue
        if a == "--no-gr":
            opts["gr"] = False
            i += 1
            continue
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
    return (files[0] if files else ""), preview, cli, opts


# ---------- Notes helpers ----------

def com_call(obj, name, *args):
    """Notes COM 方法：先 _FlagAsMethod 再呼叫（晚期繫結拿不到型別資訊時會把方法誤判成屬性）"""
    try:
        obj._FlagAsMethod(name)
    except Exception:
        pass
    return getattr(obj, name)(*args)


def item_text(doc, name):
    """讀欄位文字。Notes 文件可有多個同名 item（GetFirstItem 可能拿到空的那個），
    所以：先 GetFirstItem；空的話掃全部同名 item 取第一個非空；還是空再退回 GetItemValue。"""
    try:
        if not com_call(doc, "HasItem", name):
            return ""
    except Exception:
        return ""
    try:
        it = com_call(doc, "GetFirstItem", name)
        if it is not None:
            t = (it.Text or "").strip()
            if t:
                return t
    except Exception:
        pass
    try:
        for it in doc.Items:
            try:
                if (it.Name or "").lower() == name.lower():
                    t = (it.Text or "").strip()
                    if t:
                        return t
            except Exception:
                continue
    except Exception:
        pass
    try:
        v = com_call(doc, "GetItemValue", name)
        if isinstance(v, (list, tuple)):
            return "; ".join(str(x) for x in v if x not in (None, ""))
        return "" if v is None else str(v)
    except Exception:
        return ""


def parse_fallbacks(spec):
    """'A21:C11,A26:C15' -> {'A21': ['C11'], 'A26': ['C15']}（備援可多個，用 / 分隔：A21:C11/D11）"""
    m = {}
    for part in (spec or "").split(","):
        if ":" not in part:
            continue
        a, b = part.split(":", 1)
        a = a.strip()
        fbs = [x.strip() for x in b.split("/") if x.strip()]
        if a and fbs:
            m.setdefault(a.upper(), []).extend(fbs)
    return m


FALLBACKS = {}


def item_text_fb(doc, name):
    """主欄位有值就用；空的話依 FALLBACKS 讀備援欄"""
    t = item_text(doc, name)
    if t:
        return t
    for fb in FALLBACKS.get(name.upper(), []):
        t = item_text(doc, fb)
        if t:
            return t
    return ""


def dump_items(doc, maxlen=200):
    """列出文件所有 item：name | type | text（同名會重複出現，方便看出空 item 擋在前面）"""
    rows = []
    try:
        for it in doc.Items:
            try:
                t = (it.Text or "")
            except Exception:
                t = "<無法讀取>"
            if len(t) > maxlen:
                t = t[:maxlen] + "…"
            try:
                ty = it.Type
            except Exception:
                ty = "?"
            rows.append((it.Name, ty, t))
    except Exception as e:
        rows.append(("<Items 讀取失敗>", "", str(e)))
    return rows


def inspect_key(cfg, key, key_field=None, use_gr=False):
    """找一筆文件並回傳 (描述, [(name,type,text),...])；給「檢查欄位」用"""
    if use_gr:
        lk = NotesLookup(cfg["gr_server"], cfg["gr_dbfile"], cfg["gr_view"], cfg["gr_key_field"], "", "")
    else:
        kf = (key_field or cfg["epms_key_field"] or "DOCNO").strip()
        view = cfg["docno_view"] if kf.upper() == cfg["docno_field"].upper() else ""
        lk = NotesLookup(cfg["server"], cfg["dbfile"], view, kf, "", "")
    docs = lk.find_all(key)
    if not docs:
        return "%s 在 %s 查無（依 %s）" % (key, lk.dbfile, lk.docno_field), []
    d = docs[0]
    head = "%s 在 %s 找到 %d 筆（依 %s），顯示第 1 筆：%s" % (key, lk.dbfile, len(docs), lk.docno_field, doc_link(d, lk.server, lk.replica_id))
    return head, dump_items(d)


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


def ensure_fields(fields_csv, must):
    """確保 must 裡的欄位都在清單中（保留原順序、去重，缺的補在後面）"""
    fields, low = [], set()
    for f in (fields_csv or "").split(","):
        f = f.strip()
        if f and f.lower() not in low:
            fields.append(f)
            low.add(f.lower())
    for m in must:
        if m.lower() not in low:
            fields.append(m)
            low.add(m.lower())
    return ",".join(fields)


def class_of(value):
    """CUSILPE-B12 → CUSILPE；沒有 - 就整串；空值回 (空白)"""
    v = (value or "").strip()
    if not v:
        return "(空白)"
    return v.split("-", 1)[0].strip() or v


# ---------- Excel ----------

def parse_pasted_keys(text):
    """貼上的清單 → key 列表：以換行/逗號/分號/Tab/空白分隔，去空白去重保序；順便過濾掉看起來是標題的 'name'"""
    keys, seen = [], set()
    for tok in re.split(r"[\r\n,;\t ]+", text or ""):
        k = tok.strip().strip('"').strip("'")
        if not k or k.lower() == "name":
            continue
        if k.upper() not in seen:
            seen.add(k.upper())
            keys.append(k)
    return keys


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

    if cat_idx is None:          # 沒比對 ePMS 就沒有分類
        wb.save(path)
        return

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
        self.dbfile = dbfile
        self.view = None
        self.view_candidates = [v.strip() for v in (view_name or "").split(";") if v.strip()]
        self.view_checked = False
        self.server = server
        self.fields = [f.strip() for f in (fetch_fields or "").split(",") if f.strip()]
        self.class_field = (class_field or "").strip()
        self.docno_field = key_field

    def pick_view(self, sample_keys):
        """用樣本 key 驗證候選檢視：GetDocumentByKey 回來的文件 key 欄位要等於 key 才算。"""
        self.view_checked = True
        if not self.view_candidates:
            logging.info("[%s] 未設定檢視，用公式搜尋 %s", self.dbfile, self.docno_field)
            return
        samples = [k for k in sample_keys if k][:5]
        for name in self.view_candidates:
            try:
                v = com_call(self.db, "GetView", name)
            except Exception:
                v = None
            if v is None:
                logging.info("[%s] 檢視 %s 不存在，跳過", self.dbfile, name)
                continue
            if not samples:
                self.view = v
                logging.info("[%s] 採用檢視 %s（無樣本可驗證）", self.dbfile, name)
                return
            hit = 0
            for k in samples:
                try:
                    d = com_call(v, "GetDocumentByKey", k, True)
                except Exception:
                    d = None
                if d is not None and item_text_fb(d, self.docno_field).strip().upper() == k.upper():
                    hit += 1
            if hit > 0:
                self.view = v
                logging.info("[%s] 採用檢視 %s（樣本 %d/%d 命中）", self.dbfile, name, hit, len(samples))
                return
            logging.info("[%s] 檢視 %s 查不到樣本，跳過", self.dbfile, name)
        logging.warning("[%s] 候選檢視都不適用，改用公式搜尋 %s（較慢）", self.dbfile, self.docno_field)

    def key_fields(self):
        return [self.docno_field] + FALLBACKS.get(self.docno_field.upper(), [])

    def prefetch(self, keys, chunk=150):
        """沒有可用檢視時：用 @IsMember 一次把一批 key 的文件撈回來快取，整個庫只掃 ceil(n/chunk) 次。
        cache: key(upper) -> [doc, ...]（同 key 可多筆）"""
        self.cache = {}
        keys = sorted({k.strip() for k in keys if k and k.strip()})
        if not keys:
            return
        if not hasattr(self, "_since"):
            self._since = com_call(self.ns, "CreateDateTime", "01/01/1980")
        got = 0
        for i in range(0, len(keys), chunk):
            part = keys[i:i + chunk]
            lst = ":".join('"%s"' % k.replace('"', '""') for k in part)
            formula = " | ".join('@IsMember(@Text(%s); %s)' % (f, lst) for f in self.key_fields())
            try:
                dc = com_call(self.db, "Search", formula, self._since, 0)
                d = com_call(dc, "GetFirstDocument")
                while d is not None:
                    k = item_text_fb(d, self.docno_field).strip()
                    if k:
                        self.cache.setdefault(k.upper(), []).append(d)
                        got += 1
                    d = com_call(dc, "GetNextDocument", d)
            except Exception as e:
                logging.warning("[%s] 批次公式搜尋失敗（第 %d 批）：%s", self.dbfile, i // chunk + 1, e)
        logging.info("[%s] 批次公式搜尋：%d 個 key → 撈到 %d 筆", self.dbfile, len(keys), got)

    def find_all(self, key):
        """回傳符合 key 的所有文件（list）。"""
        if not self.view_checked:
            self.pick_view([key])
        cache = getattr(self, "cache", None)
        if cache is not None:
            return list(cache.get(key.strip().upper(), []))
        if self.view is not None:
            try:
                dc = com_call(self.view, "GetAllDocumentsByKey", key, True)
                docs = []
                d = com_call(dc, "GetFirstDocument")
                while d is not None:
                    docs.append(d)
                    d = com_call(dc, "GetNextDocument", d)
                if docs:
                    return docs
            except Exception:
                pass
        kq = key.replace('"', '""')
        formula = " | ".join('@Text(%s) = "%s"' % (f, kq) for f in self.key_fields())
        try:
            # 第二參數必須是 NotesDateTime（pywin32 傳 None 會「類型不符」）：給一個很早的日期＝不限時間
            if not hasattr(self, "_since"):
                self._since = com_call(self.ns, "CreateDateTime", "01/01/1980")
            dc = com_call(self.db, "Search", formula, self._since, 0)
            docs = []
            d = com_call(dc, "GetFirstDocument")
            while d is not None:
                docs.append(d)
                d = com_call(dc, "GetNextDocument", d)
            return docs
        except Exception as e:
            logging.warning("公式搜尋失敗 %s: %s", key, e)
        return []

    def find(self, key):
        docs = self.find_all(key)
        return docs[0] if docs else None

    def row_from_doc(self, d):
        vals = [item_text_fb(d, f) for f in self.fields]
        cat = class_of(item_text_fb(d, self.class_field)) if self.class_field else ""
        return vals, cat, doc_link(d, self.server, self.replica_id)

    def lookup_rows(self, key, multi="all"):
        """回傳 [(vals, cat, link), ...]；查無回 []"""
        if not key:
            return []
        docs = self.find_all(key)
        if not docs:
            return []
        if multi != "all":
            docs = docs[:1]
        return [self.row_from_doc(d) for d in docs]

    def lookup_row(self, key):
        """回傳 (vals, cat, link, found)（單筆相容介面）"""
        rows = self.lookup_rows(key, "first")
        if not rows:
            return [""] * len(self.fields), "(查無)", "", False
        vals, cat, link = rows[0]
        return vals, cat, link, True


def run_lookup(cfg, header, data, key_idx, progress=None, epms_on=None, gr_on=None, key_field=None):
    """對每列查 Notes。回傳 dict(out_rows, found, missing, extra_cols, cat_idx, link_idx, amt_idx, ...)。
    epms_on / gr_on：是否比對 ePMS / 驗收系統(GR)；key_field：ePMS 用哪個欄位對 Excel key（DOCNO/A21/A22）。
    progress(n, total, found) 可選。"""
    truthy = lambda v: str(v).lower() in ("1", "true", "yes")
    if epms_on is None:
        epms_on = truthy(cfg["epms_enabled"])
    if gr_on is None:
        gr_on = truthy(cfg["gr_enabled"])
    if not epms_on and not gr_on:
        raise LookupError_(2, "ePMS 與驗收系統至少要勾選一個")
    key_field = (key_field or cfg["epms_key_field"] or "DOCNO").strip()
    multi = (cfg.get("multi_match") or "all").strip().lower()

    ns = None
    lk = gr = None
    if epms_on:
        # 搜尋欄位不是 DOCNO 時，(ByDocNo) 檢視沒用，直接走批次公式；且把 DOCNO 一起帶出來好對照
        view = cfg["docno_view"] if key_field.upper() == cfg["docno_field"].upper() else ""
        fields_csv = ensure_fields(cfg["fetch_fields"], ["A21", "A22", "VendorNameList"])
        if key_field.upper() != cfg["docno_field"].upper():
            # 用 A21/A22 搜尋時把 DOCNO 放在最前面，方便對照是哪張單
            fields_csv = ensure_fields(cfg["docno_field"] + "," + fields_csv, [])
        lk = NotesLookup(cfg["server"], cfg["dbfile"], view, key_field, fields_csv, cfg["class_field"])
        ns = lk.ns
    if gr_on:
        gr = NotesLookup(cfg["gr_server"], cfg["gr_dbfile"], cfg["gr_view"], cfg["gr_key_field"],
                         cfg["gr_fetch_fields"], "", ns=ns)
        ns = gr.ns

    extra_cols = []
    cat_idx = link_idx = None
    if lk is not None:
        extra_cols += lk.fields + ["機台類別", "ePMS Link"]
        cat_idx = len(header) + len(lk.fields)
        link_idx = cat_idx + 1
    if gr is not None:
        extra_cols += ["GR " + f for f in gr.fields] + ["GR Link"]
    gr_found = 0
    logging.info("比對範圍：ePMS=%s（依 %s）  驗收系統=%s", "是" if epms_on else "否", key_field if epms_on else "-", "是" if gr_on else "否")
    if FALLBACKS:
        logging.info("備援欄位：%s", ", ".join("%s←%s" % (k, "/".join(v)) for k, v in FALLBACKS.items()))
    amt_idx = None
    amt_name = cfg["amt_column"].strip().lower()
    if amt_name:
        for i, h in enumerate(header):
            if h.lower() == amt_name:
                amt_idx = i
                break
    all_keys = [str(r[key_idx]).strip() for r in data if r[key_idx] is not None and str(r[key_idx]).strip()]
    sample_keys = all_keys[:5]
    for lookup in [x for x in (lk, gr) if x is not None]:
        lookup.pick_view(sample_keys)
        if lookup.view is None:
            if progress:
                progress(0, len(data), 0)
            lookup.prefetch(all_keys)
    out_rows, found, missing, multi_keys = [], 0, [], 0
    for n, row in enumerate(data, 1):
        key = "" if row[key_idx] is None else str(row[key_idx]).strip()
        gr_part = []
        if gr is not None:
            gvals, _, glink, gok = gr.lookup_row(key)
            if gok:
                gr_found += 1
            gr_part = gvals + [glink]
        if lk is not None:
            hits = lk.lookup_rows(key, multi)
            if hits:
                found += 1
                if len(hits) > 1:
                    multi_keys += 1
                for vals, cat, link in hits:
                    out_rows.append(list(row) + vals + [cat, link] + gr_part)
            else:
                if key:
                    missing.append(key)
                out_rows.append(list(row) + [""] * len(lk.fields) + ["(查無)", ""] + gr_part)
        else:
            out_rows.append(list(row) + gr_part)
        if progress and (n % 10 == 0 or n == len(data)):
            progress(n, len(data), found if lk is not None else gr_found)
    if lk is not None and multi_keys:
        logging.info("ePMS：%d 個 key 對到多筆（multi_match=%s，輸出 %d 列）", multi_keys, multi, len(out_rows))
    if gr is not None:
        logging.info("GR（%s）命中 %d / %d", cfg["gr_dbfile"], gr_found, len(data))
    link_cols = [i for i, c in enumerate(header + extra_cols) if c.endswith("Link")]
    return {"out_rows": out_rows, "found": found, "missing": missing, "extra_cols": extra_cols,
            "cat_idx": cat_idx, "link_idx": link_idx, "amt_idx": amt_idx, "link_cols": link_cols,
            "gr_found": gr_found, "epms_on": lk is not None, "gr_on": gr is not None, "key_field": key_field}


# ---------- CLI ----------

def run_cli(cfg, base, in_path, preview, opts=None):
    opts = opts or {}
    if opts.get("keys"):
        spec = opts["keys"]
        if os.path.isfile(spec):
            spec = open(spec, "rb").read().decode("utf-8-sig", "replace")
        keys = parse_pasted_keys(spec)
        if not keys:
            logging.error("--keys 沒有解析到任何 key")
            return 6
        header, data, key_idx = [cfg["key_column"] or "name"], [[k] for k in keys], 0
        in_path = os.path.join(base, "keys_notes.xlsx")
        logging.info("輸入：--keys %d 個 key", len(data))
    else:
        if not in_path or not os.path.isfile(in_path):
            logging.error("找不到輸入 Excel：%s（命令列給路徑、--keys 清單，或在 ini 設 input_xlsx）", in_path or "(未指定)")
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
        res = run_lookup(cfg, header, data, key_idx, prog, opts.get("epms"), opts.get("gr"), opts.get("key"))
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

    out_path = abs_path(base, cfg["output_xlsx"]) or (in_path if opts.get("keys") else os.path.splitext(in_path)[0] + "_notes.xlsx")
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
    btn_inspect = ttk.Button(top, text="檢查欄位…")
    btn_inspect.pack(side="left", padx=(12, 0))

    # ---- 第二列：搜尋依據 / 比對哪些系統 ----
    opt = ttk.Frame(root, padding=(8, 0, 8, 4))
    opt.pack(fill="x")
    ttk.Label(opt, text="Excel key 欄對應：").pack(side="left")
    key_opts = [k.strip() for k in cfg["epms_key_options"].split(",") if k.strip()]
    key_labels = {"DOCNO": "DOCNO（單號）", "A21": "A21（material）", "A22": "A22（原廠料號）"}
    key_cb = ttk.Combobox(opt, state="readonly", width=18, values=[key_labels.get(k, k) for k in key_opts])
    init_key = (cfg["epms_key_field"] or "DOCNO").strip()
    key_cb.current(key_opts.index(init_key) if init_key in key_opts else 0)
    key_cb.pack(side="left", padx=(0, 16))
    truthy = lambda v: str(v).lower() in ("1", "true", "yes")
    epms_var = tk.BooleanVar(value=truthy(cfg["epms_enabled"]))
    gr_var = tk.BooleanVar(value=truthy(cfg["gr_enabled"]))
    ttk.Checkbutton(opt, text="比對 ePMS（eri\\pms.nsf）", variable=epms_var).pack(side="left")
    ttk.Checkbutton(opt, text="比對驗收系統（eop\\gr.nsf，依 DCNUM）", variable=gr_var).pack(side="left", padx=(12, 0))
    ttk.Label(opt, text="   A21 / A22 / VendorNameList 一定會帶出；A21/A22 搜尋對到多筆時每筆各一列", foreground="#667").pack(side="left")

    def selected_key():
        return key_opts[key_cb.current()]

    # ---- 輸入來源：Excel 檔 / 直接貼上 ----
    src_var = tk.StringVar(value="paste" if not (in_path and os.path.isfile(in_path)) else "excel")
    src = ttk.Frame(root, padding=(8, 0, 8, 4))
    src.pack(fill="x")
    ttk.Label(src, text="輸入來源：").pack(side="left")
    ttk.Radiobutton(src, text="Excel 檔（上方選檔）", variable=src_var, value="excel").pack(side="left")
    ttk.Radiobutton(src, text="直接貼上清單", variable=src_var, value="paste").pack(side="left", padx=(10, 0))
    ttk.Label(src, text="   貼上：一行一個或用逗號/空白分隔，可直接從 Excel 複製整欄貼進來；多筆一起查",
              foreground="#667").pack(side="left")
    paste_fr = ttk.Frame(root, padding=(8, 0, 8, 4))
    paste_box = tk.Text(paste_fr, height=4, font=("Consolas", 10), undo=True)
    paste_sb = ttk.Scrollbar(paste_fr, orient="vertical", command=paste_box.yview)
    paste_box.configure(yscrollcommand=paste_sb.set)
    paste_box.pack(side="left", fill="both", expand=True)
    paste_sb.pack(side="right", fill="y")
    paste_cnt = ttk.Label(root, text="", foreground="#667", padding=(8, 0))

    def refresh_src(*_):
        if src_var.get() == "paste":
            paste_fr.pack(fill="x", after=src)
            paste_cnt.pack(fill="x", after=paste_fr)
        else:
            paste_fr.pack_forget()
            paste_cnt.pack_forget()
    src_var.trace_add("write", refresh_src)

    def update_cnt(_=None):
        n = len(parse_pasted_keys(paste_box.get("1.0", "end")))
        paste_cnt.configure(text="已貼上 %d 個 key" % n if n else "")
    paste_box.bind("<KeyRelease>", update_cnt)
    paste_box.bind("<<Paste>>", lambda e: root.after(50, update_cnt))
    refresh_src()

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
        if not res or res["cat_idx"] is None:
            cat_cb["values"] = ["（全部）"]
            cat_cb.set("（全部）")
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
                cat = r[res["cat_idx"]] if res["cat_idx"] is not None else ""
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
        ci_ = state["res"]["cat_idx"]
        tags = ["miss"] if (ci_ is not None and r[ci_] == "(查無)") else []
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
        if src_var.get() == "paste":
            keys = parse_pasted_keys(paste_box.get("1.0", "end"))
            if not keys:
                messagebox.showerror("沒有資料", "請先在文字框貼上要查詢的 key（一行一個）")
                return
            p = "(貼上清單)"
            header, data, key_idx = [cfg["key_column"] or "name"], [[k] for k in keys], 0
        else:
            p = path_var.get().strip()
            if not p or not os.path.isfile(p):
                messagebox.showerror("找不到檔案", "請先選擇存在的 Excel 檔，或切到「直接貼上清單」")
                return
            try:
                header, data, key_idx = read_input(p, cfg["sheet"], int(cfg["header_row"] or 1), cfg["key_column"])
            except Exception as e:
                messagebox.showerror("讀取 Excel 失敗", str(e))
                return
        if not epms_var.get() and not gr_var.get():
            messagebox.showwarning("未選系統", "「比對 ePMS」與「比對驗收系統」至少要勾一個")
            return
        state["header"] = header
        btn_run.configure(state="disabled")
        btn_export.configure(state="disabled")
        bar["maximum"] = max(len(data), 1)
        bar["value"] = 0
        status.configure(text="查詢中… 0/%d" % len(data))
        logging.info("輸入：%s，%d 列，key 欄「%s」", p, len(data), header[key_idx])
        e_on, g_on, kf = epms_var.get(), gr_var.get(), selected_key()

        def worker():
            try:
                res = run_lookup(cfg, header, data, key_idx, lambda n, t, f: q.put(("prog", n, t, f)), e_on, g_on, kf)
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
                elif kind == "inspect":
                    show_inspect(msg[1], msg[2])
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
        if src_var.get() == "paste" or not path_var.get().strip():
            default = abs_path(base, cfg["output_xlsx"]) or os.path.join(base, "paste_%s_notes.xlsx" % __import__("datetime").datetime.now().strftime("%Y%m%d_%H%M"))
        else:
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

    def do_inspect():
        from tkinter import simpledialog
        key = simpledialog.askstring("檢查欄位", "輸入要檢查的 key（依目前「Excel key 欄對應」的欄位查 ePMS）：", parent=root)
        if not key or not key.strip():
            return
        key = key.strip()
        kf = selected_key()
        use_gr = (not epms_var.get()) and gr_var.get()

        def worker():
            try:
                head, rows = inspect_key(cfg, key, kf, use_gr)
                q.put(("inspect", head, rows))
            except Exception as e:
                q.put(("err", "檢查欄位失敗：%s" % e))
        threading.Thread(target=worker, daemon=True).start()

    def show_inspect(head, rows):
        win = tk.Toplevel(root)
        win.title("檢查欄位")
        win.geometry("900x600")
        ttk.Label(win, text=head, wraplength=880, padding=6).pack(fill="x")
        fr = ttk.Frame(win)
        fr.pack(fill="both", expand=True, padx=6, pady=(0, 6))
        t = ttk.Treeview(fr, columns=("n", "ty", "v"), show="headings")
        for c, w, h in (("n", 200, "item 名稱"), ("ty", 60, "type"), ("v", 580, "值（Text）")):
            t.heading(c, text=h)
            t.column(c, width=w, anchor="w")
        sb = ttk.Scrollbar(fr, orient="vertical", command=t.yview)
        t.configure(yscrollcommand=sb.set)
        t.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        t.tag_configure("empty", foreground="#9ca3af")
        for n, ty, v in rows:
            t.insert("", "end", values=(n, ty, v), tags=("empty",) if not v else ())
        ttk.Label(win, text="灰字＝該 item 是空的。同名 item 出現多個時，程式會取第一個非空的。表單上看得到但這裡完全沒有的欄位＝計算顯示欄，文件沒存。",
                  foreground="#667", padding=6, wraplength=880).pack(fill="x")

    btn_inspect.configure(command=do_inspect)
    btn_export.configure(command=do_export)
    btn_open.configure(command=lambda: os.startfile(state["out_path"]) if state["out_path"] else None)

    if in_path and os.path.isfile(in_path):
        root.after(300, do_run)
    root.mainloop()
    return 0


# ---------- 主流程 ----------

def main():
    global FALLBACKS
    cfg = load_config()
    FALLBACKS = parse_fallbacks(cfg.get("field_fallbacks", ""))
    base = app_dir()
    cli_in, preview, cli, opts = cli_input_and_preview()
    in_path = abs_path(base, cli_in or cfg["input_xlsx"])
    setup_logging(abs_path(base, cfg["log_file"]))
    if opts.get("dump"):
        try:
            head, rows = inspect_key(cfg, opts["dump"], opts.get("key"), opts.get("epms") is False)
        except LookupError_ as e:
            logging.error("%s", e)
            return e.code
        logging.info(head)
        for n, ty, t in rows:
            logging.info("  %-24s type=%-5s %s", n, ty, t)
        return 0
    if cli:
        return run_cli(cfg, base, in_path, preview, opts)
    return run_gui(cfg, base, in_path)


if __name__ == "__main__":
    sys.exit(main())
