#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
fetch_ecms_status.py

從自己的 Notes 信箱找最新一封「ECMS P58 Current Online Status Report (yyyymmdd)」，
抽出 CSV 附件，依 EQPID 樣式篩選後，輸出「篩選後 CSV」與「HTML 表格片段」，
供機況表（ASPX 網頁）讀取顯示。設計成由 Windows 工作排程器每天跑 4 次。【唯讀，不修改信件】

前置需求：
  - Windows + 已開啟並登入的 Notes client
  - pip install pywin32   （Python 位元數必須跟 Notes client 一致，Notes 多半是 32 位元）
執行：
  python fetch_ecms_status.py          或   FetchEcmsStatus.exe（PyInstaller 打包，見 build_exe.bat）

設定：程式同資料夾若有 fetch_ecms_status.ini 就讀它覆蓋下方 DEFAULTS（見 fetch_ecms_status.ini.example）。

結束代碼：0 成功 / 3 信箱開不了 / 4 Notes COM 連不上(含 pywin32 未安裝) /
          5 找不到符合主旨的信 / 6 沒有 .csv 附件 / 7 CSV 無資料或找不到 EQPID 欄 / 8 寫檔失敗
"""
import configparser
import csv
import datetime as dt
import html
import io
import json
import logging
import os
import re
import sys

DEFAULTS = {
    # --- Notes 信箱 ---
    "server":         "UMCM74/UMC",
    "dbfile":         r"mail\00047829.nsf",
    "subject_prefix": "ECMS P58 Current Online Status Report",
    "lookback_days":  "2",
    # --- 篩選 ---
    "eqpid_column":   "EQPID",
    "eqpid_like":     "NISACVD-B%;SACVD-B%",     # ; 分隔，% 萬用字元，符合任一即保留
    # --- 輸出 ---
    "out_dir":        "",                        # 留空＝程式所在資料夾；通常填 ASPX 站台資料夾
    "out_csv":        "ecms_status.csv",
    "out_html":       "ecms_status.html",
    "keep_raw_copy":  "true",
    "csv_charset":    "auto",                    # auto = 依序試 utf-8-sig, big5, cp950, latin-1
    "log_file":       "fetch_ecms_status_log.txt",
}

EMBED_ATTACHMENT = 1454
RICHTEXT = 1


# ---------- 基礎 ----------

def app_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def load_config():
    cfg = dict(DEFAULTS)
    ini = os.path.join(app_dir(), "fetch_ecms_status.ini")
    if os.path.isfile(ini):
        # 記事本可能存成 ANSI(Big5)，依序試 UTF-8 / Big5 / cp950 / latin-1
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
        if cp.has_section("ecms"):
            for k, v in cp.items("ecms"):
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


def preview_rows():
    """命令列 --preview [N]：回傳要預覽的列數（預設 10），沒帶參數回 0"""
    args = sys.argv[1:]
    if "--preview" not in args:
        return 0
    i = args.index("--preview")
    if i + 1 < len(args) and args[i + 1].isdigit():
        return int(args[i + 1])
    return 10


def com_call(obj, name, *args):
    """呼叫 Notes COM 方法。拿不到型別資訊時（常見於跨位元），pywin32 會把方法誤判成屬性而回
    「找不到成員」；先 _FlagAsMethod 標成方法再呼叫就正常。"""
    try:
        obj._FlagAsMethod(name)
    except Exception:
        pass
    return getattr(obj, name)(*args)


def like_to_regex(pattern):
    """% 萬用字元 → regex（整串、不分大小寫）"""
    return re.compile("^" + ".*".join(re.escape(p) for p in pattern.split("%")) + "$", re.IGNORECASE)


# ---------- Notes helpers ----------

def item_text(doc, name):
    try:
        if not com_call(doc, "HasItem", name):
            return ""
        it = com_call(doc, "GetFirstItem", name)
        return (it.Text or "") if it is not None else ""
    except Exception:
        return ""


def mail_when(doc):
    """收信時間：DeliveredDate / PostedDate 文字能 parse 就用，否則 doc.Created"""
    for fld in ("DeliveredDate", "PostedDate"):
        s = item_text(doc, fld).strip()
        if s:
            for f in ("%Y/%m/%d %H:%M:%S", "%Y/%m/%d %H:%M", "%m/%d/%Y %H:%M:%S", "%m/%d/%Y %H:%M",
                      "%Y/%m/%d %p %I:%M:%S", "%Y/%m/%d %p %I:%M"):
                try:
                    return dt.datetime.strptime(s, f)
                except ValueError:
                    pass
    try:
        c = doc.Created
        return dt.datetime(c.year, c.month, c.day, c.hour, c.minute, c.second)
    except Exception:
        return dt.datetime.min


def _stream_bytes(stream):
    """把 NotesStream 整個讀成 bytes（COM 回來可能是 tuple[int] / bytes / memoryview）"""
    chunks = []
    while True:
        data = com_call(stream, "Read", 65536)
        if not data:
            break
        if isinstance(data, (bytes, bytearray)):
            chunks.append(bytes(data))
        elif isinstance(data, memoryview):
            chunks.append(data.tobytes())
        else:
            chunks.append(bytes(bytearray(int(b) & 0xFF for b in data)))
        if len(chunks[-1]) == 0:
            break
    return b"".join(chunks)


_FILENAME_RE = re.compile(r'(?:filename|name)\s*=\s*"?([^";]+)"?', re.IGNORECASE)


def _mime_filename(entity):
    """從 Content-Disposition / Content-Type 標頭取附件檔名；沒有回 "" """
    for hdr_name in ("Content-Disposition", "Content-Type"):
        try:
            hdr = com_call(entity, "GetNthHeader", hdr_name)
            if hdr is None:
                continue
            val = com_call(hdr, "GetHeaderValAndParams", False, True) or ""
            m = _FILENAME_RE.search(val)
            if m:
                return m.group(1).strip()
        except Exception:
            continue
    return ""


def _mime_walk(entity, out):
    """深度優先收集所有 MIME entity"""
    while entity is not None:
        out.append(entity)
        try:
            child = com_call(entity, "GetFirstChildEntity")
        except Exception:
            child = None
        if child is not None:
            _mime_walk(child, out)
        try:
            entity = com_call(entity, "GetNextSibling")
        except Exception:
            entity = None


def extract_csv_via_mime(ns, db, doc):
    """用 MIME 結構把 .csv 附件解碼到 NotesStream 再讀成 bytes（不落地）。非 MIME 信回 None。
    只在這段期間把 ConvertMIME 關掉並重新載入文件，結束一定還原，不影響其他路徑。"""
    try:
        ns.ConvertMIME = False
    except Exception:
        return None
    try:
        try:
            d2 = com_call(db, "GetDocumentByUNID", doc.UniversalID)
            root = com_call(d2, "GetMIMEEntity", "Body")
        except Exception as e:
            logging.info("GetMIMEEntity 失敗：%s", e)
            return None
        if root is None:
            logging.info("MIME：這封信不是 MIME 格式（GetMIMEEntity 回 None）")
            return None
        ents = []
        _mime_walk(root, ents)
        names = [_mime_filename(e) for e in ents]
        logging.info("MIME：共 %d 個 part，檔名：%s", len(ents), [n for n in names if n] or "（無）")
        return _mime_pick_csv(ns, ents, names)
    finally:
        try:
            ns.ConvertMIME = True
        except Exception:
            pass


def _mime_pick_csv(ns, ents, names):
    for ent, name in zip(ents, names):
        if not name.lower().endswith(".csv"):
            continue
        try:
            stream = com_call(ns, "CreateStream")
            com_call(ent, "GetContentAsBytes", stream, True)   # True = 已解碼（去 base64）
            try:
                stream.Position = 0
            except Exception:
                pass
            data = _stream_bytes(stream)
            try:
                com_call(stream, "Close")
            except Exception:
                pass
            if data:
                return (data, name, "MIME")
        except Exception as e:
            logging.warning("MIME 附件讀取失敗（%s）：%s", name, e)
    return None


def extract_first_csv(ns, db, doc, tmp_dir):
    """回傳 (bytes, name, how)；找不到回 (None, None, None)。
    順序：NotesEmbeddedObject.InputStream（記憶體）→ MIME GetContentAsBytes（記憶體）
          → ExtractFile 到 tmp_dir 讀後刪除（最後手段）。"""
    def try_eo(eo):
        try:
            if eo.Type != EMBED_ATTACHMENT:
                return None
            name = eo.Source or ""
            if not name.lower().endswith(".csv"):
                return None
        except Exception:
            return None

        # 方法 1：InputStream（記憶體）
        try:
            stream = eo.InputStream
            if stream is not None:
                try:
                    data = _stream_bytes(stream)
                finally:
                    try:
                        com_call(stream, "Close")
                    except Exception:
                        pass
                if data:
                    return (data, name, "InputStream")
        except Exception as e:
            logging.info("InputStream 不可用（%s），改試 MIME", e)

        # 方法 2：MIME（不落地）
        if not try_eo.mime_tried:
            try_eo.mime_tried = True
            r = extract_csv_via_mime(ns, db, doc)
            if r:
                return r

        # 方法 3：ExtractFile → 讀 bytes → 刪檔
        try:
            os.makedirs(tmp_dir, exist_ok=True)
            path = os.path.join(tmp_dir, name)
            if os.path.exists(path):
                os.remove(path)
            com_call(eo, "ExtractFile", path)
            with open(path, "rb") as f:
                data = f.read()
            try:
                os.remove(path)
                os.rmdir(tmp_dir)          # 用完就把空的 _tmp 收掉
            except OSError:
                pass
            return (data, name, "ExtractFile")
        except Exception as e:
            logging.warning("附件抽取失敗: %s", e)
            return None

    try_eo.mime_tried = False
    try:
        eos = doc.EmbeddedObjects
        if eos:
            for eo in eos:
                r = try_eo(eo)
                if r:
                    return r
    except Exception:
        pass
    try:
        for it in doc.Items:
            try:
                if it.Type != RICHTEXT:
                    continue
                reos = it.EmbeddedObjects
                if reos:
                    for eo in reos:
                        r = try_eo(eo)
                        if r:
                            return r
            except Exception:
                continue
    except Exception:
        pass
    # MIME 信件 EmbeddedObjects 常是空的：最後再用 MIME 找一次
    if not try_eo.mime_tried:
        r = extract_csv_via_mime(ns, db, doc)
        if r:
            return r
    return (None, None, None)


# ---------- CSV ----------

def decode_csv(raw, charset):
    tries = [charset] if charset and charset.lower() != "auto" else ["utf-8-sig", "big5", "cp950", "latin-1"]
    last = None
    for enc in tries:
        try:
            return raw.decode(enc), enc
        except UnicodeDecodeError as e:
            last = e
    raise last


def parse_rows(text):
    first = next((ln for ln in text.splitlines() if ln.strip()), "")
    delim = max([",", "\t", ";"], key=first.count) if first else ","
    rows = []
    for r in csv.reader(io.StringIO(text), delimiter=delim):
        cells = [c.strip() for c in r]
        if any(cells):
            rows.append(cells)
    return rows


# ---------- 輸出 ----------

def build_html(header, rows, subject, when, eqpid_like):
    esc = html.escape
    th = "<th style='border:1px solid #aac;background:#f2f5f9;color:#1f3b57;padding:4px 8px;text-align:left;'>{}</th>"
    td = "<td style='border:1px solid #aac;padding:3px 8px;'>{}</td>"
    out = ["<div class='ecms-status'>",
           "  <div class='ecms-meta' style='font-size:12px;color:#667;margin:4px 0;'>"
           f"來源：{esc(subject)}｜收信 {when:%Y/%m/%d %H:%M}｜更新 {dt.datetime.now():%Y/%m/%d %H:%M}"
           f"｜篩選 EQPID LIKE {esc(eqpid_like.replace(';', ' / '))}｜共 {len(rows)} 筆</div>",
           "  <table class='ecms-table' cellspacing='0' cellpadding='4' style='border-collapse:collapse;font-size:13px;'>",
           "    <thead><tr>" + "".join(th.format(esc(h)) for h in header) + "</tr></thead>",
           "    <tbody>"]
    for r in rows:
        cells = list(r) + [""] * (len(header) - len(r))
        out.append("      <tr>" + "".join(td.format(esc(c)) for c in cells[:len(header)]) + "</tr>")
    out += ["    </tbody>", "  </table>", "</div>", ""]
    return "\n".join(out)


# ---------- 主流程 ----------

def main():
    cfg = load_config()
    base = app_dir()
    out_dir = abs_path(base, cfg["out_dir"]) or base
    os.makedirs(out_dir, exist_ok=True)
    setup_logging(abs_path(base, cfg["log_file"]))

    try:
        import pythoncom  # noqa
        import pywintypes  # noqa
        import win32com.client  # noqa
    except ImportError as e:
        logging.error("載入 pywin32 失敗：%s", e)
        logging.error("Python=%s (%d-bit) frozen=%s", sys.executable, 64 if sys.maxsize > 2**32 else 32, getattr(sys, "frozen", False))
        logging.error("請用「打包時同一個 python」執行 pip install pywin32，並以 --hidden-import win32com.client 等參數打包（見 build_exe.bat）")
        return 4

    try:
        # 先試早期繫結（有型別庫時方法/屬性都分得清楚），不行再退回晚期繫結
        try:
            ns = win32com.client.gencache.EnsureDispatch("Notes.NotesSession")
        except Exception:
            ns = win32com.client.Dispatch("Notes.NotesSession")
    except Exception as e:
        logging.error("無法建立 Notes.NotesSession（Notes 未開/未登入？或 Python 位元數與 Notes 不符）: %s", e)
        return 4

    server, dbfile = cfg["server"], cfg["dbfile"]
    try:
        db = com_call(ns, "GetDatabase", server, dbfile)
        if not db.IsOpen:
            com_call(db, "Open", server, dbfile)
        if not db.IsOpen:
            raise RuntimeError("IsOpen=False")
    except Exception as e:
        logging.error("無法開啟信箱 [%s] %s: %s", server, dbfile, e)
        if "-2147352573" in str(e) or "找不到成員" in str(e) or "Member not found" in str(e):
            logging.error("Notes COM 回「找不到成員」：多半是 Python(%d-bit) 與 Notes client 位元數不同，請改用與 Notes 相同位元的 Python 打包。",
                          64 if sys.maxsize > 2**32 else 32)
        return 3

    lookback = int(cfg["lookback_days"] or 2)
    since = com_call(ns, "CreateDateTime", "Today")
    com_call(since, "AdjustDay", -lookback)
    prefix = cfg["subject_prefix"]
    formula = '@Begins(Subject; "%s")' % prefix.replace('"', '""')
    try:
        dc = com_call(db, "Search", formula, since, 0)
    except Exception as e:
        logging.error("搜尋信件失敗: %s", e)
        return 5
    if dc.Count == 0:
        logging.error("最近 %d 天內找不到主旨以「%s」開頭的信。", lookback, prefix)
        return 5

    best, best_when = None, None
    doc = com_call(dc, "GetFirstDocument")
    while doc is not None:
        w = mail_when(doc)
        if best is None or w > best_when:
            best, best_when = doc, w
        doc = com_call(dc, "GetNextDocument", doc)

    subject = item_text(best, "Subject")
    logging.info("找到 %d 封，取最新：%s（%s）", dc.Count, subject, best_when.strftime("%Y/%m/%d %H:%M"))

    raw_bytes, csv_name, how = extract_first_csv(ns, db, best, os.path.join(out_dir, "_tmp"))
    if raw_bytes is None:
        logging.error("這封信裡沒有 .csv 附件。")
        return 6
    logging.info("附件：%s（%d bytes，%s）", csv_name, len(raw_bytes), how)

    if cfg["keep_raw_copy"].lower() in ("1", "true", "yes"):
        raw_dir = os.path.join(out_dir, "raw")
        os.makedirs(raw_dir, exist_ok=True)
        try:
            with open(os.path.join(raw_dir, f"{best_when:%Y%m%d_%H%M}_{csv_name}"), "wb") as f:
                f.write(raw_bytes)
        except OSError as e:
            logging.warning("原始附件備份失敗: %s", e)

    try:
        text, enc = decode_csv(raw_bytes, cfg["csv_charset"])
    except Exception as e:
        logging.error("CSV 解碼失敗（試改 csv_charset）: %s", e)
        return 7
    rows = parse_rows(text)
    if len(rows) < 2:
        logging.error("CSV 沒有資料列（只有標題或是空檔）。")
        return 7
    header, data = rows[0], rows[1:]
    col = cfg["eqpid_column"].strip().upper()
    try:
        idx = [h.strip().upper() for h in header].index(col)
    except ValueError:
        logging.error("CSV 標題列找不到欄位「%s」。實際標題：%s", cfg["eqpid_column"], " | ".join(header))
        return 7

    pats = [like_to_regex(p.strip()) for p in cfg["eqpid_like"].split(";") if p.strip()]
    kept = [r for r in data if len(r) > idx and any(p.match(r[idx]) for p in pats)]

    # --preview [N]：把標題與篩選後前 N 列印到畫面/ log，方便不開檔確認抓到的內容
    n_prev = preview_rows()
    if n_prev:
        logging.info("預覽（篩選後前 %d 列，共 %d 列）：", min(n_prev, len(kept)), len(kept))
        logging.info("  %s", " | ".join(header))
        for r in kept[:n_prev]:
            logging.info("  %s", " | ".join(r))

    out_csv = os.path.join(out_dir, cfg["out_csv"])
    out_html = os.path.join(out_dir, cfg["out_html"])
    # 先寫 .tmp 再 rename，避免 ASPX 網頁讀到寫一半的檔
    try:
        with open(out_csv + ".tmp", "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            w.writerow(header)
            w.writerows(kept)
        with open(out_html + ".tmp", "w", encoding="utf-8") as f:
            f.write(build_html(header, kept, subject, best_when, cfg["eqpid_like"]))
        # 中繼資料：主旨 / 收信時間 / 抓取時間 / 筆數，給網頁面板顯示
        out_meta = os.path.join(out_dir, os.path.splitext(cfg["out_csv"])[0] + ".meta.json")
        meta = {
            "subject": subject,
            "mail_time": best_when.strftime("%Y-%m-%d %H:%M:%S"),
            "fetch_time": dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "attachment": csv_name,
            "rows_total": len(data),
            "rows_kept": len(kept),
            "eqpid_like": cfg["eqpid_like"],
        }
        with open(out_meta + ".tmp", "w", encoding="utf-8") as f:
            json.dump(meta, f, ensure_ascii=False, indent=2)
        os.replace(out_csv + ".tmp", out_csv)
        os.replace(out_html + ".tmp", out_html)
        os.replace(out_meta + ".tmp", out_meta)
    except OSError as e:
        logging.error("寫入輸出檔失敗（檔案可能被網站/Excel 鎖住）: %s", e)
        return 8

    logging.info("OK 原始 %d 筆 → 篩選後 %d 筆（編碼 %s）；已輸出 %s、%s、%s",
                 len(data), len(kept), enc, out_csv, out_html, out_meta)
    return 0


if __name__ == "__main__":
    sys.exit(main())
