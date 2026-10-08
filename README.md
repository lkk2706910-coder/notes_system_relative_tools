# notes_system_relative_tools

一組操作 HCL / IBM Lotus Notes（Domino）的 VBScript 小工具，透過 Notes COM 介面（`Notes.NotesSession`）做批次任務。所有腳本以 `cscript` 在 Windows 執行，且需先開啟並登入 Notes client。各腳本頂端都有 `設定區`，改常數即可。AI agent 工作指南見 [`AGENTS.md`](AGENTS.md)。

> 本機編輯 `.vbs` 後請存成 **UTF-16 LE with BOM**（記事本「Unicode」），cscript 才能正確讀中文。

---

## `FetchEcmsStatus.vbs` — 抓 ECMS 狀態信附件、篩選後更新機況表資料

從自己的 Notes 信箱找最近 `LOOKBACK_DAYS` 天內、主旨以 `ECMS P58 Current Online Status Report` 開頭的信，取**最新一封**，抽出 **CSV 附件**，依 `EQPID` 樣式篩選後輸出：

- `ecms_status.csv`：篩選後資料（UTF-8 + BOM，Excel 可直接開）
- `ecms_status.html`：HTML 表格片段（無 `<html>/<body>`，最上方一行 meta：來源主旨 / 收信時間 / 更新時間 / 篩選條件 / 筆數）
- `raw\yyyymmdd_hhmm_<原附件名>.csv`：原始附件備份（`KEEP_RAW_COPY=True` 時）

**唯讀，不修改信件。** 設計成由 Windows 工作排程器每天在 4 封信到後各跑一次。

**設定**

| 常數 | 說明 |
|------|------|
| `SERVER` / `DBFILE` | 信箱伺服器與信箱檔，例 `"UMCM74/UMC"` / `"mail\00047829.nsf"` |
| `SUBJECT_PREFIX` | 主旨開頭（`@Begins` 比對） |
| `LOOKBACK_DAYS` | 只在最近 N 天的信裡找 |
| `EQPID_COLUMN` | CSV 標題列中的 EQPID 欄名（不分大小寫） |
| `EQPID_LIKE` | 篩選樣式，多個用 `;` 分隔，`%` 為萬用字元；符合任一即保留。預設 `NISACVD-B%;SACVD-B%` |
| `OUT_DIR` | 輸出資料夾；留空＝本 `.vbs` 所在資料夾，通常填 ASPX 站台的資料夾 |
| `OUT_CSV` / `OUT_HTML` | 輸出檔名 |
| `KEEP_RAW_COPY` | `True`＝另存原始附件到 `OUT_DIR\raw\` |
| `CSV_CHARSET` | 附件 CSV 編碼，預設 `utf-8`；中文亂碼改 `big5` |

**執行**

```bat
cscript //nologo FetchEcmsStatus.vbs
```
或雙擊 / 排程 `fetch_ecms_status.bat`（輸出寫入同資料夾 `fetch_ecms_status_log.txt`）。

**排程（Windows 工作排程器）**：信在 08:00 / 15:30 / 20:30 / 23:30 到，建議各延後 10 分鐘跑。用系統管理員命令提示字元：

```bat
schtasks /Create /TN "ECMS Status 0810" /TR "\"C:\path\to\fetch_ecms_status.bat\"" /SC DAILY /ST 08:10 /F
schtasks /Create /TN "ECMS Status 1540" /TR "\"C:\path\to\fetch_ecms_status.bat\"" /SC DAILY /ST 15:40 /F
schtasks /Create /TN "ECMS Status 2040" /TR "\"C:\path\to\fetch_ecms_status.bat\"" /SC DAILY /ST 20:40 /F
schtasks /Create /TN "ECMS Status 2340" /TR "\"C:\path\to\fetch_ecms_status.bat\"" /SC DAILY /ST 23:40 /F
```
排程要在「使用者已登入、Notes client 開著」的 session 下執行（工作排程器選「只有在使用者登入時才執行」）。

---

## 機況表（`eq_status.aspx`）整合

你的看板每 `REFRESH_MS` 會自己 `fetch` 一次頁面、把 `.board` 的 `innerHTML` 整個換掉，
而 `phTable` 就在 `.board` 裡——所以**只要在 code-behind 把 ECMS 表格 append 進 `phTable`，就會跟著 PM 面板一起自動刷新**，不需要 iframe、不用再寫前端 JS。

現成程式碼在 [`aspx_integration/EcmsPanel.snippet.cs`](aspx_integration/EcmsPanel.snippet.cs)，做法完全比照既有的 PM 面板（`BuildPmHtml`）：依檔案修改時間快取、`Monitor.TryEnter` 不排隊、讀檔失敗顯示舊資料並掛 `data-warn`、沿用 `pm-panel / pm-table / pm-alert` class（深色模式自動套用）。

### 步驟

1. **輸出指到站台資料夾**（`fetch_ecms_status.ini`）：
   ```ini
   out_dir = C:\inetpub\wwwroot\<你的站台>\data
   ```
   權限：排程帳號對 `data\` 要「寫入」；IIS 應用程式集區帳號（`IIS AppPool\<集區名>`）要「讀取」。

2. **貼程式碼**到 `eq_status.aspx.cs`：
   - 把 snippet 的「區塊 A」整段貼進 `class EQ_Status`（例如放在 `BuildPmHtml` 後面）。
   - 把「區塊 B」這幾行貼進 `Page_Load`，位置在 PM 區塊 `if (PmEnabled()) { ... }` **之後**、`phTable.Controls.Clear();` **之前**：
     ```csharp
     string ecmsWarning;
     DateTime ecmsTime;
     List<string[]> ecmsRows = LoadEcmsCached(out ecmsWarning, out ecmsTime);
     html.Append(BuildEcmsHtml(ecmsRows, ecmsWarning, ecmsTime));
     ```
   - 既有 `using` 已足夠（`System.IO`、`System.Text`、`System.Threading`、`System.Globalization`、`System.Collections.Generic`）。

3. **（選用）`web.config` `<appSettings>`**，都有預設值：

   | key | 預設 | 說明 |
   |---|---|---|
   | `EcmsCsvPath` | `~/data/ecms_status.csv` | CSV 位置（相對站台根目錄） |
   | `EcmsAlertStatus` | `ON` | `EC_CHECKING_STATUS` **不等於**此值的列整列標紅（`pm-alert`）；留空＝不標 |
   | `EcmsColumns` | （全部） | 只顯示這些欄，如 `EQPID,EC_CHECKING_STATUS,UPDATE_TMST` |
   | `EcmsTitle` | `ECMS Online Status` | 面板標題 |

4. 存檔後重新整理頁面，ECMS 表格會出現在 PM 面板下方，標題旁顯示「資料時間」＝CSV 最後寫入時間，之後隨看板每次 AJAX 更新自動換新。

### 常見問題

- **顯示舊資料 / 沒更新**：看面板「資料時間」是否跟著排程走；沒變代表排程沒寫到 `out_dir`（看 `fetch_ecms_status_log.txt`）。
- **出現 ⚠ 找不到 ECMS 資料檔**：`out_dir` 與 `EcmsCsvPath` 指的不是同一個檔，或 IIS 帳號沒讀取權。
- **中文亂碼**：CSV 是 UTF-8 + BOM、程式用 `Encoding.UTF8` 讀；頁面雖是 `big5`，ASP.NET 輸出時會自動轉碼，一般無須處理。
- **檔案被鎖、程式寫不進去**：確認沒有人用 Excel 開著 `ecms_status.csv`；程式會回 exit code 8。

**Python / exe 版（`fetch_ecms_status.py`）**

與 VBS 版邏輯相同，走 `pywin32` 的同一個 `Notes.NotesSession` COM，可用 PyInstaller 打成單一 `FetchEcmsStatus.exe`，不需在目標機器裝 Python。

- 設定：exe（或 .py）同資料夾放 `fetch_ecms_status.ini`（從 `fetch_ecms_status.ini.example` 複製），只列要覆蓋的項目；沒有 ini 就用內建預設。
- 打包：`build_exe.bat`（會 `pip install -r requirements.txt` 再 `pyinstaller --onefile`），產出 `dist\FetchEcmsStatus.exe`。
- **位元數**：Python 必須跟 Notes client 同位元（Notes 多半 32 位元 → 用 32 位元 Python 打包），否則 `Dispatch("Notes.NotesSession")` 會失敗。查法：`python -c "import struct;print(struct.calcsize('P')*8)"`。
- 排程：`schtasks` 的 `/TR` 直接指到 `FetchEcmsStatus.exe`；log 寫在 exe 同資料夾 `fetch_ecms_status_log.txt`。
- `csv_charset = auto` 會依序試 `utf-8-sig / big5 / cp950 / latin-1`。

**結束代碼**（VBS 與 Python 版相同）

| Code | 意義 |
|------|------|
| 0 | 成功 |
| 3 | 找不到 / 無法開啟信箱 |
| 4 | 無法建立 `Notes.NotesSession`（Notes 未開 / 未登入） |
| 5 | 搜尋失敗，或最近 N 天內沒有符合主旨的信 |
| 6 | 最新那封信沒有 `.csv` 附件 |
| 7 | CSV 無資料列，或標題列找不到 `EQPID_COLUMN` |
| 8 | 寫入輸出檔失敗（檔案被網站 / Excel 鎖住） |

---

其他通用 Notes 工具（列檢視、查文件、下載附件、寄信）在 `main` branch；本 branch 只保留 ECMS 狀態抓取相關檔案。
