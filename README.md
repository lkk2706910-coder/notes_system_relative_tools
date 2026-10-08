# notes_system_relative_tools — ECMS Online Status 機況表整合

每天 4 次從自己的 Notes 信箱抓「ECMS P58 Current Online Status Report」信件的 CSV 附件，篩選出指定機台後寫成網頁要讀的檔，讓 `eq_status.aspx` 機況看板在 PM 面板下方顯示 ECMS 狀態表。

```
Notes 信箱 ──FetchEcmsStatus.exe（排程 4 次/天）──▶ data\ecms_status.csv ──▶ eq_status.aspx 看板（自動刷新）
```

- 抓取端：`fetch_ecms_status.py` → PyInstaller 打成 `FetchEcmsStatus.exe`，走 `pywin32` 的 `Notes.NotesSession` COM，**唯讀、不修改信件**。
- 網頁端：`aspx_integration/EcmsPanel.snippet.cs` 貼進 `eq_status.aspx.cs`，比照既有 PM 面板做法。
- AI agent 工作指南見 [`AGENTS.md`](AGENTS.md)；其他通用 Notes 工具在 `main` branch。

---

## 1. 抓取端：`FetchEcmsStatus.exe`

### 做什麼

1. 在信箱用 `@Begins(Subject; "ECMS P58 Current Online Status Report")` 找最近 `lookback_days` 天的信，取**最新一封**（依 `DeliveredDate`，沒有就用文件建立時間）。
2. **不落地**讀出第一個 `.csv` 附件：用 `NotesEmbeddedObject.InputStream` 直接把 bytes 讀進記憶體（Notes 8.5.1 以上）；舊版沒有 `InputStream` 才退回 `ExtractFile` 到 `_tmp\` 讀完即刪。
3. 自動判斷分隔（逗號/Tab/分號）與編碼（`csv_charset=auto` 依序試 `utf-8-sig / big5 / cp950 / latin-1`）。
4. 依 `eqpid_like` 篩選 `EQPID` 欄（`%` 萬用字元、`;` 分隔多個、符合任一即保留；預設 `NISACVD-B%;SACVD-B%`）。
5. 輸出到 `out_dir`（先寫 `.tmp` 再 rename，網頁不會讀到半成品）：
   - `ecms_status.csv`：篩選後資料，UTF-8 + BOM（網頁讀這個）
   - `ecms_status.html`：同內容的 HTML 表格片段（備用，想直接嵌別的頁面可用）
   - `raw\yyyymmdd_hhmm_<原附件名>.csv`：原始附件備份（`keep_raw_copy=true` 時）
6. log 寫在 exe 同資料夾 `fetch_ecms_status_log.txt`。

### 設定：`fetch_ecms_status.ini`

放在 exe 同資料夾（從 `fetch_ecms_status.ini.example` 複製），只列要覆蓋的項目；沒有 ini 就用內建預設。

| key | 預設 | 說明 |
|---|---|---|
| `server` / `dbfile` | `UMCM74/UMC` / `mail\00047829.nsf` | 信箱伺服器與信箱檔 |
| `subject_prefix` | `ECMS P58 Current Online Status Report` | 主旨開頭 |
| `lookback_days` | `2` | 只在最近 N 天的信裡找 |
| `eqpid_column` | `EQPID` | CSV 標題列中的欄名（不分大小寫） |
| `eqpid_like` | `NISACVD-B%;SACVD-B%` | 篩選樣式 |
| `out_dir` | （exe 所在資料夾） | **填 ASPX 站台的 `data\` 資料夾**，例 `C:\inetpub\wwwroot\<站台>\data` |
| `out_csv` / `out_html` | `ecms_status.csv` / `ecms_status.html` | 輸出檔名 |
| `keep_raw_copy` | `true` | 是否備份原始附件 |
| `csv_charset` | `auto` | 附件編碼；中文亂碼可固定 `big5` |
| `log_file` | `fetch_ecms_status_log.txt` | log 檔 |

### 打包成 exe

```bat
build_exe.bat
```
會用**同一個** Python 安裝 `pywin32`、`pyinstaller`，先確認 `import win32com.client` 成功，再以 `--hidden-import`（`pythoncom`、`pywintypes`、`win32com.client`…）打包，產出 `dist\FetchEcmsStatus.exe`。機器上有多個 Python 時，改 bat 開頭的 `set PY=`（例如 `py -3-32`）。

手動打包等價指令：
```bat
python -m pip install --upgrade pywin32 pyinstaller
python -m PyInstaller --onefile --console --clean --name FetchEcmsStatus ^
  --hidden-import pythoncom --hidden-import pywintypes ^
  --hidden-import win32com --hidden-import win32com.client ^
  --hidden-import win32com.client.dynamic --hidden-import win32com.client.gencache ^
  --hidden-import win32timezone --collect-submodules win32com ^
  fetch_ecms_status.py
```
exe 若印 `載入 pywin32 失敗`，幾乎都是「安裝 pywin32 的 Python」和「打包用的 Python」不是同一個，或漏了上面的 `--hidden-import`。

> **位元數**：Python 必須跟 Notes client 同位元（Notes 多半是 32 位元 → 用 32 位元 Python 打包），否則 `Dispatch("Notes.NotesSession")` 會失敗。查法：`python -c "import struct;print(struct.calcsize('P')*8)"` 要印 `32`。

### 執行與排程

手動：`FetchEcmsStatus.exe`（或開發時 `python fetch_ecms_status.py`）。前提：**Notes client 已開啟並登入**。

### 怎麼確認真的抓到附件（不落地也看得到）

- **log / 畫面**：`附件：<檔名>（<bytes> bytes，InputStream）` 代表已把附件讀進記憶體；`OK 原始 N 筆 → 篩選後 M 筆` 是整份附件的列數與符合條件的列數。
- **`--preview [N]`**：加這個參數會把標題列與篩選後前 N 列（預設 10）直接印在畫面與 log，不用開任何檔：
  ```bat
  FetchEcmsStatus.exe --preview 20
  ```
- **`keep_raw_copy=true`**（預設）：把記憶體裡的附件原樣寫到 `out_dir\raw\yyyymmdd_hhmm_<附件名>.csv`，可拿去跟信裡的附件比對；確認沒問題後可改 `false` 完全不留原檔。

信在 08:00 / 15:30 / 20:30 / 23:30 到，建議各延後 10 分鐘跑。系統管理員命令提示字元：

```bat
schtasks /Create /TN "ECMS Status 0810" /TR "\"C:\path\to\FetchEcmsStatus.exe\"" /SC DAILY /ST 08:10 /F
schtasks /Create /TN "ECMS Status 1540" /TR "\"C:\path\to\FetchEcmsStatus.exe\"" /SC DAILY /ST 15:40 /F
schtasks /Create /TN "ECMS Status 2040" /TR "\"C:\path\to\FetchEcmsStatus.exe\"" /SC DAILY /ST 20:40 /F
schtasks /Create /TN "ECMS Status 2340" /TR "\"C:\path\to\FetchEcmsStatus.exe\"" /SC DAILY /ST 23:40 /F
```
工作排程器要選「只有在使用者登入時才執行」（COM 需要登入中的 Notes session），「起始於」填 exe 所在資料夾（ini 與 log 才找得到）。

### 結束代碼

| Code | 意義 |
|------|------|
| 0 | 成功 |
| 3 | 找不到 / 無法開啟信箱 |
| 4 | 無法建立 `Notes.NotesSession`（Notes 未開 / 未登入 / pywin32 未裝 / 位元數不符） |
| 5 | 搜尋失敗，或最近 N 天內沒有符合主旨的信 |
| 6 | 最新那封信沒有 `.csv` 附件 |
| 7 | CSV 解碼失敗、無資料列，或標題列找不到 `eqpid_column` |
| 8 | 寫入輸出檔失敗（檔案被網站 / Excel 鎖住） |

---

## 2. 網頁端：`eq_status.aspx` 加 ECMS 面板

### 原理

看板每 `REFRESH_MS` 會自己 `fetch` 一次頁面、把 `.board` 的 `innerHTML` 整個換掉，而 `phTable` 就在 `.board` 裡——所以 ECMS 表格由 **code-behind append 進 `phTable`**，就會跟 PM 面板一起自動刷新，不需要 iframe、不用再寫前端 JS。

程式碼在 [`aspx_integration/EcmsPanel.snippet.cs`](aspx_integration/EcmsPanel.snippet.cs)，做法完全比照既有 PM 面板（`BuildPmHtml`）。

### 面板功能

| 功能 | 說明 |
|---|---|
| 位置 | 看板最下方、PM 面板之後 |
| 外觀 | 沿用 `pm-panel / pm-title / pm-sub / pm-table / pm-wrap / pm-alert / pm-empty / data-warn` class，**深色模式自動套用**，不必新增 CSS |
| 標題列 | 面板標題（`EcmsTitle`）＋「資料時間 MM-dd HH:mm」＝ `ecms_status.csv` 最後寫入時間，一眼看出排程有沒有更新 |
| 欄位 | 預設顯示 CSV 全部欄位（`EQP_TYPE, EQP_MODEL, EQPID, EC_CHECKING_STATUS, UPDATEUSER, UPDATE_TMST`）；`EcmsColumns` 可只挑幾欄、順序照設定 |
| 異常標紅 | `EC_CHECKING_STATUS` **不等於** `EcmsAlertStatus`（預設 `ON`）的列整列紅（`pm-alert`，與 PM 面板同色）；留空＝不標 |
| 快取 | 依 CSV 檔案修改時間快取，檔沒變就不重新解析；多個瀏覽者共用 |
| 保護 | 同時只有一個請求讀檔（`Monitor.TryEnter`，其餘直接用舊資料不排隊）；讀檔失敗顯示上一次資料並掛 ⚠ 警示，不會整頁 500 |
| 無資料 | 檔不存在或沒有資料列 → 顯示「目前沒有 ECMS 資料」（檔不存在另掛 ⚠ 路徑提示） |
| 自動更新 | 隨看板 AJAX 更新一併換新，不需額外設定 |

### 安裝步驟

1. **輸出指到站台資料夾**（`fetch_ecms_status.ini`）：
   ```ini
   out_dir = C:\inetpub\wwwroot\<你的站台>\data
   ```
   權限：排程帳號對 `data\` 要「寫入」；IIS 應用程式集區帳號（`IIS AppPool\<集區名>`）要「讀取」。

2. **貼程式碼**到 `eq_status.aspx.cs`：
   - snippet「區塊 A」整段貼進 `class EQ_Status`（例如放在 `BuildPmHtml` 後面）。
   - 「區塊 B」貼進 `Page_Load`，位置在 `if (PmEnabled()) { ... }` **之後**、`phTable.Controls.Clear();` **之前**：
     ```csharp
     string ecmsWarning;
     DateTime ecmsTime;
     List<string[]> ecmsRows = LoadEcmsCached(out ecmsWarning, out ecmsTime);
     html.Append(BuildEcmsHtml(ecmsRows, ecmsWarning, ecmsTime));
     ```
   - 既有 `using` 已足夠（`System.IO`、`System.Text`、`System.Threading`、`System.Globalization`、`System.Collections.Generic`）。

3. **（選用）`web.config` `<appSettings>`**：

   | key | 預設 | 說明 |
   |---|---|---|
   | `EcmsCsvPath` | `~/data/ecms_status.csv` | CSV 位置（相對站台根目錄；要與 `out_dir` 指同一個檔） |
   | `EcmsAlertStatus` | `ON` | 不等於此值的列標紅；留空＝不標 |
   | `EcmsColumns` | （全部） | 例 `EQPID,EC_CHECKING_STATUS,UPDATE_TMST` |
   | `EcmsTitle` | `ECMS Online Status` | 面板標題 |

   ```xml
   <appSettings>
     <add key="EcmsColumns" value="EQPID,EC_CHECKING_STATUS,UPDATEUSER,UPDATE_TMST" />
   </appSettings>
   ```

4. 存檔後重新整理頁面即可看到面板。

### 常見問題

- **面板「資料時間」沒跟著排程走**：排程沒寫到 `out_dir`，看 `fetch_ecms_status_log.txt`。
- **⚠ 找不到 ECMS 資料檔**：`out_dir` 與 `EcmsCsvPath` 指的不是同一個檔，或 IIS 帳號沒讀取權。
- **中文亂碼**：CSV 是 UTF-8 + BOM、程式用 `Encoding.UTF8` 讀；頁面雖是 `big5`，ASP.NET 輸出會自動轉碼。
- **一開就 `UnicodeDecodeError ... load_config`**：`fetch_ecms_status.ini` 被記事本存成 ANSI(Big5)。現在程式會自動辨識 UTF-8 / Big5，若仍出錯請把 ini 另存為 UTF-8。
- **exe 寫檔失敗（exit 8）**：有人用 Excel 開著 `ecms_status.csv`。
- **`com_error (-2147352573, '找不到成員')`**：pywin32 拿不到 Notes 型別資訊，把方法當屬性呼叫。程式已改成先 `_FlagAsMethod` 再呼叫；若仍出現，幾乎可確定是 Python 與 Notes client **位元數不同**（64 位元 Python 對 32 位元 Notes），請改用 32 位元 Python（`py -3-32`）打包。
- **exe 連不上 Notes（exit 4）**：Notes 沒開/沒登入、排程不是跑在登入 session、或 Python 位元數與 Notes 不符。
