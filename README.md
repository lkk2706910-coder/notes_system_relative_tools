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

**機況表（ASPX）嵌入方式**

把 `OUT_DIR` 指到站台資料夾（例 `C:\inetpub\wwwroot\status\data`），然後擇一：

1. 最簡單——`<iframe>`：
   ```html
   <iframe src="data/ecms_status.html" style="width:100%;border:0;height:600px"></iframe>
   ```
2. 直接嵌進頁面（code-behind 讀檔塞進 `Literal`）：
   ```aspx
   <asp:Literal ID="litEcms" runat="server" />
   ```
   ```csharp
   // Page_Load
   var p = Server.MapPath("~/data/ecms_status.html");
   litEcms.Text = File.Exists(p) ? File.ReadAllText(p, Encoding.UTF8) : "<p>尚無資料</p>";
   ```
3. 要自己排版：讀 `ecms_status.csv` 綁到 `GridView`。

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

## 其他工具

`ListNotesViews.vbs`（列檢視）、`TestFindDoc.vbs`（測試查文件）、`DownloadNotesAttachments.vbs` + `download_attachments.bat`（下載附件）、`SendNotesMail.vbs`（MIME 寄 HTML 信）。各檔頂端註解有完整說明與設定。
