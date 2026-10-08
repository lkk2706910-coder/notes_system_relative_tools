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

## 機況表（ASPX）整合

程式每次跑完會覆寫兩個檔（先寫 `.tmp` 再 rename，網頁不會讀到半成品）：

| 檔 | 內容 | 適合 |
|---|---|---|
| `ecms_status.html` | 一段 `<div class="ecms-status">…<table>…</table></div>`，含 meta 行（來源主旨 / 收信時間 / 更新時間 / 筆數） | 直接嵌進頁面，零程式碼 |
| `ecms_status.csv` | 篩選後資料，UTF-8 + BOM | 想自己排版 / 加條件格式 |

### 1. 把輸出指到站台資料夾

在 `fetch_ecms_status.ini` 設：
```ini
out_dir = C:\inetpub\wwwroot\status\data
```
（路徑換成你站台實際位置，`data` 子資料夾會自動建立。）

**權限**：排程執行的 Windows 帳號要對該資料夾有「寫入」；IIS 應用程式集區帳號（通常 `IIS AppPool\<集區名>`）要有「讀取」。資料夾內容 → 安全性 → 新增這兩個帳號即可。

### 2A. 最簡單：`<iframe>`（不用動 code-behind）

在機況表頁面底部放：
```html
<iframe id="ecms" src="data/ecms_status.html" style="width:100%;border:0;height:600px"></iframe>
<script>
  // 每 5 分鐘重載一次，並加時間戳避免瀏覽器快取
  setInterval(function () {
    document.getElementById("ecms").src = "data/ecms_status.html?t=" + Date.now();
  }, 5 * 60 * 1000);
</script>
```
高度可視內容調整，或用 `onload` 量 iframe 內容高度自動撐開。

### 2B. 嵌進同一頁：`Literal` 讀檔（版面統一、可套同一份 CSS）

ASPX：
```aspx
<asp:Literal ID="litEcms" runat="server" Mode="PassThrough" />
```

C# code-behind：
```csharp
using System.IO;
using System.Text;

protected void Page_Load(object sender, EventArgs e)
{
    string p = Server.MapPath("~/data/ecms_status.html");
    litEcms.Text = File.Exists(p)
        ? File.ReadAllText(p, Encoding.UTF8)
        : "<p style='color:#999'>ECMS 狀態尚無資料</p>";
}
```

VB.NET code-behind：
```vb
Imports System.IO
Imports System.Text

Protected Sub Page_Load(sender As Object, e As EventArgs) Handles Me.Load
    Dim p As String = Server.MapPath("~/data/ecms_status.html")
    If File.Exists(p) Then
        litEcms.Text = File.ReadAllText(p, Encoding.UTF8)
    Else
        litEcms.Text = "<p style='color:#999'>ECMS 狀態尚無資料</p>"
    End If
End Sub
```

頁面要自動更新畫面，在 `<head>` 加 `<meta http-equiv="refresh" content="300">`（每 5 分鐘），或用 `UpdatePanel` + `Timer`。

### 2C. 自己排版：讀 CSV 綁 `GridView`

ASPX：
```aspx
<asp:GridView ID="gvEcms" runat="server" AutoGenerateColumns="true" CssClass="ecms-table" />
```

C#：
```csharp
using System.Data;
using System.IO;
using System.Text;

void BindEcms()
{
    string p = Server.MapPath("~/data/ecms_status.csv");
    if (!File.Exists(p)) return;
    var lines = File.ReadAllLines(p, Encoding.UTF8);
    if (lines.Length == 0) return;
    var dt = new DataTable();
    foreach (var h in SplitCsv(lines[0])) dt.Columns.Add(h);
    for (int i = 1; i < lines.Length; i++)
    {
        if (string.IsNullOrWhiteSpace(lines[i])) continue;
        var cells = SplitCsv(lines[i]);
        var row = dt.NewRow();
        for (int c = 0; c < dt.Columns.Count && c < cells.Length; c++) row[c] = cells[c];
        dt.Rows.Add(row);
    }
    gvEcms.DataSource = dt;
    gvEcms.DataBind();
}

// 處理「有雙引號、逗號在引號內」的 CSV 欄位
static string[] SplitCsv(string line)
{
    var result = new System.Collections.Generic.List<string>();
    var sb = new StringBuilder(); bool inQ = false;
    for (int i = 0; i < line.Length; i++)
    {
        char ch = line[i];
        if (inQ) { if (ch == '"') { if (i + 1 < line.Length && line[i + 1] == '"') { sb.Append('"'); i++; } else inQ = false; } else sb.Append(ch); }
        else if (ch == '"') inQ = true;
        else if (ch == ',') { result.Add(sb.ToString()); sb.Clear(); }
        else sb.Append(ch);
    }
    result.Add(sb.ToString());
    return result.ToArray();
}
```
在 `Page_Load` 裡 `if (!IsPostBack) BindEcms();`。接著就能用 `RowDataBound` 做條件格式（例如 `EC_CHECKING_STATUS` 不是 `ON` 就標紅）。

### 常見問題

- **網頁顯示舊資料**：瀏覽器快取 → `iframe` 用上面帶 `?t=` 的寫法；或 IIS 對 `data/` 關掉輸出快取。
- **中文亂碼**：兩個輸出檔都是 UTF-8；`Literal` 讀檔請用 `Encoding.UTF8`；頁面 `<meta charset="utf-8">`。
- **IIS 不給下載 `.html`/`.csv`**：靜態檔預設可讀；若站台鎖了 MIME 類型，在 IIS → MIME 類型補 `.csv → text/csv`。
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
