# 任務：在 `eq_status.aspx` 機況看板加入「ECMS Online Status」面板

請修改 `eq_status.aspx.cs`（`public partial class EQ_Status`），在看板最下方、現有 **PM 狀態面板之後**，加入一個 ECMS Online Status 表格面板。**只做加法，不要改動既有行為**。

## 背景（已知事實，不用再查）

- 看板前端每 `REFRESH_MS` 會 `fetch` 自己一次、把 `.board` 的 `innerHTML` 整個換掉；`phTable` 在 `.board` 內。因此面板只要由 code-behind append 進 `phTable`，就會跟 PM 面板一起自動刷新，**不要加 iframe、不要新增前端 JS、不要新增 CSS**。
- 資料檔由另一支排程程式每天 4 次（約 08:10 / 15:40 / 20:40 / 23:40）覆寫到站台 `data\` 資料夾：
  - `~/data/ecms_status.csv`：UTF-8 + BOM、逗號分隔、第一列標題。欄位：`EQP_TYPE, EQP_MODEL, EQPID, EC_CHECKING_STATUS, UPDATEUSER, UPDATE_TMST`（約 27 列）
  - `~/data/ecms_status.meta.json`：`{"subject","mail_time","fetch_time","attachment","rows_total","rows_kept","eqpid_like"}`，時間格式 `yyyy-MM-dd HH:mm:ss`
- 既有可直接沿用的東西：`Setting(key, fallback)`、`SplitSetting(key, fallback)`、CSS class `pm-panel / pm-title / pm-sub / pm-wrap / pm-table / pm-alert / pm-empty / data-warn`（含深色模式）。`using` 已有 `System.IO`、`System.Text`、`System.Threading`、`System.Globalization`、`System.Collections.Generic`、`System.Web.Script.Serialization`，不需新增。
- 專案是舊版 ASP.NET Web Site（`CodeFile`），C# 請用與檔案一致的保守語法（`out` 變數先宣告、不用 `?.`、不用字串內插）。頁面 `<meta charset="big5">`，輸出一律 `Server.HtmlEncode`。

## 要做的事

1. 把下方「區塊 A」整段貼進 `class EQ_Status`（建議放在 `BuildPmHtml` 方法之後）。
2. 在 `Page_Load` 裡、`if (PmEnabled()) { ... }` 區塊**之後**、`phTable.Controls.Clear();` **之前**插入：
   ```csharp
   // ECMS Online Status（Notes 信件 CSV）放在 PM 面板下方；同在 .board 內，AJAX 更新會一併換新
   string ecmsWarning;
   DateTime ecmsTime;
   List<string[]> ecmsRows = LoadEcmsCached(out ecmsWarning, out ecmsTime);
   html.Append(BuildEcmsHtml(ecmsRows, ecmsWarning, ecmsTime));
   ```
3. （選用）`web.config` `<appSettings>` 支援以下 key，全部有預設值、不設也能跑：

   | key | 預設 | 意義 |
   |---|---|---|
   | `EcmsCsvPath` | `~/data/ecms_status.csv` | CSV 位置 |
   | `EcmsAlertStatus` | `ON` | `EC_CHECKING_STATUS` 不等於此值的列加 `pm-alert`（整列紅）；留空＝不標 |
   | `EcmsColumns` | （全部） | 只顯示指定欄，逗號分隔，例 `EQPID,EC_CHECKING_STATUS,UPDATEUSER,UPDATE_TMST` |
   | `EcmsTitle` | `ECMS Online Status` | 面板標題 |

## 面板行為規格

- 標題列：`<標題>` + `pm-sub`：`收信 MM-dd HH:mm　抓取 MM-dd HH:mm　N 筆` + 信件主旨（來自 meta.json；meta 不存在時退回顯示 CSV 最後寫入時間）。
- 依 CSV 的 `LastWriteTimeUtc` 快取解析結果；檔沒變不重讀。同時只允許一個請求讀檔（`Monitor.TryEnter`，其餘直接用舊資料不等待）。讀檔失敗 → 顯示上一次資料並在面板內掛 `data-warn` 警示，**不可整頁 500**。
- CSV 不存在或沒有資料列 → 顯示 `pm-empty`「目前沒有 ECMS 資料」（不存在另掛 ⚠ 路徑提示）。
- CSV 切欄需引號感知（欄位內可含逗號、`""` 轉義）。

## 區塊 A（直接貼）

```csharp
private const string EcmsDefaultPath = "~/data/ecms_status.csv";
private static readonly object EcmsLock = new object();
private static List<string[]> EcmsRows;                 // [0] 是標題列
private static DateTime EcmsFileTimeUtc = DateTime.MinValue;
private static string EcmsLastError;

// 依檔案修改時間快取；別的請求正在讀就直接用舊資料，不排隊
private List<string[]> LoadEcmsCached(out string warning, out DateTime fileTimeLocal)
{
    warning = null;
    fileTimeLocal = DateTime.MinValue;

    string path = Server.MapPath(Setting("EcmsCsvPath", EcmsDefaultPath));
    if (!File.Exists(path))
    {
        warning = "找不到 ECMS 資料檔：" + path;
        return EcmsRows;
    }

    DateTime mt = File.GetLastWriteTimeUtc(path);
    fileTimeLocal = mt.ToLocalTime();
    if (EcmsRows != null && mt == EcmsFileTimeUtc)
    {
        if (EcmsLastError != null) warning = "ECMS 資料讀取失敗，顯示舊資料：" + EcmsLastError;
        return EcmsRows;
    }

    if (!Monitor.TryEnter(EcmsLock))
        return EcmsRows ?? new List<string[]>();

    try
    {
        if (EcmsRows != null && mt == EcmsFileTimeUtc) return EcmsRows;

        var rows = new List<string[]>();
        foreach (string line in File.ReadAllLines(path, Encoding.UTF8))
        {
            if (line.Trim().Length == 0) continue;
            rows.Add(SplitCsvLine(line));
        }
        EcmsRows = rows;
        EcmsFileTimeUtc = mt;
        EcmsLastError = null;
        return EcmsRows;
    }
    catch (Exception ex)
    {
        EcmsLastError = ex.Message;
        warning = "ECMS 資料讀取失敗，顯示舊資料：" + ex.Message;
        return EcmsRows ?? new List<string[]>();
    }
    finally
    {
        Monitor.Exit(EcmsLock);
    }
}

// 引號感知的 CSV 切欄（欄位內可含逗號、雙引號以 "" 轉義）
private static string[] SplitCsvLine(string line)
{
    var result = new List<string>();
    var sb = new StringBuilder();
    bool inQ = false;
    for (int i = 0; i < line.Length; i++)
    {
        char ch = line[i];
        if (inQ)
        {
            if (ch == '"')
            {
                if (i + 1 < line.Length && line[i + 1] == '"') { sb.Append('"'); i++; }
                else inQ = false;
            }
            else sb.Append(ch);
        }
        else if (ch == '"') inQ = true;
        else if (ch == ',') { result.Add(sb.ToString().Trim()); sb.Length = 0; }
        else sb.Append(ch);
    }
    result.Add(sb.ToString().Trim());
    return result.ToArray();
}

// 讀 exe 一起寫出的 ecms_status.meta.json（主旨 / 收信時間 / 抓取時間 / 筆數）；沒有就回 null
private Dictionary<string, object> LoadEcmsMeta()
{
    try
    {
        string csvPath = Server.MapPath(Setting("EcmsCsvPath", EcmsDefaultPath));
        string metaPath = Path.ChangeExtension(csvPath, null) + ".meta.json";
        if (!File.Exists(metaPath)) return null;
        JavaScriptSerializer ser = new JavaScriptSerializer();
        return ser.Deserialize<Dictionary<string, object>>(File.ReadAllText(metaPath, Encoding.UTF8));
    }
    catch (Exception)
    {
        return null;
    }
}

private static string MetaStr(Dictionary<string, object> meta, string key)
{
    object v;
    if (meta == null || !meta.TryGetValue(key, out v) || v == null) return "";
    return Convert.ToString(v, CultureInfo.InvariantCulture);
}

// "yyyy-MM-dd HH:mm:ss" → "MM-dd HH:mm"；解析失敗就原樣回傳
private static string ShortTime(string s)
{
    DateTime d;
    if (DateTime.TryParseExact(s, "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture, DateTimeStyles.None, out d))
        return d.ToString("MM-dd HH:mm", CultureInfo.InvariantCulture);
    return s;
}

private string BuildEcmsHtml(List<string[]> rows, string warning, DateTime fileTimeLocal)
{
    Dictionary<string, object> meta = LoadEcmsMeta();
    StringBuilder sb = new StringBuilder();
    sb.Append("<div class='pm-panel ecms-panel'>");
    sb.Append("<div class='pm-title'>" + Server.HtmlEncode(Setting("EcmsTitle", "ECMS Online Status")) + "<span class='pm-sub'>");
    string mailTime = MetaStr(meta, "mail_time");
    string fetchTime = MetaStr(meta, "fetch_time");
    if (mailTime != "") sb.Append("收信 " + Server.HtmlEncode(ShortTime(mailTime)));
    if (fetchTime != "") sb.Append("　抓取 " + Server.HtmlEncode(ShortTime(fetchTime)));
    if (mailTime == "" && fetchTime == "" && fileTimeLocal != DateTime.MinValue)
        sb.Append("資料時間 " + fileTimeLocal.ToString("MM-dd HH:mm", CultureInfo.InvariantCulture));
    if (rows != null && rows.Count > 1)
        sb.Append("　" + (rows.Count - 1).ToString(CultureInfo.InvariantCulture) + " 筆");
    string subj = MetaStr(meta, "subject");
    if (subj != "") sb.Append("</span><span class='pm-sub' title='" + Server.HtmlEncode(subj) + "'>" + Server.HtmlEncode(subj) + "");
    sb.Append("</span></div>");

    if (!string.IsNullOrEmpty(warning))
        sb.Append("<div class='data-warn'>⚠ " + Server.HtmlEncode(warning) + "</div>");

    if (rows == null || rows.Count < 2)
    {
        sb.Append("<div class='pm-empty'>目前沒有 ECMS 資料</div></div>");
        return sb.ToString();
    }

    string[] header = rows[0];

    // 顯示欄位：appSettings 可指定；只留標題裡真的有的欄位，一個都對不上就全部顯示
    string[] wanted = SplitSetting("EcmsColumns", "");
    var colIdx = new List<int>();
    foreach (string w in wanted)
    {
        for (int i = 0; i < header.Length; i++)
            if (string.Equals(header[i], w, StringComparison.OrdinalIgnoreCase)) { colIdx.Add(i); break; }
    }
    if (colIdx.Count == 0) for (int i = 0; i < header.Length; i++) colIdx.Add(i);

    // 狀態欄：EC_CHECKING_STATUS 不等於 EcmsAlertStatus（預設 ON）整列標紅
    int statusIdx = -1;
    for (int i = 0; i < header.Length; i++)
        if (string.Equals(header[i], "EC_CHECKING_STATUS", StringComparison.OrdinalIgnoreCase)) { statusIdx = i; break; }
    string alertStatus = Setting("EcmsAlertStatus", "ON");

    sb.Append("<div class='pm-wrap'><table class='pm-table'><thead><tr>");
    foreach (int c in colIdx) sb.Append("<th>" + Server.HtmlEncode(header[c]) + "</th>");
    sb.Append("</tr></thead><tbody>");

    for (int r = 1; r < rows.Count; r++)
    {
        string[] row = rows[r];
        bool alert = statusIdx >= 0 && alertStatus != "" && statusIdx < row.Length
                  && !string.Equals(row[statusIdx], alertStatus, StringComparison.OrdinalIgnoreCase);
        sb.Append(alert ? "<tr class='pm-alert'>" : "<tr>");
        foreach (int c in colIdx)
        {
            string v = c < row.Length ? row[c] : "";
            sb.Append("<td>" + Server.HtmlEncode(v) + "</td>");
        }
        sb.Append("</tr>");
    }
    sb.Append("</tbody></table></div></div>");
    return sb.ToString();
}
```

## 驗收

1. 編譯通過，頁面載入無錯。
2. PM 面板下方出現 ECMS 面板，標題旁有收信/抓取時間與筆數，表格 6 欄、約 27 列；`EC_CHECKING_STATUS` 為 `NA` 的列整列紅。
3. 深色模式下外觀與 PM 面板一致。
4. 把 `data\ecms_status.csv` 暫時改名 → 面板顯示 ⚠ 與「目前沒有 ECMS 資料」，頁面其他部分正常；改回後下一次刷新恢復。
5. 既有功能（機況磁磚、異常清單、PM 面板、寵物、深色切換、AJAX 更新）行為不變。
