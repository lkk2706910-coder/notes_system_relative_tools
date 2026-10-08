// =====================================================================
//  ECMS Online Status 面板 — 貼進 eq_status.aspx.cs（class EQ_Status 內）
//  資料來源：FetchEcmsStatus 排程寫出的 ~/data/ecms_status.csv（UTF-8 + BOM）
//
//  使用方式：
//   1. 把本檔「區塊 A」整段貼到 class EQ_Status 內（例如放在 BuildPmHtml 後面）
//   2. 把「區塊 B」那 3 行貼到 Page_Load 裡 PM 區塊之後、phTable.Controls.Clear() 之前
//   3. web.config <appSettings> 可選設定（都有預設值）：
//        EcmsCsvPath      ~/data/ecms_status.csv
//        EcmsAlertStatus  ON        （EC_CHECKING_STATUS 不等於此值的列整列標紅；留空＝不標）
//        EcmsColumns      留空＝顯示 CSV 全部欄位；或填 "EQPID,EC_CHECKING_STATUS,UPDATE_TMST"
//        EcmsTitle        ECMS Online Status
//  與 PM 面板同一套保護：讀檔失敗顯示上一次資料並掛警示；依檔案修改時間快取，不重複解析。
//  面板沿用 pm-panel / pm-table / pm-alert 等既有 class，深色模式自動套用。
// =====================================================================

// ---------------- 區塊 A：貼進 class EQ_Status ----------------

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

private string BuildEcmsHtml(List<string[]> rows, string warning, DateTime fileTimeLocal)
{
    StringBuilder sb = new StringBuilder();
    sb.Append("<div class='pm-panel ecms-panel'>");
    sb.Append("<div class='pm-title'>" + Server.HtmlEncode(Setting("EcmsTitle", "ECMS Online Status")) + "<span class='pm-sub'>");
    if (fileTimeLocal != DateTime.MinValue)
        sb.Append("資料時間 " + fileTimeLocal.ToString("MM-dd HH:mm", CultureInfo.InvariantCulture));
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

// ---------------- 區塊 B：貼進 Page_Load（PM 區塊之後、phTable.Controls.Clear() 之前） ----------------
/*
        // ECMS Online Status（Notes 信件 CSV）放在 PM 面板下方；同在 .board 內，AJAX 更新會一併換新
        string ecmsWarning;
        DateTime ecmsTime;
        List<string[]> ecmsRows = LoadEcmsCached(out ecmsWarning, out ecmsTime);
        html.Append(BuildEcmsHtml(ecmsRows, ecmsWarning, ecmsTime));
*/
