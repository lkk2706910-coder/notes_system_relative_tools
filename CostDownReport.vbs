Option Explicit

' CostDownReport.vbs
' 用途：從 CDPAS 資料庫的檢視「6. 依分類 (Last 2 Years)」讀取 cost-down 項目，
'       篩選「Forecast_Rls_Date 有值 且 Actual_Date_Final 無值」的資料，
'       依 Forecast_Rls_Date 月份分組整理成 HTML 表格，透過 Notes 寄信到指定收件人，
'       並另存一份「互動版」HTML（瀏覽器開啟：點欄位標題排序、第二列逐欄篩選）當附件 / 存檔。
'       Project ID 做成 Notes 文件連結（＝文件內容 → Meta → Identifier 的 notes:// 網址），可點擊回原文件。
'       Notes 信件內文不能跑 JavaScript，所以信件內文是靜態表格，排序 / 篩選請開附件。
'
' 前置需求：先開啟並登入 Notes client；需對目標資料庫有 Reader 權限。
' 執行：  cscript //nologo CostDownReport.vbs
'
' ============================ 設定區（請依你的環境修改） ============================

' --- Notes 端 ---
Const SERVER   = "UMCD46/UMC"                 ' Domino 伺服器
Const DBFILE   = "fab\cdpas.nsf"              ' 資料庫路徑
Const VIEWNAME = "6. 依分類 (Last 2 Years)"   ' 要掃的檢視

' 欄位 item name（如 Notes Fields 分頁所示）
Const F_PROJECT_ID   = "Project_ID"
Const F_SECTION      = "Section"
Const F_COST_POOL    = "CostPool"
Const F_EQP_TYPE     = "eqptype"
Const F_PROJECT_DESC = "Project_Description"
Const F_METHOD       = "Method"
Const F_FORECAST_DT  = "Forecast_Rls_Date"    ' 篩選＋分組用
Const F_MONTHLY      = "Benefit_Per_Month"
Const F_TOTAL        = "Month_Total"
Const F_ACTUAL_DT    = "Actual_Date_Final"    ' 需為空才保留
Const F_STATUS       = "Status"               ' 額外篩選欄
Const STATUS_LIKE    = "New Project%"         ' 對 Status 的比對樣式；% 為萬用字元(任意字串)；留空＝不套用
Const F_STATION_T11  = "StationTime11"        ' 值格式如 "11^New Project - Approved^07/24/2026 16:25^^^"，取最早的 MM/DD/YYYY

' --- 寄信 ---
' 收件人：多人用分號 ; 分隔（可用 & _ 續行方便閱讀）。
' 注意 VBScript 的 Const 不允許多字串串接，這裡改用 Dim。
Dim MAIL_TO : MAIL_TO = "bo_hsiang_kao@umc.com;" & _
                        "g_t_tsai@umc.com;" & _
                        "jackson_lee@umc.com;" & _
                        "chun_shun_lin@umc.com;" & _
                        "yung_feng_tsou@umc.com;" & _
                        "allen_hung@umc.com;" & _
                        "yuan_ying_chen@umc.com;" & _
                        "chien_hsien_li@umc.com;" & _
                        "yi_ting_lin@umc.com"
Const MAIL_SUBJECT = "Cost Down Project 預估報表（未實際結案）"

' --- 其他 ---
' 三個開關可同時開：預設寄信＋附上互動版＋存檔。只想測試不寄信就把 SEND_MAIL 改 False。
Const SEND_MAIL       = True     ' True＝透過 Notes 寄信給 MAIL_TO（內文＝靜態分月表格）
Const ATTACH_HTML     = True     ' True＝信件附上互動版 HTML（可排序 / 篩選；用瀏覽器開）
Const SAVE_HTML       = True     ' True＝把互動版 HTML 存到 OUTFILE（先存檔再寄信，寄信失敗也有檔案）
Const OUTFILE         = "cost_down_report.html"   ' 相對路徑＝以本 .vbs 所在資料夾為基準
Const PROGRESS_EVERY  = 200

' ==================================================================================

' MIME encoding constants (見 HCL Notes 文件)
Const ENC_IDENTITY_BINARY = 1729
Const ENC_BASE64          = 1727

Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
Dim scriptDir : scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)

' ---------- 連 Notes ----------
Dim ns
On Error Resume Next
Set ns = CreateObject("Notes.NotesSession")
If Err.Number <> 0 Then
  WScript.Echo "ERROR: 無法建立 Notes.NotesSession (Notes 未開/未登入?) " & Err.Description
  WScript.Quit 4
End If
On Error Goto 0

Dim db : Set db = ns.GetDatabase(SERVER, DBFILE)
If db Is Nothing Then
  WScript.Echo "ERROR: 找不到資料庫: [" & SERVER & "] " & DBFILE
  WScript.Quit 3
End If
On Error Resume Next
If Not db.IsOpen Then Call db.Open(SERVER, DBFILE)
If Err.Number <> 0 Or Not db.IsOpen Then
  WScript.Echo "ERROR: 無法開啟資料庫: [" & SERVER & "] " & DBFILE & ". " & Err.Description
  WScript.Quit 3
End If
On Error Goto 0

Dim view : Set view = db.GetView(VIEWNAME)
If view Is Nothing Then
  WScript.Echo "ERROR: 找不到檢視: " & VIEWNAME
  WScript.Quit 5
End If

' ---------- 掃描並收集 ----------
' groups: Dictionary key=YYYY/MM -> ArrayList of Dictionary(欄位->值)
Dim groups : Set groups = CreateObject("Scripting.Dictionary")
Dim seenNoteIDs : Set seenNoteIDs = CreateObject("Scripting.Dictionary")

WScript.Echo "開始掃描檢視 " & VIEWNAME & " …"

Dim scanned : scanned = 0
Dim kept : kept = 0

Dim doc : Set doc = view.GetFirstDocument()
Do While Not (doc Is Nothing)
  scanned = scanned + 1
  If PROGRESS_EVERY > 0 Then
    If (scanned Mod PROGRESS_EVERY) = 0 Then _
      WScript.StdOut.Write vbCr & "  掃描中… 已處理 " & scanned & " 筆，保留 " & kept & " 筆"
  End If

  Dim nid : nid = doc.NoteID
  If Not seenNoteIDs.Exists(nid) Then
    seenNoteIDs.Add nid, True

    Dim pid : pid = Trim(ItemText(doc, F_PROJECT_ID))
    Dim fd : fd = ItemDate(doc, F_FORECAST_DT)
    Dim ad : ad = ItemDate(doc, F_ACTUAL_DT)
    Dim stVal : stVal = ItemText(doc, F_STATUS)

    ' 保留條件：Project_ID 有值 AND Forecast_Rls_Date 有值 AND Actual_Date_Final 無值 AND Status 符合 STATUS_LIKE
    If pid <> "" And (Not IsEmpty(fd)) And IsEmpty(ad) And LikeMatchPct(stVal, STATUS_LIKE) Then
      Dim mkey : mkey = FmtMonth(fd)
      If Not groups.Exists(mkey) Then _
        groups.Add mkey, CreateObject("System.Collections.ArrayList")

      Dim rec : Set rec = CreateObject("Scripting.Dictionary")
      rec.Add "project_id",   pid
      rec.Add "section",      ItemText(doc, F_SECTION)
      rec.Add "cost_pool",    ItemText(doc, F_COST_POOL)
      rec.Add "eqp_type",     ItemText(doc, F_EQP_TYPE)
      rec.Add "project_desc", ItemText(doc, F_PROJECT_DESC)
      rec.Add "method",       ItemText(doc, F_METHOD)
      rec.Add "forecast_dt",  FmtDate(fd)
      rec.Add "monthly",      ItemNumber(doc, F_MONTHLY)
      rec.Add "total",        ItemNumber(doc, F_TOTAL)
      rec.Add "approve_date", ExtractApproveDate(ItemText(doc, F_STATION_T11))
      rec.Add "link",         DocLink(doc)
      rec.Add "month",        mkey
      rec.Add "sortkey",      CDbl(Year(fd)) * 10000 + Month(fd) * 100 + Day(fd)
      groups(mkey).Add rec
      kept = kept + 1
    End If
  End If

  Set doc = view.GetNextDocument(doc)
Loop

If PROGRESS_EVERY > 0 Then _
  WScript.StdOut.Write vbCr & "  掃描完成：共 " & scanned & " 筆，保留 " & kept & " 筆" & vbCrLf

If kept = 0 Then
  WScript.Echo "沒有符合條件的項目（Forecast_Rls_Date 有值且 Actual_Date_Final 無值），不寄信。"
  WScript.Quit 0
End If

' ---------- 組 HTML（信件內文：靜態分月表格） ----------
Dim monthKeys : monthKeys = SortedKeys(groups)
Dim html : html = HtmlHeader(scanned, kept, UBound(monthKeys) + 1)
If SEND_MAIL And ATTACH_HTML Then _
  html = html & "<p style='font-size:12px;margin:0 0 10px;color:#1f3b57;'>" & _
                "需要依欄位<b>排序 / 篩選</b>請開啟附件 <b>" & Esc(fso.GetFileName(OUTFILE)) & "</b>（用瀏覽器開啟）。</p>"
Dim mk
Dim mi
For mi = 0 To UBound(monthKeys)
  mk = monthKeys(mi)
  html = html & "<h3 style='margin:18px 0 6px;color:#1f3b57;'>" & Esc(mk) & _
                " <span style='font-weight:normal;color:#667;font-size:13px;'>（" & groups(mk).Count & " 筆）</span></h3>"
  html = html & "<table cellspacing='0' cellpadding='5' style='border-collapse:collapse;font-size:13px;border:1px solid #aac;'>"
  html = html & "<thead><tr style='background:#f2f5f9;color:#1f3b57;'>" & _
                Th("Project ID") & Th("SEC") & Th("Cost Pool") & Th("EQP Type") & Th("Project Desc.") & _
                Th("Method") & Th("Benefit Occur (Forecast)") & Th("Monthly Benefit") & Th("Total Benefit (Forecast)") & Th("Approve Date") & _
                "</tr></thead><tbody>"

  Dim rows : rows = SortRowsBySortKey(groups(mk))
  Dim ri
  For ri = 0 To UBound(rows)
    Dim r : Set r = rows(ri)
    html = html & "<tr>" & _
      TdLink(r("project_id"), r("link")) & Td(r("section")) & Td(r("cost_pool")) & Td(r("eqp_type")) & _
      Td(r("project_desc")) & Td(r("method")) & Td(r("forecast_dt")) & _
      TdN(r("monthly")) & TdN(r("total")) & Td(r("approve_date")) & _
      "</tr>"
  Next
  html = html & "</tbody></table>"
Next
html = html & HtmlFooter()

' ---------- 互動版 HTML（單一平面表格 + 排序 / 篩選），存檔與附件共用 ----------
Dim ihtml : ihtml = BuildInteractiveHtml(monthKeys, scanned, kept)

' ---------- 存 HTML / 寄信（可同時） ----------
Dim summary : summary = "（月份區塊 " & (UBound(monthKeys) + 1) & " 個，共 " & kept & " 筆）"
Dim attachPath : attachPath = ""

If SAVE_HTML Then
  Dim outPath : outPath = AbsPath(OUTFILE)
  EnsureFolder fso.GetParentFolderName(outPath)
  WriteTextUtf8 outPath, ihtml
  attachPath = outPath
  WScript.Echo "OK 已把互動版 HTML 存到：" & outPath & " " & summary
End If

If SEND_MAIL Then
  Dim tempAttach : tempAttach = False
  If ATTACH_HTML And attachPath = "" Then
    ' 沒存檔但要附件：先寫到暫存資料夾，寄完刪掉
    attachPath = fso.BuildPath(fso.GetSpecialFolder(2).Path, fso.GetFileName(OUTFILE))
    WriteTextUtf8 attachPath, ihtml
    tempAttach = True
  End If
  SendHtmlMail MAIL_TO, MAIL_SUBJECT, html, IIf(ATTACH_HTML, attachPath, "")
  If tempAttach Then
    On Error Resume Next
    fso.DeleteFile attachPath
    Err.Clear : On Error Goto 0
  End If
  WScript.Echo "OK 已寄出報表到：" & MAIL_TO & IIf(ATTACH_HTML, "（附互動版 HTML）", "") & " " & summary
End If

If Not SAVE_HTML And Not SEND_MAIL Then
  WScript.Echo "注意：SEND_MAIL 與 SAVE_HTML 都是 False，沒有任何輸出 " & summary
End If

WScript.Quit 0

' ================================ 函式 ================================

' 讀文字欄位；欄位不存在或 Null 回 ""
Function ItemText(d, fld)
  ItemText = ""
  If Not d.HasItem(fld) Then Exit Function
  Dim it : Set it = d.GetFirstItem(fld)
  If it Is Nothing Then Exit Function
  ItemText = it.Text
End Function

' 讀數值欄位；非數字回 0（走 NotesItem.Text，去掉千分位逗號後 CDbl）
Function ItemNumber(d, fld)
  ItemNumber = 0
  Dim s : s = ItemText(d, fld)
  If Trim(s) = "" Then Exit Function
  s = Replace(s, ",", "")
  On Error Resume Next
  Dim x : x = CDbl(s)
  If Err.Number = 0 Then ItemNumber = x
  Err.Clear
  On Error Goto 0
End Function

' 讀日期欄位；欄位不存在或值不是日期回 Empty（走 NotesItem.Text，避開物件轉型）
Function ItemDate(d, fld)
  ItemDate = Empty
  Dim s : s = ItemText(d, fld)
  If Trim(s) = "" Then Exit Function
  On Error Resume Next
  Dim x : x = CDate(s)
  If Err.Number = 0 Then ItemDate = x
  Err.Clear
  On Error Goto 0
End Function

Function FmtDate(d)
  If Not IsDate(d) Then FmtDate = "" : Exit Function
  FmtDate = Right("0000" & Year(d), 4) & "/" & Right("0" & Month(d), 2) & "/" & Right("0" & Day(d), 2)
End Function

Function FmtMonth(d)
  If Not IsDate(d) Then FmtMonth = "" : Exit Function
  FmtMonth = Right("0000" & Year(d), 4) & "/" & Right("0" & Month(d), 2)
End Function

Function FmtNum(n)
  If Not IsNumeric(n) Then FmtNum = CStr(n & "") : Exit Function
  FmtNum = FormatNumber(n, 0, -1, 0, -1)   ' 0 小數位、含千分位
End Function

' 從 Station Time11 抽出核准日期。以 ^ 拆，掃描每段開頭是否為 MM/DD/YYYY，
' 收集所有合法日期，回傳「最早」那一個 (格式 MM/DD/YYYY，補零)。
' 例1："11^New Project - Approved^07/24/2026 16:25^^^"                              -> "07/24/2026"
' 例2："11^New Project - Approved^07/24/2026 14:00^07/27/2026 11:18^2,157.3^"       -> "07/24/2026"
Function ExtractApproveDate(raw)
  ExtractApproveDate = ""
  Dim s : s = Trim(CStr(raw & ""))
  If s = "" Then Exit Function
  Dim parts : parts = Split(s, "^")
  Dim earliest, haveDate : haveDate = False
  Dim i, seg, dp, sp, d
  For i = 0 To UBound(parts)
    seg = Trim(parts(i))
    If seg <> "" Then
      sp = InStr(seg, " ")
      If sp > 0 Then dp = Left(seg, sp - 1) Else dp = seg
      d = ParseMdy(dp)
      If Not IsEmpty(d) Then
        If (Not haveDate) Or (d < earliest) Then
          earliest = d
          haveDate = True
        End If
      End If
    End If
  Next
  If haveDate Then
    ExtractApproveDate = Right("0" & Month(earliest), 2) & "/" & _
                        Right("0" & Day(earliest), 2) & "/" & _
                        Right("0000" & Year(earliest), 4)
  End If
End Function

' 解析 "MM/DD/YYYY" 字串為日期；不合格式或超出範圍回 Empty
Function ParseMdy(s)
  ParseMdy = Empty
  Dim parts : parts = Split(s, "/")
  If UBound(parts) <> 2 Then Exit Function
  If Not IsNumeric(parts(0)) Then Exit Function
  If Not IsNumeric(parts(1)) Then Exit Function
  If Not IsNumeric(parts(2)) Then Exit Function
  Dim m : m = CInt(parts(0))
  Dim dd : dd = CInt(parts(1))
  Dim y : y = CInt(parts(2))
  If m < 1 Or m > 12 Then Exit Function
  If dd < 1 Or dd > 31 Then Exit Function
  If y < 1900 Or y > 2999 Then Exit Function
  On Error Resume Next
  ParseMdy = DateSerial(y, m, dd)
  If Err.Number <> 0 Then ParseMdy = Empty : Err.Clear
  On Error Goto 0
End Function

' 三元運算輔助（VBScript 沒有內建 IIf）
Function IIf(cond, a, b)
  If cond Then IIf = a Else IIf = b
End Function

' 比對「% 為萬用字元(任意字串)」的樣式；不分大小寫。pattern 空＝視為符合
Function LikeMatchPct(s, pattern)
  If pattern = "" Then LikeMatchPct = True : Exit Function
  Dim rx : Set rx = CreateObject("VBScript.RegExp")
  rx.IgnoreCase = True
  rx.Global = False
  Dim out, i, c
  out = "^"
  For i = 1 To Len(pattern)
    c = Mid(pattern, i, 1)
    Select Case c
      Case "%" : out = out & ".*"
      Case ".", "\", "+", "*", "?", "(", ")", "[", "]", "{", "}", "^", "$", "|"
        out = out & "\" & c
      Case Else : out = out & c
    End Select
  Next
  rx.Pattern = out & "$"
  LikeMatchPct = rx.Test(CStr(s & ""))
End Function

Function Esc(s)
  Esc = Replace(CStr(s & ""), "&", "&amp;")
  Esc = Replace(Esc, "<", "&lt;")
  Esc = Replace(Esc, ">", "&gt;")
  Esc = Replace(Esc, """", "&quot;")
End Function

Function Th(txt)
  Th = "<th style='border:1px solid #aac;padding:5px 8px;text-align:left;'>" & Esc(txt) & "</th>"
End Function

Function Td(v)
  Td = "<td style='border:1px solid #aac;padding:4px 8px;'>" & Esc(v) & "</td>"
End Function

Function TdN(v)
  TdN = "<td style='border:1px solid #aac;padding:4px 8px;text-align:right;font-variant-numeric:tabular-nums;'>" & Esc(FmtNum(v)) & "</td>"
End Function

' Project ID 儲存格：有連結就做成 <a href="notes://…">，點擊用 Notes 開回原文件（信件內文用，inline style）
Function TdLink(txt, url)
  If Trim(url & "") = "" Then
    TdLink = Td(txt)
  Else
    TdLink = "<td style='border:1px solid #aac;padding:4px 8px;white-space:nowrap;'>" & _
             "<a href='" & Esc(url) & "' style='color:#1a5fb4;'>" & Esc(txt) & "</a></td>"
  End If
End Function

' ---- 互動版 HTML 用的儲存格（靠 CSS class，排序 / 篩選用 data-* 屬性） ----
' 標題：typ = text / num / date（決定排序與數字篩選）；filt = "select"＝下拉選單篩選，否則文字框
Function ThI(txt, typ, filt)
  ThI = "<th data-t='" & typ & "'" & IIf(filt = "select", " data-f='select'", "") & _
        IIf(typ = "num", " class='num'", "") & ">" & Esc(txt) & "</th>"
End Function

' 一般 / 日期格：disp 顯示文字，sortv 排序用值（空＝用顯示文字）
Function TdV(disp, sortv)
  TdV = "<td" & IIf(sortv <> "", " data-v='" & Esc(sortv) & "'", "") & ">" & Esc(disp) & "</td>"
End Function

' 數字格：顯示千分位，data-v 放原始數值
Function TdNV(v)
  TdNV = "<td class='num' data-v='" & Esc(CStr(v & "")) & "'>" & Esc(FmtNum(v)) & "</td>"
End Function

Function TdLinkV(txt, url)
  If Trim(url & "") = "" Then
    TdLinkV = TdV(txt, "")
  Else
    TdLinkV = "<td data-v='" & Esc(txt) & "' class='nw'><a href='" & Esc(url) & "'>" & Esc(txt) & "</a></td>"
  End If
End Function

' 文件連結：優先用 NotesURL（＝文件內容 → Meta → Identifier 顯示的 notes:// 網址），
' 拿不到才用 notes://<server>/<replicaID>/0/<UNID> 自己組
Function DocLink(d)
  Dim u : u = ""
  On Error Resume Next
  u = d.NotesURL
  If Err.Number <> 0 Then u = "" : Err.Clear
  On Error Goto 0
  If Trim(u & "") = "" Then
    Dim srv : srv = Split(SERVER & "/", "/")(0)
    u = "notes://" & srv & "/" & db.ReplicaID & "/0/" & d.UniversalID
  End If
  DocLink = u
End Function

' "MM/DD/YYYY" → "YYYYMMDD"（互動版排序用）；格式不符回 ""
Function MdyToKey(s)
  MdyToKey = ""
  Dim p : p = Split(Trim(s & ""), "/")
  If UBound(p) <> 2 Then Exit Function
  MdyToKey = Right("0000" & p(2), 4) & Right("0" & p(0), 2) & Right("0" & p(1), 2)
End Function

' 報表說明段（信件內文與互動版共用）
Function HtmlIntro(scannedCount, keptCount, monthCount)
  HtmlIntro = "<h2 style='margin:0 0 6px;color:#1f3b57;'>Cost Down Project 預估報表</h2>" & _
    "<p style='color:#667;font-size:12px;margin:0 0 10px;'>" & _
    "來源：<b>" & Esc(SERVER) & "</b> !! <b>" & Esc(DBFILE) & "</b> → <b>" & Esc(VIEWNAME) & "</b><br>" & _
    "篩選：<code>Project_ID</code> 有值 且 <code>Forecast_Rls_Date</code> 有值 且 <code>Actual_Date_Final</code> 無值" & _
    IIf(STATUS_LIKE <> "", " 且 <code>Status</code> LIKE '<b>" & Esc(STATUS_LIKE) & "</b>'", "") & "<br>" & _
    "掃描 " & scannedCount & " 筆文件，保留 <b>" & keptCount & "</b> 筆，依 Forecast_Rls_Date 月份分成 <b>" & monthCount & "</b> 組。" & _
    "Project ID 可點擊，用 Notes 開回原文件。</p>"
End Function

Function HtmlStamp(what)
  HtmlStamp = "<p style='color:#999;font-size:11px;margin-top:16px;'>" & Esc(what) & "由 CostDownReport.vbs 自動產生（" & _
              FmtDate(Now) & " " & Right("0" & Hour(Now), 2) & ":" & Right("0" & Minute(Now), 2) & "）。</p>"
End Function

' 信件內文：Notes 內文 HTML 引擎陽春（無 JavaScript），一律 inline style
Function HtmlHeader(scannedCount, keptCount, monthCount)
  HtmlHeader = "<html><body style='font-family:Microsoft JhengHei,Segoe UI,Arial,sans-serif;color:#222;'>" & _
               HtmlIntro(scannedCount, keptCount, monthCount)
End Function

Function HtmlFooter()
  HtmlFooter = HtmlStamp("本信") & "</body></html>"
End Function

' 互動版 HTML：全部資料放一張平面表格，多一欄「預估月份」；標題可點排序，第二列逐欄篩選，
' 底列顯示目前篩選結果的筆數與 Monthly / Total 合計。瀏覽器開啟（Notes 內文無法執行 JavaScript）。
Function BuildInteractiveHtml(monthKeys, scannedCount, keptCount)
  Dim h
  h = "<!DOCTYPE html><html><head><meta charset='utf-8'><title>Cost Down Project 預估報表</title>" & _
      "<style>" & InteractiveCss() & "</style></head><body>" & _
      HtmlIntro(scannedCount, keptCount, UBound(monthKeys) + 1) & _
      "<div class='bar'><input id='q' type='text' placeholder='全文搜尋…'>" & _
      "<button type='button' onclick='cdClear()'>清除篩選 / 排序</button>" & _
      "<span id='cnt'></span>" & _
      "<span class='hint'>點欄位標題排序（再點反向）；第二列逐欄篩選，數字欄可輸入 &gt;1000、&lt;=500、500-2000</span></div>" & _
      "<table id='t'><thead><tr>" & _
      ThI("預估月份", "text", "select") & ThI("Project ID", "text", "") & ThI("SEC", "text", "select") & _
      ThI("Cost Pool", "text", "select") & ThI("EQP Type", "text", "select") & ThI("Project Desc.", "text", "") & _
      ThI("Method", "text", "select") & ThI("Benefit Occur (Forecast)", "date", "") & _
      ThI("Monthly Benefit", "num", "") & ThI("Total Benefit (Forecast)", "num", "") & ThI("Approve Date", "date", "") & _
      "</tr></thead><tbody>"
  Dim mi, mk, rows, ri, r
  For mi = 0 To UBound(monthKeys)
    mk = monthKeys(mi)
    rows = SortRowsBySortKey(groups(mk))
    For ri = 0 To UBound(rows)
      Set r = rows(ri)
      h = h & "<tr>" & _
        TdV(mk, "") & TdLinkV(r("project_id"), r("link")) & TdV(r("section"), "") & TdV(r("cost_pool"), "") & _
        TdV(r("eqp_type"), "") & TdV(r("project_desc"), "") & TdV(r("method"), "") & _
        TdV(r("forecast_dt"), Replace(r("forecast_dt"), "/", "")) & _
        TdNV(r("monthly")) & TdNV(r("total")) & _
        TdV(r("approve_date"), MdyToKey(r("approve_date"))) & _
        "</tr>"
    Next
  Next
  h = h & "</tbody></table>" & HtmlStamp("本檔") & "<script>" & InteractiveJs() & "</script></body></html>"
  BuildInteractiveHtml = h
End Function

Function InteractiveCss()
  Dim c : c = ""
  c = c & "body{font-family:Microsoft JhengHei,Segoe UI,Arial,sans-serif;color:#222;margin:16px}"
  c = c & ".bar{margin:10px 0;display:flex;flex-wrap:wrap;gap:8px;align-items:center;font-size:13px}"
  c = c & ".bar input{padding:4px 6px;width:220px}.bar button{padding:4px 10px}.hint{color:#667}#cnt{font-weight:bold}"
  c = c & "table{border-collapse:collapse;font-size:13px;border:1px solid #aac}"
  c = c & "th,td{border:1px solid #aac;padding:4px 8px;text-align:left;vertical-align:top}"
  c = c & "thead tr:first-child th{background:#f2f5f9;color:#1f3b57;cursor:pointer;user-select:none;white-space:nowrap;position:sticky;top:0}"
  c = c & "thead tr:first-child th:hover{background:#e3e9f2}"
  c = c & "th.asc:after{content:' \25B2'}th.desc:after{content:' \25BC'}"
  c = c & "tr.fl th{background:#fafbfd;padding:2px 4px;position:sticky;top:29px}"
  c = c & "tr.fl input,tr.fl select{width:100%;box-sizing:border-box;font-size:12px;padding:2px 4px}"
  c = c & "td.num,th.num{text-align:right;font-variant-numeric:tabular-nums}td.nw{white-space:nowrap}"
  c = c & "tbody tr:nth-child(even){background:#fbfcfe}tbody tr:hover{background:#fff7dc}"
  c = c & "tfoot td{background:#f2f5f9;font-weight:bold}a{color:#1a5fb4}"
  InteractiveCss = c
End Function

' 傳回排序後的月份鍵陣列（字典序＝時間序，因為格式 YYYY/MM）
Function SortedKeys(dict)
  If dict.Count = 0 Then
    SortedKeys = Array() : Exit Function
  End If
  Dim arr : ReDim arr(dict.Count - 1)
  Dim i : i = 0
  Dim k
  For Each k In dict.Keys
    arr(i) = k : i = i + 1
  Next
  Dim a, b, tmp
  For a = 0 To UBound(arr) - 1
    For b = 0 To UBound(arr) - 1 - a
      If arr(b) > arr(b + 1) Then
        tmp = arr(b) : arr(b) = arr(b + 1) : arr(b + 1) = tmp
      End If
    Next
  Next
  SortedKeys = arr
End Function

' 依「SEC 升冪 → 相同 SEC 內再按日期(sortkey)升冪」排序 ArrayList of Dictionary，回 Variant 陣列
Function SortRowsBySortKey(list)
  If list.Count = 0 Then
    SortRowsBySortKey = Array() : Exit Function
  End If
  Dim arr : ReDim arr(list.Count - 1)
  Dim i : i = 0
  Dim r
  For Each r In list
    Set arr(i) = r : i = i + 1
  Next
  Dim a, b, tmp
  For a = 0 To UBound(arr) - 1
    For b = 0 To UBound(arr) - 1 - a
      If RowCmp(arr(b), arr(b + 1)) > 0 Then
        Set tmp = arr(b) : Set arr(b) = arr(b + 1) : Set arr(b + 1) = tmp
      End If
    Next
  Next
  SortRowsBySortKey = arr
End Function

' 比較兩筆：先按 SEC (不分大小寫，空字串排最後)、再按 sortkey (日期)
Function RowCmp(x, y)
  Dim sx : sx = LCase(Trim(CStr(x("section") & "")))
  Dim sy : sy = LCase(Trim(CStr(y("section") & "")))
  ' 空字串排到最後
  If sx = "" And sy <> "" Then RowCmp = 1 : Exit Function
  If sy = "" And sx <> "" Then RowCmp = -1 : Exit Function
  If sx < sy Then RowCmp = -1 : Exit Function
  If sx > sy Then RowCmp = 1 : Exit Function
  If x("sortkey") < y("sortkey") Then RowCmp = -1 : Exit Function
  If x("sortkey") > y("sortkey") Then RowCmp = 1 : Exit Function
  RowCmp = 0
End Function

' 透過 Notes 用 MIME 寄 HTML 信；attachPath 非空時用 multipart/mixed 附上該檔案
Sub SendHtmlMail(mailTo, subject, htmlBody, attachPath)
  Dim mdb : Set mdb = ns.GetDatabase("", "")
  On Error Resume Next
  Call mdb.OpenMail()
  If Err.Number <> 0 Or Not mdb.IsOpen Then
    WScript.Echo "ERROR: OpenMail 失敗。" & Err.Description
    WScript.Quit 5
  End If
  On Error Goto 0
  ns.ConvertMIME = False

  Dim memo : Set memo = mdb.CreateDocument()
  Call memo.ReplaceItemValue("Form", "Memo")
  Call memo.ReplaceItemValue("SendTo", Split(mailTo, ";"))
  Call memo.ReplaceItemValue("Subject", subject)

  Dim body : Set body = memo.CreateMIMEEntity("Body")
  Dim stream

  If attachPath <> "" Then
    ' --- multipart/mixed：HTML 內文 + 附件（boundary 要明確給，部分 client 才解析得到） ---
    Dim ctHeader : Set ctHeader = body.CreateHeader("Content-Type")
    Call ctHeader.SetHeaderValAndParams("multipart/mixed; boundary=""==CDREPORTMIXED==""")

    Dim childHtml : Set childHtml = body.CreateChildEntity()
    Set stream = ns.CreateStream()
    Call stream.WriteText(htmlBody)
    Call childHtml.SetContentFromText(stream, "text/html; charset=UTF-8", ENC_IDENTITY_BINARY)
    Call stream.Close()

    Dim fname : fname = fso.GetFileName(attachPath)
    Dim childAtt : Set childAtt = body.CreateChildEntity()
    Dim dispHeader : Set dispHeader = childAtt.CreateHeader("Content-Disposition")
    Call dispHeader.SetHeaderValAndParams("attachment; filename=""" & fname & """")
    Set stream = ns.CreateStream()
    If Not stream.Open(attachPath, "binary") Then
      WScript.Echo "ERROR: 無法開啟附件檔：" & attachPath
      ns.ConvertMIME = True
      WScript.Quit 8
    End If
    ' 用 octet-stream 讓 Notes 當一般附件顯示（依副檔名 .html 用瀏覽器開），不會被當內文 inline 顯示
    Call childAtt.SetContentFromBytes(stream, "application/octet-stream; name=""" & fname & """", ENC_BASE64)
    Call stream.Close()
  Else
    ' --- 純 text/html ---
    Set stream = ns.CreateStream()
    Call stream.WriteText(htmlBody)
    Call body.SetContentFromText(stream, "text/html; charset=UTF-8", ENC_IDENTITY_BINARY)
    Call stream.Close()
  End If

  ns.ConvertMIME = True
  On Error Resume Next
  Call memo.Send(False)
  If Err.Number <> 0 Then
    WScript.Echo "ERROR: 寄信失敗。" & Err.Description
    WScript.Quit 5
  End If
  On Error Goto 0
End Sub

' 寫 UTF-8 + BOM 文字檔（SAVE_HTML 用）
Sub WriteTextUtf8(path, text)
  Dim st : Set st = CreateObject("ADODB.Stream")
  st.Type = 2 : st.Charset = "utf-8" : st.Open
  st.WriteText text
  st.SaveToFile path, 2
  st.Close
End Sub

' 相對路徑轉成以本 .vbs 所在資料夾為基準的絕對路徑
Function AbsPath(p)
  If p = "" Then AbsPath = "" : Exit Function
  If InStr(p, ":\") > 0 Or Left(p, 2) = "\\" Then
    AbsPath = p
  Else
    AbsPath = fso.BuildPath(scriptDir, p)
  End If
End Function

Sub EnsureFolder(folder)
  If folder = "" Then Exit Sub
  If Not fso.FolderExists(folder) Then
    EnsureFolder fso.GetParentFolderName(folder)
    fso.CreateFolder folder
  End If
End Sub

' 互動版 HTML 的 JavaScript（排序 / 篩選 / 合計）。字串裡只用單引號，方便放進 VBScript。
' 修改時請在瀏覽器測過再回填；此處每行對應原始 JS 一行。
Function InteractiveJs()
  Dim js : js = ""
  js = js & "(function(){" & vbLf
  js = js & "var T=document.getElementById('t'),TB=T.tBodies[0],HD=T.tHead.rows[0].cells,n=HD.length;" & vbLf
  js = js & "var rows=[].slice.call(TB.rows),q=document.getElementById('q'),fs=[],sc=-1,sd=1;" & vbLf
  js = js & "var fr=T.tHead.insertRow(-1);fr.className='fl';" & vbLf
  js = js & "var ft=T.createTFoot().insertRow(-1),fc=[];" & vbLf
  js = js & "for(var i=0;i<n;i++){" & vbLf
  js = js & " var c=fr.insertCell(-1),e;" & vbLf
  js = js & " if(HD[i].getAttribute('data-f')=='select'){" & vbLf
  js = js & "  e=document.createElement('select');var vs={},o;" & vbLf
  js = js & "  rows.forEach(function(r){vs[r.cells[i].textContent.trim()]=1;});" & vbLf
  js = js & "  o=document.createElement('option');o.value='';o.textContent='(全部)';e.appendChild(o);" & vbLf
  js = js & "  Object.keys(vs).sort().forEach(function(v){o=document.createElement('option');o.value=v;o.textContent=v||'(空白)';e.appendChild(o);});" & vbLf
  js = js & " }else{" & vbLf
  js = js & "  e=document.createElement('input');e.type='text';" & vbLf
  js = js & "  e.placeholder=HD[i].getAttribute('data-t')=='num'?'例 >1000 或 500-2000':'篩選…';" & vbLf
  js = js & " }" & vbLf
  js = js & " e.addEventListener('input',apply);e.addEventListener('change',apply);c.appendChild(e);fs.push(e);" & vbLf
  js = js & " HD[i].title='點一下排序，再點一下反向';" & vbLf
  js = js & " HD[i].addEventListener('click',(function(k){return function(){if(sc==k){sd=-sd;}else{sc=k;sd=1;}sort();};})(i));" & vbLf
  js = js & " var f=ft.insertCell(-1);if(HD[i].getAttribute('data-t')=='num'){f.className='num';}fc.push(f);" & vbLf
  js = js & "}" & vbLf
  js = js & "q.addEventListener('input',apply);" & vbLf
  js = js & "function val(r,i){var d=r.cells[i].getAttribute('data-v');return d==null?r.cells[i].textContent.trim():d;}" & vbLf
  js = js & "function numOk(v,f){" & vbLf
  js = js & " f=f.replace(/,/g,'').trim();var x=parseFloat(v)||0;" & vbLf
  js = js & " var m=f.match(/^(>=|<=|>|<|=)?\s*(-?\d+(\.\d+)?)$/);" & vbLf
  js = js & " if(m){var y=parseFloat(m[2]),op=m[1]||'=';return op=='>'?x>y:op=='<'?x<y:op=='>='?x>=y:op=='<='?x<=y:x==y;}" & vbLf
  js = js & " var g=f.match(/^(-?\d+(\.\d+)?)\s*[-~]\s*(-?\d+(\.\d+)?)$/);" & vbLf
  js = js & " if(g){return x>=parseFloat(g[1])&&x<=parseFloat(g[3]);}" & vbLf
  js = js & " return String(v).indexOf(f)>=0;" & vbLf
  js = js & "}" & vbLf
  js = js & "function fmt(x){var s=Math.round(x).toString();return s.replace(/\B(?=(\d{3})+(?!\d))/g,',');}" & vbLf
  js = js & "function apply(){" & vbLf
  js = js & " var g=q.value.trim().toLowerCase(),shown=0,sums=[];" & vbLf
  js = js & " for(var i=0;i<n;i++){sums.push(0);}" & vbLf
  js = js & " rows.forEach(function(r){" & vbLf
  js = js & "  var ok=true;" & vbLf
  js = js & "  for(var i=0;i<n&&ok;i++){" & vbLf
  js = js & "   var f=fs[i].value;if(!f){continue;}" & vbLf
  js = js & "   var t=r.cells[i].textContent.trim();" & vbLf
  js = js & "   if(fs[i].tagName=='SELECT'){ok=(t==f);}" & vbLf
  js = js & "   else if(HD[i].getAttribute('data-t')=='num'){ok=numOk(val(r,i),f);}" & vbLf
  js = js & "   else{ok=t.toLowerCase().indexOf(f.toLowerCase())>=0;}" & vbLf
  js = js & "  }" & vbLf
  js = js & "  if(ok&&g){ok=r.textContent.toLowerCase().indexOf(g)>=0;}" & vbLf
  js = js & "  r.style.display=ok?'':'none';" & vbLf
  js = js & "  if(ok){shown++;for(var k=0;k<n;k++){if(HD[k].getAttribute('data-t')=='num'){sums[k]+=parseFloat(val(r,k))||0;}}}" & vbLf
  js = js & " });" & vbLf
  js = js & " document.getElementById('cnt').textContent='顯示 '+shown+' / '+rows.length+' 筆';" & vbLf
  js = js & " fc[0].textContent='合計（'+shown+' 筆）';" & vbLf
  js = js & " for(var k=1;k<n;k++){fc[k].textContent=HD[k].getAttribute('data-t')=='num'?fmt(sums[k]):'';}" & vbLf
  js = js & "}" & vbLf
  js = js & "function sort(){" & vbLf
  js = js & " var t=HD[sc].getAttribute('data-t');" & vbLf
  js = js & " rows.sort(function(a,b){" & vbLf
  js = js & "  var x=val(a,sc),y=val(b,sc);" & vbLf
  js = js & "  if(t=='num'){return ((parseFloat(x)||0)-(parseFloat(y)||0))*sd;}" & vbLf
  js = js & "  if(x===''&&y!==''){return 1;}if(y===''&&x!==''){return -1;}" & vbLf
  js = js & "  return (x<y?-1:x>y?1:0)*sd;" & vbLf
  js = js & " });" & vbLf
  js = js & " rows.forEach(function(r){TB.appendChild(r);});" & vbLf
  js = js & " mark();" & vbLf
  js = js & "}" & vbLf
  js = js & "function mark(){for(var i=0;i<n;i++){HD[i].className=(HD[i].getAttribute('data-t')=='num'?'num ':'')+((i==sc)?(sd>0?'asc':'desc'):'');}}" & vbLf
  js = js & "window.cdClear=function(){fs.forEach(function(e){e.value='';});q.value='';sc=-1;apply();mark();};" & vbLf
  js = js & "apply();" & vbLf
  js = js & "})();" & vbLf
  InteractiveJs = js
End Function
