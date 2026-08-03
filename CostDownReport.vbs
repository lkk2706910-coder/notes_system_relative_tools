Option Explicit

' CostDownReport.vbs
' 用途：從 CDPAS 資料庫的檢視「6. 依分類 (Last 2 Years)」讀取 cost-down 項目，
'       篩選「Forecast_Rls_Date 有值 且 Actual_Date_Final 無值」的資料，
'       依 Forecast_Rls_Date 月份分組整理成 HTML 表格，透過 Notes 寄信到指定收件人。
'       Project ID 會做成 Notes 文件連結 (Copy as Document Link 的效果)，可點擊回原文件。
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
Const F_STATION_T11  = "Station Time11"       ' 值格式如 "11^New Project - Approved^07/24/2026 16:25^^^"，取第 3 段的日期部分

' --- 寄信 ---
Const MAIL_TO      = "bo_hsiang_kao@umc.com"          ' 收件人，多人分號分隔
Const MAIL_SUBJECT = "Cost Down Project 預估報表（未實際結案）"

' --- 其他 ---
Const DRY_RUN         = False    ' True＝不寄信，改把 HTML 存到 OUTFILE（測試用）
Const OUTFILE         = "cost_down_report.html"
Const PROGRESS_EVERY  = 200

' ==================================================================================

' MIME encoding constant (見 HCL Notes 文件)
Const ENC_IDENTITY_BINARY = 1729

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

' ---------- 組 HTML ----------
Dim monthKeys : monthKeys = SortedKeys(groups)
Dim html : html = HtmlHeader(scanned, kept, UBound(monthKeys) + 1)
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
      Td(r("project_id")) & Td(r("section")) & Td(r("cost_pool")) & Td(r("eqp_type")) & _
      Td(r("project_desc")) & Td(r("method")) & Td(r("forecast_dt")) & _
      TdN(r("monthly")) & TdN(r("total")) & Td(r("approve_date")) & _
      "</tr>"
  Next
  html = html & "</tbody></table>"
Next
html = html & HtmlFooter()

' ---------- 寄信 / DRY RUN ----------
If DRY_RUN Then
  Dim outPath : outPath = AbsPath(OUTFILE)
  EnsureFolder fso.GetParentFolderName(outPath)
  WriteTextUtf8 outPath, html
  WScript.Echo "DRY_RUN=True，已把 HTML 存到：" & outPath & "（未寄信）"
Else
  SendHtmlMail MAIL_TO, MAIL_SUBJECT, html
  WScript.Echo "OK 已寄出報表到：" & MAIL_TO & "（月份區塊 " & (UBound(monthKeys) + 1) & " 個，共 " & kept & " 筆）"
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

Function HtmlHeader(scannedCount, keptCount, monthCount)
  HtmlHeader = "<html><body style='font-family:Microsoft JhengHei,Segoe UI,Arial,sans-serif;color:#222;'>" & _
    "<h2 style='margin:0 0 6px;color:#1f3b57;'>Cost Down Project 預估報表</h2>" & _
    "<p style='color:#667;font-size:12px;margin:0 0 10px;'>" & _
    "來源：<b>" & Esc(SERVER) & "</b> !! <b>" & Esc(DBFILE) & "</b> → <b>" & Esc(VIEWNAME) & "</b><br>" & _
    "篩選：<code>Project_ID</code> 有值 且 <code>Forecast_Rls_Date</code> 有值 且 <code>Actual_Date_Final</code> 無值" & _
    IIf(STATUS_LIKE <> "", " 且 <code>Status</code> LIKE '<b>" & Esc(STATUS_LIKE) & "</b>'", "") & "<br>" & _
    "掃描 " & scannedCount & " 筆文件，保留 <b>" & keptCount & "</b> 筆，依 Forecast_Rls_Date 月份分成 <b>" & monthCount & "</b> 組。</p>"
End Function

Function HtmlFooter()
  HtmlFooter = "<p style='color:#999;font-size:11px;margin-top:16px;'>本信由 CostDownReport.vbs 自動產生（" & _
               FmtDate(Now) & " " & Right("0" & Hour(Now), 2) & ":" & Right("0" & Minute(Now), 2) & "）。</p></body></html>"
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

' 透過 Notes 用 MIME 寄 HTML 信
Sub SendHtmlMail(mailTo, subject, htmlBody)
  Dim mdb : Set mdb = ns.GetDatabase("", "")
  Call mdb.OpenMail()
  If Err.Number <> 0 Then
    WScript.Echo "ERROR: OpenMail 失敗。" & Err.Description
    WScript.Quit 5
  End If
  ns.ConvertMIME = False

  Dim memo : Set memo = mdb.CreateDocument()
  Call memo.ReplaceItemValue("Form", "Memo")
  Call memo.ReplaceItemValue("SendTo", Split(mailTo, ";"))
  Call memo.ReplaceItemValue("Subject", subject)

  Dim body : Set body = memo.CreateMIMEEntity("Body")
  Dim stream : Set stream = ns.CreateStream()
  Call stream.WriteText(htmlBody)
  Call body.SetContentFromText(stream, "text/html; charset=UTF-8", ENC_IDENTITY_BINARY)
  Call stream.Close()

  ns.ConvertMIME = True
  Call memo.Send(False)
End Sub

' 寫 UTF-8 文字檔（DRY_RUN 模式用）
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
