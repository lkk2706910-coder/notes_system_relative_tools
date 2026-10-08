Option Explicit

' FetchEcmsStatus.vbs
' 用途：從自己的 Notes 信箱找最新一封「ECMS P58 Current Online Status Report (yyyymmdd)」，
'       抽出 CSV 附件，依 EQPID 樣式篩選後，輸出成「篩選後 CSV」與「HTML 表格片段」，
'       供機況表（ASPX 網頁）讀取顯示。設計成由 Windows 工作排程器每天在信到後跑 4 次。
'       【唯讀，不修改任何信件】
'
' 前置需求：先開啟並登入 Notes client（排程執行時 Notes 也要是開著/登入狀態）。
' 執行：  cscript //nologo FetchEcmsStatus.vbs
'
' ============================ 設定區（請依你的環境修改） ============================

' --- Notes 信箱 ---
Const SERVER         = "UMCM74/UMC"            ' 信箱所在伺服器
Const DBFILE         = "mail\00047829.nsf"     ' 自己的信箱檔
Const SUBJECT_PREFIX = "ECMS P58 Current Online Status Report"   ' 主旨開頭（後面接 (yyyymmdd)）
Const LOOKBACK_DAYS  = 2                       ' 只在最近 N 天的信裡找（加快搜尋）

' --- 篩選 ---
Const EQPID_COLUMN   = "EQPID"                 ' CSV 標題列中的欄位名（不分大小寫）
Const EQPID_LIKE     = "NISACVD-B%;SACVD-B%"   ' 多個樣式用 ; 分隔；% 為萬用字元；符合任一即保留

' --- 輸出（機況表網頁要讀的檔） ---
Const OUT_DIR        = ""                      ' 留空＝本 .vbs 所在資料夾；通常填 ASPX 站台的資料夾，例 "C:\inetpub\wwwroot\status\data"
Const OUT_CSV        = "ecms_status.csv"       ' 篩選後 CSV（UTF-8 + BOM）
Const OUT_HTML       = "ecms_status.html"      ' HTML 表格片段（無 <html>/<body>，可直接嵌入）
Const KEEP_RAW_COPY  = True                    ' True＝另存原始附件到 OUT_DIR\raw\
Const CSV_CHARSET    = "utf-8"                 ' 附件 CSV 的編碼；內容有中文且亂碼時改 "big5"

' ==================================================================================

Const EMBED_ATTACHMENT = 1454
Const RICHTEXT         = 1

Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
Dim scriptDir : scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
Dim outDir : outDir = AbsPath(OUT_DIR)
If outDir = "" Then outDir = scriptDir
EnsureFolder outDir

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
  WScript.Echo "ERROR: 無法開啟信箱: [" & SERVER & "] " & DBFILE & ". " & Err.Description
  WScript.Quit 3
End If
On Error Goto 0

' ---------- 找最近 N 天內、主旨符合的信，取最新一封 ----------
Dim since : Set since = ns.CreateDateTime("Today")
Call since.AdjustDay(-LOOKBACK_DAYS)

Dim formula : formula = "@Begins(Subject; """ & Replace(SUBJECT_PREFIX, """", """""") & """)"
Dim dc
On Error Resume Next
Set dc = db.Search(formula, since, 0)
If Err.Number <> 0 Then
  WScript.Echo "ERROR: 搜尋信件失敗: " & Err.Description
  WScript.Quit 5
End If
On Error Goto 0

If dc.Count = 0 Then
  WScript.Echo "ERROR: 最近 " & LOOKBACK_DAYS & " 天內找不到主旨以「" & SUBJECT_PREFIX & "」開頭的信。"
  WScript.Quit 5
End If

Dim best, bestWhen
Set best = Nothing
Dim d : Set d = dc.GetFirstDocument()
Do Until d Is Nothing
  Dim w : w = MailWhen(d)
  If best Is Nothing Then
    Set best = d : bestWhen = w
  ElseIf w > bestWhen Then
    Set best = d : bestWhen = w
  End If
  Set d = dc.GetNextDocument(d)
Loop

Dim subject : subject = ItemText(best, "Subject")
WScript.Echo "找到 " & dc.Count & " 封，取最新：" & subject & "（" & FmtDateTime(bestWhen) & "）"

' ---------- 抽出 CSV 附件 ----------
Dim rawDir : rawDir = fso.BuildPath(outDir, "raw")
Dim tmpDir : tmpDir = fso.BuildPath(outDir, "_tmp")
EnsureFolder tmpDir
If KEEP_RAW_COPY Then EnsureFolder rawDir

Dim csvPath : csvPath = ""
Dim csvName : csvName = ""
ExtractFirstCsv best, tmpDir, csvPath, csvName
If csvPath = "" Then
  WScript.Echo "ERROR: 這封信裡沒有 .csv 附件。"
  WScript.Quit 6
End If
WScript.Echo "附件：" & csvName

If KEEP_RAW_COPY Then
  On Error Resume Next
  fso.CopyFile csvPath, fso.BuildPath(rawDir, FmtStamp(bestWhen) & "_" & csvName), True
  On Error Goto 0
End If

' ---------- 解析 + 篩選 ----------
Dim text : text = ReadText(csvPath, CSV_CHARSET)
Dim delim : delim = DetectDelim(text)
Dim rows : Set rows = ParseDelimited(text, delim)
If rows.Count < 2 Then
  WScript.Echo "ERROR: CSV 沒有資料列（只有標題或是空檔）。"
  WScript.Quit 7
End If

Dim header : header = rows.Item(0)
Dim eqpIdx : eqpIdx = FindCol(header, EQPID_COLUMN)
If eqpIdx < 0 Then
  WScript.Echo "ERROR: CSV 標題列找不到欄位「" & EQPID_COLUMN & "」。實際標題：" & Join(header, " | ")
  WScript.Quit 7
End If

Dim patterns : patterns = Split(EQPID_LIKE, ";")
Dim kept : Set kept = CreateObject("System.Collections.ArrayList")
Dim i, r, v, p, hit
For i = 1 To rows.Count - 1
  r = rows.Item(i)
  If UBound(r) >= eqpIdx Then
    v = Trim(r(eqpIdx))
    hit = False
    For Each p In patterns
      If Trim(p) <> "" Then
        If LikeMatchPct(v, Trim(p)) Then hit = True : Exit For
      End If
    Next
    If hit Then kept.Add r
  End If
Next

' ---------- 輸出 ----------
Dim outCsv : outCsv = fso.BuildPath(outDir, OUT_CSV)
Dim outHtml : outHtml = fso.BuildPath(outDir, OUT_HTML)

Dim csvOut : csvOut = CsvRow(header) & vbCrLf
For Each r In kept : csvOut = csvOut & CsvRow(r) & vbCrLf : Next

Dim html : html = ""
html = html & "<div class='ecms-status'>" & vbCrLf
html = html & "  <div class='ecms-meta' style='font-size:12px;color:#667;margin:4px 0;'>" & _
              "來源：" & Esc(subject) & "｜收信 " & FmtDateTime(bestWhen) & _
              "｜更新 " & FmtDateTime(Now) & "｜篩選 EQPID LIKE " & Esc(Replace(EQPID_LIKE, ";", " / ")) & _
              "｜共 " & kept.Count & " 筆</div>" & vbCrLf
html = html & "  <table class='ecms-table' cellspacing='0' cellpadding='4' style='border-collapse:collapse;font-size:13px;'>" & vbCrLf
html = html & "    <thead><tr>"
Dim c
For c = 0 To UBound(header) : html = html & Th(header(c)) : Next
html = html & "</tr></thead>" & vbCrLf & "    <tbody>" & vbCrLf
For Each r In kept
  html = html & "      <tr>"
  For c = 0 To UBound(header)
    If c <= UBound(r) Then html = html & Td(r(c)) Else html = html & Td("")
  Next
  html = html & "</tr>" & vbCrLf
Next
html = html & "    </tbody>" & vbCrLf & "  </table>" & vbCrLf & "</div>" & vbCrLf

On Error Resume Next
WriteTextUtf8 outCsv, csvOut
WriteTextUtf8 outHtml, html
If Err.Number <> 0 Then
  WScript.Echo "ERROR: 寫入輸出檔失敗（檔案可能被網站/Excel 鎖住）: " & Err.Description
  WScript.Quit 8
End If
fso.DeleteFile csvPath, True
On Error Goto 0

WScript.Echo "OK 原始 " & (rows.Count - 1) & " 筆 → 篩選後 " & kept.Count & " 筆；已輸出 " & outCsv & " 與 " & outHtml
WScript.Quit 0

' ================================ 函式 ================================

' 信件時間：優先 DeliveredDate / PostedDate，否則文件建立時間
Function MailWhen(doc)
  Dim s, x
  s = ItemText(doc, "DeliveredDate")
  If Trim(s) = "" Then s = ItemText(doc, "PostedDate")
  On Error Resume Next
  If Trim(s) <> "" Then x = CDate(s)
  If Err.Number <> 0 Or Trim(s) = "" Then
    Err.Clear
    x = CDate(doc.Created)
  End If
  On Error Goto 0
  MailWhen = x
End Function

' 從信件抽出第一個 .csv 附件到 dir；回傳路徑與檔名（ByRef）
Sub ExtractFirstCsv(doc, dir, ByRef outPath, ByRef outName)
  outPath = "" : outName = ""
  Dim eos, eo, it, reos
  On Error Resume Next
  eos = Empty : eos = doc.EmbeddedObjects
  If IsArray(eos) Then
    For Each eo In eos
      If TryExtract(eo, dir, outPath, outName) Then Exit Sub
    Next
  End If
  Err.Clear
  For Each it In doc.Items
    If it.Type = RICHTEXT Then
      reos = Empty : reos = it.EmbeddedObjects
      If IsArray(reos) Then
        For Each eo In reos
          If TryExtract(eo, dir, outPath, outName) Then Exit Sub
        Next
      End If
    End If
    Err.Clear
  Next
  On Error Goto 0
End Sub

Function TryExtract(eo, dir, ByRef outPath, ByRef outName)
  TryExtract = False
  On Error Resume Next
  If eo.Type <> EMBED_ATTACHMENT Then Exit Function
  Dim nm : nm = eo.Source
  If LCase(fso.GetExtensionName(nm)) <> "csv" Then Exit Function
  Dim pth : pth = fso.BuildPath(dir, nm)
  If fso.FileExists(pth) Then fso.DeleteFile pth, True
  eo.ExtractFile pth
  If Err.Number <> 0 Then Err.Clear : Exit Function
  outPath = pth : outName = nm
  TryExtract = True
End Function

' 讀文字檔（指定編碼）
Function ReadText(path, charset)
  Dim st : Set st = CreateObject("ADODB.Stream")
  st.Type = 2 : st.Charset = charset : st.Open
  st.LoadFromFile path
  ReadText = st.ReadText
  st.Close
  If Len(ReadText) > 0 Then
    If AscW(Left(ReadText, 1)) = &HFEFF Then ReadText = Mid(ReadText, 2)
  End If
End Function

' 偵測分隔：第一個非空行裡逗號/Tab/分號哪個最多
Function DetectDelim(text)
  Dim lines : lines = Split(Replace(text, vbCr, ""), vbLf)
  Dim ln, i
  For i = 0 To UBound(lines)
    If Trim(lines(i)) <> "" Then ln = lines(i) : Exit For
  Next
  Dim nc : nc = Len(ln) - Len(Replace(ln, ",", ""))
  Dim nt : nt = Len(ln) - Len(Replace(ln, vbTab, ""))
  Dim ns2 : ns2 = Len(ln) - Len(Replace(ln, ";", ""))
  DetectDelim = ","
  If nt > nc And nt >= ns2 Then DetectDelim = vbTab
  If ns2 > nc And ns2 > nt Then DetectDelim = ";"
End Function

' 引號感知的分隔文字解析；回傳 ArrayList，每筆為 Variant 陣列（已略過整列空白）
Function ParseDelimited(text, delim)
  Dim rows : Set rows = CreateObject("System.Collections.ArrayList")
  Dim fields : Set fields = CreateObject("System.Collections.ArrayList")
  Dim field : field = ""
  Dim inQ : inQ = False
  Dim i, ch, n : n = Len(text)
  i = 1
  Do While i <= n
    ch = Mid(text, i, 1)
    If inQ Then
      If ch = """" Then
        If Mid(text, i + 1, 1) = """" Then
          field = field & """" : i = i + 2
        Else
          inQ = False : i = i + 1
        End If
      Else
        field = field & ch : i = i + 1
      End If
    Else
      If ch = """" Then
        inQ = True : i = i + 1
      ElseIf ch = delim Then
        fields.Add field : field = "" : i = i + 1
      ElseIf ch = vbCr Then
        i = i + 1
      ElseIf ch = vbLf Then
        fields.Add field : field = ""
        AddRowIfNotBlank rows, fields
        Set fields = CreateObject("System.Collections.ArrayList")
        i = i + 1
      Else
        field = field & ch : i = i + 1
      End If
    End If
  Loop
  If field <> "" Or fields.Count > 0 Then
    fields.Add field
    AddRowIfNotBlank rows, fields
  End If
  Set ParseDelimited = rows
End Function

Sub AddRowIfNotBlank(rows, fields)
  Dim arr() : ReDim arr(fields.Count - 1)
  Dim k : k = 0
  Dim f, anyVal : anyVal = False
  For Each f In fields
    arr(k) = Trim(CStr(f & "")) : If arr(k) <> "" Then anyVal = True
    k = k + 1
  Next
  If anyVal Then rows.Add arr
End Sub

' 在標題陣列中找欄位（不分大小寫、去空白）；找不到回 -1
Function FindCol(header, name)
  FindCol = -1
  Dim k
  For k = 0 To UBound(header)
    If UCase(Trim(header(k))) = UCase(Trim(name)) Then FindCol = k : Exit Function
  Next
End Function

' % 萬用字元樣式比對（不分大小寫，整串）
Function LikeMatchPct(s, pattern)
  If pattern = "" Then LikeMatchPct = True : Exit Function
  Dim rx : Set rx = CreateObject("VBScript.RegExp")
  rx.IgnoreCase = True : rx.Global = False
  Dim out, i, c : out = "^"
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

Function ItemText(d, fld)
  ItemText = ""
  If Not d.HasItem(fld) Then Exit Function
  Dim it : Set it = d.GetFirstItem(fld)
  If it Is Nothing Then Exit Function
  ItemText = it.Text
End Function

Function FmtDateTime(v)
  If Not IsDate(v) Then FmtDateTime = "" : Exit Function
  Dim x : x = CDate(v)
  FmtDateTime = Right("0000" & Year(x), 4) & "/" & Right("0" & Month(x), 2) & "/" & Right("0" & Day(x), 2) & _
                " " & Right("0" & Hour(x), 2) & ":" & Right("0" & Minute(x), 2)
End Function

Function FmtStamp(v)
  Dim x : x = CDate(v)
  FmtStamp = Right("0000" & Year(x), 4) & Right("0" & Month(x), 2) & Right("0" & Day(x), 2) & "_" & _
             Right("0" & Hour(x), 2) & Right("0" & Minute(x), 2)
End Function

Function Esc(s)
  Esc = Replace(CStr(s & ""), "&", "&amp;")
  Esc = Replace(Esc, "<", "&lt;")
  Esc = Replace(Esc, ">", "&gt;")
  Esc = Replace(Esc, """", "&quot;")
End Function

Function Th(t)
  Th = "<th style='border:1px solid #aac;background:#f2f5f9;color:#1f3b57;padding:4px 8px;text-align:left;'>" & Esc(t) & "</th>"
End Function

Function Td(t)
  Td = "<td style='border:1px solid #aac;padding:3px 8px;'>" & Esc(t) & "</td>"
End Function

Function CsvRow(arr)
  Dim out, k, cell : out = ""
  For k = LBound(arr) To UBound(arr)
    cell = CStr(arr(k) & "")
    If InStr(cell, ",") > 0 Or InStr(cell, """") > 0 Or InStr(cell, vbCr) > 0 Or InStr(cell, vbLf) > 0 Then
      cell = """" & Replace(cell, """", """""") & """"
    End If
    If k > LBound(arr) Then out = out & ","
    out = out & cell
  Next
  CsvRow = out
End Function

Sub WriteTextUtf8(path, text)
  EnsureFolder fso.GetParentFolderName(path)
  Dim st : Set st = CreateObject("ADODB.Stream")
  st.Type = 2 : st.Charset = "utf-8" : st.Open
  st.WriteText text
  st.SaveToFile path, 2
  st.Close
End Sub

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
