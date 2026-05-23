Option Explicit

' DownloadNotesAttachments.vbs
' 從 Notes/Domino 資料庫的文件抽出指定副檔名的附件 (預設 Excel) 到本機資料夾。
'
' 用法：cscript //nologo DownloadNotesAttachments.vbs
' 設定改下方「設定區」。需 Notes client 已開啟且已登入，且帳號對該庫有讀取權。

' ====================== 設定區 ======================
Const SERVER   = "UMCD06/UMC"            ' Domino server
Const DBFILE   = "doclib\cmp.nsf"        ' 資料庫路徑 (相對 Notes data 目錄)

' 挑文件的方式 (依序判斷，擇一)：
'  1) DOCNO        ：用「文件編號」在 By Document No. 檢視精準比對 (最快，ASCII 免編碼問題)
'  2) SUBJECTFILE  ：把主旨/關鍵字放 UTF-8 檔，用全文檢索 FTSearch (含中文走這條)
'  3) VIEWNAME     ：指定檢視，抓該檢視全部文件
'  4) SEARCHFORMULA：Notes 公式
'  5) 以上皆空      ：全部文件
Const DOCNO      = "SHCP-DPX3NC"         ' 文件編號 (精準比對)；留空=不使用
Const DOCNO_VIEW = "All Document\3. By Document No."   ' 依文件編號排序的檢視
Const SUBJECTFILE   = ""                 ' UTF-8 檔 (主旨/關鍵字, FTSearch)；留空=不使用
Const VIEWNAME      = ""                 ' 例：All Document\5. By Subject
Const SEARCHFORMULA = ""                 ' 例：Subject = ""SACVD 月報""

Const TARGETDIR = "C:\NotesDownload"     ' 存檔資料夾 (不存在會自動建立)
Const EXTLIST   = ".xls;.xlsx;.xlsm"     ' 要下載的副檔名 (分號分隔，小寫)
Const OVERWRITE = True                   ' True=同名覆寫(先刪舊檔)；False=同名時加 NoteID 前綴
' ===================================================

Const EMBED_ATTACHMENT = 1454
Const RICHTEXT = 1

Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
Dim seen : Set seen = CreateObject("Scripting.Dictionary")   ' 去重 (分類檢視同文件會出現多次)
CreateFolderRecursive TARGETDIR

Dim session : Set session = CreateObject("Notes.NotesSession")
If Err.Number <> 0 Then
  WScript.Echo "ERROR: 無法建立 Notes.NotesSession (Notes 未開/未登入?) " & Err.Description
  WScript.Quit 4
End If

Dim db : Set db = session.GetDatabase(SERVER, DBFILE)
If db Is Nothing Then
  WScript.Echo "ERROR: 找不到資料庫 " & SERVER & " !! " & DBFILE
  WScript.Quit 3
End If
On Error Resume Next
If Not db.IsOpen Then db.Open "", ""
If Err.Number <> 0 Or Not db.IsOpen Then
  WScript.Echo "ERROR: 無法開啟資料庫 (權限/路徑?) " & Err.Description
  WScript.Quit 3
End If
On Error Goto 0

' 取得要處理的文件來源
Dim usingView : usingView = False
Dim view, dc, doc
If Trim(DOCNO) <> "" Then
  Set view = db.GetView(DOCNO_VIEW)
  If view Is Nothing Then
    WScript.Echo "ERROR: 找不到檢視 " & DOCNO_VIEW
    WScript.Quit 5
  End If
  On Error Resume Next
  Set dc = view.GetAllDocumentsByKey(DOCNO, True)
  If Err.Number <> 0 Then
    WScript.Echo "ERROR: 依文件編號查找失敗 " & Err.Description
    WScript.Quit 8
  End If
  On Error Goto 0
  WScript.Echo "依文件編號 [" & DOCNO & "] 找到 " & dc.Count & " 份文件"
  Set doc = dc.GetFirstDocument()
ElseIf Trim(SUBJECTFILE) <> "" Then
  Dim qpath : qpath = ResolvePath(SUBJECTFILE)
  If Not fso.FileExists(qpath) Then
    WScript.Echo "ERROR: 找不到關鍵字檔 " & qpath
    WScript.Quit 6
  End If
  Dim query : query = ReadTextUtf8(qpath)
  WScript.Echo "FTSearch 關鍵字：[" & query & "]"
  On Error Resume Next
  Set dc = db.FTSearch(query, 0)
  If Err.Number <> 0 Then
    WScript.Echo "ERROR: FTSearch 失敗 (資料庫可能未建全文索引) " & Err.Description
    WScript.Quit 7
  End If
  On Error Goto 0
  WScript.Echo "FTSearch 命中 " & dc.Count & " 份文件"
  Set doc = dc.GetFirstDocument()
ElseIf Trim(VIEWNAME) <> "" Then
  Set view = db.GetView(VIEWNAME)
  If view Is Nothing Then
    WScript.Echo "ERROR: 找不到檢視 " & VIEWNAME
    WScript.Quit 5
  End If
  usingView = True
  Set doc = view.GetFirstDocument()
ElseIf Trim(SEARCHFORMULA) <> "" Then
  Set dc = db.Search(SEARCHFORMULA, Nothing, 0)
  Set doc = dc.GetFirstDocument()
Else
  Set dc = db.AllDocuments
  Set doc = dc.GetFirstDocument()
End If

Dim total, saved
total = 0 : saved = 0

Dim attSeen, eos, eo, it, reos, e2
Do Until doc Is Nothing
  If Not seen.Exists(doc.UniversalID) Then
    seen.Add doc.UniversalID, True
    total = total + 1
    Set attSeen = CreateObject("Scripting.Dictionary")   ' 同文件附件去重 (by 檔名)

    ' 方法 A：文件層級 EmbeddedObjects
    On Error Resume Next
    eos = Empty
    eos = doc.EmbeddedObjects
    If IsArray(eos) Then
      For Each eo In eos
        ExtractIfWanted eo, doc, attSeen
      Next
    End If
    Err.Clear
    On Error Goto 0

    ' 方法 B：逐個 RichText 欄位 EmbeddedObjects (附件常在 content/Body 等欄位)
    For Each it In doc.Items
      On Error Resume Next
      If it.Type = RICHTEXT Then
        reos = Empty
        reos = it.EmbeddedObjects
        If IsArray(reos) Then
          For Each e2 In reos
            ExtractIfWanted e2, doc, attSeen
          Next
        End If
      End If
      Err.Clear
      On Error Goto 0
    Next
  End If

  If usingView Then
    Set doc = view.GetNextDocument(doc)
  Else
    Set doc = dc.GetNextDocument(doc)
  End If
Loop

WScript.Echo "OK 掃描 " & total & " 份文件，下載 " & saved & " 個附件到 " & TARGETDIR
WScript.Quit 0

' 若是要的副檔名附件就抽出 (eo=NotesEmbeddedObject)
Sub ExtractIfWanted(eo, doc, attSeen)
  On Error Resume Next
  If eo.Type <> EMBED_ATTACHMENT Then Exit Sub
  Dim nm : nm = eo.Source
  If nm = "" Then Exit Sub
  If Not MatchExt(LCase(nm)) Then Exit Sub
  If attSeen.Exists(nm) Then Exit Sub
  attSeen.Add nm, True
  Dim outpath : outpath = TARGETDIR & "\" & nm
  If fso.FileExists(outpath) Then
    If OVERWRITE Then
      fso.DeleteFile outpath, True       ' 覆蓋：先刪舊檔 (ExtractFile 遇同名會失敗)
    Else
      outpath = TARGETDIR & "\" & doc.NoteID & "_" & nm
    End If
  End If
  eo.ExtractFile outpath
  If Err.Number = 0 Then
    saved = saved + 1
    WScript.Echo "  saved: " & outpath
  Else
    WScript.Echo "  WARN: 抽取失敗 " & nm & " - " & Err.Description
    Err.Clear
  End If
End Sub


' ---------- helpers ----------
Function MatchExt(lname)
  Dim arr, e
  arr = Split(LCase(EXTLIST), ";")
  MatchExt = False
  For Each e In arr
    e = Trim(e)
    If e <> "" Then
      If Len(lname) >= Len(e) Then
        If Right(lname, Len(e)) = e Then
          MatchExt = True
          Exit Function
        End If
      End If
    End If
  Next
End Function

Sub CreateFolderRecursive(path)
  If path = "" Then Exit Sub
  If fso.FolderExists(path) Then Exit Sub
  Dim parent : parent = fso.GetParentFolderName(path)
  If parent <> "" And Not fso.FolderExists(parent) Then CreateFolderRecursive parent
  fso.CreateFolder path
End Sub

' 以 UTF-8 讀文字檔 (cscript 直接讀 .vbs 是 Big5，故中文走檔案 + ADODB.Stream)
Function ReadTextUtf8(path)
  Dim st : Set st = CreateObject("ADODB.Stream")
  st.Type = 2
  st.Charset = "utf-8"
  st.Open
  st.LoadFromFile path
  Dim t : t = st.ReadText
  st.Close
  Do While Len(t) > 0 And (Right(t, 1) = vbCr Or Right(t, 1) = vbLf)
    t = Left(t, Len(t) - 1)
  Loop
  ReadTextUtf8 = t
End Function

' 相對路徑 → 相對於本 .vbs 所在資料夾；絕對路徑直接用
Function ResolvePath(p)
  If InStr(p, ":\") > 0 Or Left(p, 2) = "\\" Then
    ResolvePath = p
  Else
    ResolvePath = fso.BuildPath(fso.GetParentFolderName(WScript.ScriptFullName), p)
  End If
End Function
