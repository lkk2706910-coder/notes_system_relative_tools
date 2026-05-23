Option Explicit

' TestFindDoc.vbs
' 測試用：用文件編號在 By Document No. 檢視查找文件，列出命中文件與其附件 (不下載)。
' 用法：cscript //nologo TestFindDoc.vbs

Const SERVER     = "UMCD06/UMC"
Const DBFILE     = "doclib\cmp.nsf"
Const DOCNO      = "SHCP-DPX3NC"
Const DOCNO_VIEW = "All Document\3. By Document No."

Const EMBED_ATTACHMENT = 1454
Const RICHTEXT = 1

Dim session : Set session = CreateObject("Notes.NotesSession")
If Err.Number <> 0 Then WScript.Echo "ERROR: 無法建立 Notes session (Notes 未開/未登入?) " & Err.Description : WScript.Quit 4

Dim db : Set db = session.GetDatabase(SERVER, DBFILE)
If db Is Nothing Then WScript.Echo "ERROR: 找不到資料庫 " & SERVER & " !! " & DBFILE : WScript.Quit 3
On Error Resume Next
If Not db.IsOpen Then db.Open "", ""
If Err.Number <> 0 Or Not db.IsOpen Then WScript.Echo "ERROR: 無法開啟資料庫 (權限/路徑?) " & Err.Description : WScript.Quit 3
On Error Goto 0

Dim view : Set view = db.GetView(DOCNO_VIEW)
If view Is Nothing Then WScript.Echo "ERROR: 找不到檢視 " & DOCNO_VIEW : WScript.Quit 5

Dim dc
On Error Resume Next
Set dc = view.GetAllDocumentsByKey(DOCNO, True)
If Err.Number <> 0 Then WScript.Echo "ERROR: 依文件編號查找失敗 " & Err.Description : WScript.Quit 8
On Error Goto 0

WScript.Echo "依文件編號 [" & DOCNO & "] 找到 " & dc.Count & " 份文件"
WScript.Echo "--------------------------------------------------------------"

Dim doc : Set doc = dc.GetFirstDocument()
Dim idx : idx = 0
Do Until doc Is Nothing
  idx = idx + 1
  WScript.Echo "#" & idx & "  NoteID=" & doc.NoteID & "  建立=" & doc.Created & "  HasEmbedded=" & doc.HasEmbedded

  Dim seenAtt : Set seenAtt = CreateObject("Scripting.Dictionary")

  ' 方法 A：文失級 EmbeddedObjects
  On Error Resume Next
  Dim eos, eo
  eos = doc.EmbeddedObjects
  If IsArray(eos) Then
    For Each eo In eos
      If eo.Type = EMBED_ATTACHMENT Then
        If Not seenAtt.Exists(eo.Source) Then
          seenAtt.Add eo.Source, True
          WScript.Echo "    [A] 附件: " & eo.Source
        End If
      End If
    Next
  End If
  Err.Clear
  On Error Goto 0

  ' 方法 B：逐個 RichText 欄位的 EmbeddedObjects (備援)
  Dim it, reos, e2
  For Each it In doc.Items
    On Error Resume Next
    If it.Type = RICHTEXT Then
      reos = Empty
      reos = it.EmbeddedObjects
      If IsArray(reos) Then
        For Each e2 In reos
          If e2.Type = EMBED_ATTACHMENT Then
            If Not seenAtt.Exists(e2.Source) Then
              seenAtt.Add e2.Source, True
              WScript.Echo "    [B] 附件(欅位 " & it.Name & "): " & e2.Source
            End If
          End If
        Next
      End If
    End If
    Err.Clear
    On Error Goto 0
  Next

  If seenAtt.Count = 0 Then WScript.Echo "    (未列出附件；若 HasEmbedded=True 再謬結果給我換方泙)"
  Set doc = dc.GetNextDocument(doc)
Loop

WScript.Echo "--------------------------------------------------------------"
If dc.Count = 0 Then
  WScript.Echo "找不到。可能：編號不宊全相同，或該檂覞第踈不是文失編號 (要改唨逎筅毟對)"
Else
  WScript.Echo "OK"
End If
WScript.Quit 0
