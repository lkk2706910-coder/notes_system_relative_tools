Option Explicit

' ListNotesViews.vbs
' 列出指定 Notes 資料庫的所有檢視 / 資料夾名稱與別名，
' 方便決定 DownloadNotesAttachments.vbs 的 VIEWNAME 該填什麼。
'
' 用法：cscript //nologo ListNotesViews.vbs

Const SERVER = "UMCD06/UMC"
Const DBFILE = "doclib\cmp.nsf"

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

WScript.Echo "資料庫 " & SERVER & " !! " & DBFILE & " 的檢視 / 資料夾："
WScript.Echo "（VIEWNAME 可填 Name 或任一 Alias）"
WScript.Echo "------------------------------------------------------------"

Dim v
On Error Resume Next
For Each v In db.Views
  WScript.Echo "Name=[" & v.Name & "]"
  WScript.Echo "    Alias=[" & AliasesToStr(v.Aliases) & "]  IsFolder=" & v.IsFolder
  If Err.Number <> 0 Then Err.Clear
Next
On Error Goto 0

WScript.Echo "------------------------------------------------------------"
WScript.Echo "OK"
WScript.Quit 0


' 把 Aliases 串成字串；任何讀取問題都吞掉，確保 Name 仍能列出
Function AliasesToStr(a)
  On Error Resume Next
  Dim s, i
  s = ""
  If IsArray(a) Then
    For i = LBound(a) To UBound(a)
      If i > LBound(a) Then s = s & ", "
      s = s & a(i)          ' & 把 Null 當空字串
    Next
  Else
    s = "" & a
  End If
  If Err.Number <> 0 Then
    s = ""
    Err.Clear
  End If
  AliasesToStr = s
  On Error Goto 0
End Function
