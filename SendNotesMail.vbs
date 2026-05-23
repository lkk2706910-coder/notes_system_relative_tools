Option Explicit

' SendNotesMail.vbs (MIME 版)
' 從 HTML 檔讀取內容，用 MIME 方式寄出，信件內文直接渲染表格。
' 可選帶一張內嵌圖片 (PNG)，以 multipart/related + Content-ID <trendchart> 內嵌，
' HTML 端用 <img src="cid:trendchart"> 參照。
'
' Usage:
'   cscript //nologo SendNotesMail.vbs "<TO>" "<SUBJECT>" "<HTML_FILE_PATH>" ["<IMAGE_FILE_PATH>"]
'   TO 多人時用分號分隔，例如 "A/UMC;B/UMC"
'   第 4 個參數 (PNG 圖片路徑) 可省略；省略時等同舊版純 text/html 行為。

' MIME 編碼常數 (見 HCL Notes 文件)
Const ENC_BASE64        = 1725
Const ENC_IDENTITY_BINARY = 1729   ' 與舊版一致 (舊註解誤標為 7BIT)
Const IMAGE_CID = "trendchart"

Dim args : Set args = WScript.Arguments
If args.Count < 3 Then
  WScript.Echo "Usage: cscript //nologo SendNotesMail.vbs ""<TO>"" ""<SUBJECT>"" ""<HTML_FILE_PATH>"" [""<IMAGE_FILE_PATH>""]"
  WScript.Quit 2
End If

Dim mailTo, subject, htmlPath, imagePath
mailTo    = args(0)
subject   = args(1)
htmlPath  = args(2)
imagePath = ""
If args.Count >= 4 Then imagePath = args(3)

' 讀 HTML 檔 (UTF-8)
Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
If Not fso.FileExists(htmlPath) Then
  WScript.Echo "ERROR: HTML file not found: " & htmlPath
  WScript.Quit 3
End If

Dim adoStream : Set adoStream = CreateObject("ADODB.Stream")
adoStream.Type = 2                ' text mode
adoStream.Charset = "utf-8"
adoStream.Open
adoStream.LoadFromFile htmlPath
Dim htmlBody : htmlBody = adoStream.ReadText
adoStream.Close

Dim hasImage : hasImage = (imagePath <> "" And fso.FileExists(imagePath))

On Error Resume Next

Dim notesSession : Set notesSession = CreateObject("Notes.NotesSession")
If Err.Number <> 0 Then
  WScript.Echo "ERROR: Cannot create Notes.NotesSession. Err=" & Err.Number & " " & Err.Description
  WScript.Quit 4
End If

Dim db : Set db = notesSession.GetDatabase("", "")
Call db.OpenMail()
If Err.Number <> 0 Then
  WScript.Echo "ERROR: OpenMail failed. Err=" & Err.Number & " " & Err.Description
  WScript.Quit 5
End If

' 關閉自動 MIME 轉換 (讓我們自己控制)
notesSession.ConvertMIME = False

Dim doc : Set doc = db.CreateDocument()
Call doc.ReplaceItemValue("Form", "Memo")

' 收件人 (用分號拆成陣列, 支援多人)
Dim toArr : toArr = Split(mailTo, ";")
Call doc.ReplaceItemValue("SendTo", toArr)
Call doc.ReplaceItemValue("Subject", subject)

Dim body : Set body = doc.CreateMIMEEntity("Body")
Dim stream

If hasImage Then
  ' --- multipart/related：HTML + 內嵌圖片 ---
  ' 必須給明確 boundary，否則部分 client 解析不到內嵌圖 (cid 變破圖)
  Dim ctHeader : Set ctHeader = body.CreateHeader("Content-Type")
  Call ctHeader.SetHeaderValAndParams("multipart/related; boundary=""==SACVDRELATED=="";type=""text/html""")

  ' child 1: HTML (引用 cid:trendchart)
  Dim childHtml : Set childHtml = body.CreateChildEntity()
  Set stream = notesSession.CreateStream()
  Call stream.WriteText(htmlBody)
  Call childHtml.SetContentFromText(stream, "text/html; charset=UTF-8", ENC_IDENTITY_BINARY)
  Call stream.Close()

  ' child 2: 圖片 (inline, 以 Content-ID 對應 cid:trendchart)
  Dim childImg : Set childImg = body.CreateChildEntity()
  Dim cidHeader : Set cidHeader = childImg.CreateHeader("Content-ID")
  Call cidHeader.SetHeaderVal("<" & IMAGE_CID & ">")
  Dim dispHeader : Set dispHeader = childImg.CreateHeader("Content-Disposition")
  Call dispHeader.SetHeaderValAndParams("inline; filename=""trend.png""")
  Set stream = notesSession.CreateStream()
  If Not stream.Open(imagePath) Then
    WScript.Echo "ERROR: Cannot open image: " & imagePath
    WScript.Quit 8
  End If
  ' 收件人為內網 Notes，用 BINARY 即可；若寄外部 SMTP 顯示有問題改 ENC_BASE64
  Call childImg.SetContentFromBytes(stream, "image/png", ENC_IDENTITY_BINARY)
  Call stream.Close()
Else
  ' --- 純 text/html (與舊版相容) ---
  Set stream = notesSession.CreateStream()
  Call stream.WriteText(htmlBody)
  Call body.SetContentFromText(stream, "text/html; charset=UTF-8", ENC_IDENTITY_BINARY)
  Call stream.Close()
End If

' 還原設定
notesSession.ConvertMIME = True

If Err.Number <> 0 Then
  WScript.Echo "ERROR: Build MIME failed. Err=" & Err.Number & " " & Err.Description
  WScript.Quit 6
End If

Call doc.Send(False)
If Err.Number <> 0 Then
  WScript.Echo "ERROR: Send failed. Err=" & Err.Number & " " & Err.Description
  WScript.Quit 7
End If

WScript.Echo "OK"
WScript.Quit 0
