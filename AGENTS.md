# AGENTS.md

給 AI coding agent 在這個 repo 工作時的指南。使用者面向的工具說明請看 [`README.md`](README.md)。

---

## 這個 repo 是什麼

一組操作 **HCL / IBM Lotus Notes（Domino）** 的 **VBScript** 小工具，透過 `Notes.NotesSession` COM 介面做批次任務（下載附件、抽資料、匯出設計、寄信、產報表等）。另可能包含瀏覽器端 **HTML** 小工具，配合匯出 CSV 在本機比對統計。所有 `.vbs` 都以 `cscript.exe`（命令列 WSH）在 **Windows** 上執行。

每支腳本都是**自含式**——沒有共用 library，helpers 直接複製到每個檔案。這是刻意的設計，方便單檔獨立部署 / 排程 / 給人拷貝。

---

## 執行環境（重要）

- **只能在 Windows 跑**（需要 Notes/Domino client + Windows COM）。你這邊在 Linux/雲端 **沒辦法實際執行**，只能寫、commit、push；使用者在自己 Windows 機器上跑。
- 使用者必須**先開啟並登入 Notes client**（腳本沿用登入身分）。
- Excel COM (`Excel.Application`) 需要本機安裝 **Excel**（某些工具會用）。
- 命令列 always 用 **`cscript`**（不用 `wscript`——`wscript` 是 GUI 模式，`WScript.Echo` 會跳對話框、`WScript.StdOut` 無法寫）。

---

## 檔案 / 命名慣例

- 主腳本：`PascalCase.vbs`（如 `DownloadNotesAttachments.vbs`）
- 排程 / 雙擊包裝：`snake_case.bat`（如 `download_attachments.bat`）
- HTML 工具：`snake_case.html`
- README 每加一支新工具，就在「工具一覽」加一節、並在「結束代碼」補上其 exit code 表
- 所有腳本開頭都要有 `Option Explicit`

### 儲存編碼（常見雷）

- Repo 內 `.vbs` 檔案存為 **UTF-8**（git 存的），這樣網頁 / GitHub 顯示 OK。
- 但 **cscript 對 UTF-8 without BOM 的支援不可靠**——會當 ANSI 讀，Big5 系統上中文 syntax error / 亂碼。
- 使用者在本機編輯後儲存，**建議存成 UTF-16 LE with BOM**（Windows 記事本「Unicode」選項；VS Code 右下角 `UTF-16 LE`）。
- README 有註記：檔案編碼是 UTF-8。如果使用者反應「跑不動 / 中文亂碼」，第一個要確認的就是編碼。

---

## 每支腳本的骨架

固定按這個結構寫，方便使用者對照 README：

```vbs
Option Explicit

' <ScriptName>.vbs
' 用途：<一句話說明>
'
' ------------------------------ 使用說明 ------------------------------
' 前置需求：先開啟並登入 Notes client；...
' 執行方式：cscript //nologo <ScriptName>.vbs [args...]
'
' ★ 執行前「必須」修改的設定（都在下方「設定區」）：
'     SERVER  → ...
'     DBFILE  → ...
' ---------------------------------------------------------------------

' ==================== 設定區（請依你的環境修改） ====================
Const SERVER  = "YourServer/YourOrg"
Const DBFILE  = "doclib\library.nsf"
' ...
' ====================================================================

' --- 主流程 ---
Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")

Dim session
On Error Resume Next
Set session = CreateObject("Notes.NotesSession")
If Err.Number <> 0 Then
  WScript.Echo "ERROR: 無法建立 Notes.NotesSession (Notes 未開/未登入?) " & Err.Description
  WScript.Quit 4
End If
On Error Goto 0

' ...
WScript.Echo "OK ..."
WScript.Quit 0

' ---------- helpers ----------
' 每支腳本自帶所需 helper，不共用檔案
```

### 標準 exit code

儘量沿用 README 的既有意義（雖然不同腳本可能有小差異）：

| Code | 意義 |
|---|---|
| 0 | 成功 |
| 2 | 命令列參數不足（有的腳本） |
| 3 | 找不到 / 無法開啟資料庫 |
| 4 | 無法建立 `Notes.NotesSession` |
| 5 | 找不到指定檢視 |
| 6 | 找不到某輸入檔（Excel、關鍵字檔等） |
| 7 | 輸入檔讀取失敗（Excel、FTSearch 等） |
| 8 | 依編號找不到文件 / 寫入輸出檔失敗 |
| 9 | DXL 匯出 / 解析失敗（多半權限或 Notes 版本問題） |
| 10 | 無法建立 `Excel.Application` |

新增新 code 時要同步更新 README 的 exit code 段。

---

## 常見 helpers（每支腳本自帶）

這些 helper 幾乎每支腳本都會用到，直接複製即可，別發明新版本：

```vbs
' 相對路徑轉成以本 .vbs 所在資料夾為基準的絕對路徑
Function AbsPath(p)
  If p = "" Then AbsPath = "" : Exit Function
  If InStr(p, ":\") > 0 Or Left(p, 2) = "\\" Then
    AbsPath = p
  Else
    AbsPath = fso.BuildPath(fso.GetParentFolderName(WScript.ScriptFullName), p)
  End If
End Function

' 建立資料夾（含上層）
Sub EnsureFolder(folder)
  If folder = "" Then Exit Sub
  If Not fso.FolderExists(folder) Then
    EnsureFolder fso.GetParentFolderName(folder)
    fso.CreateFolder folder
  End If
End Sub

' 寫 UTF-8 + BOM 文字檔（Excel 開 CSV 中文不亂碼）
Sub WriteTextUtf8(path, text)
  EnsureFolder fso.GetParentFolderName(path)
  Dim st : Set st = CreateObject("ADODB.Stream")
  st.Type = 2 : st.Charset = "utf-8" : st.Open
  st.WriteText text
  st.SaveToFile path, 2   ' 2 = 覆寫
  st.Close
End Sub

' CSV 一列：陣列轉字串，含逗號/雙引號/換行的欄位自動加雙引號
Function CsvRow(arr)
  Dim out, k, cell
  out = ""
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

' 讀 Notes 欄位文字（避開 GetItemValue 的物件/Null 型別問題，見下方 Gotcha）
Function ItemText(d, fld)
  ItemText = ""
  If Not d.HasItem(fld) Then Exit Function
  Dim it : Set it = d.GetFirstItem(fld)
  If it Is Nothing Then Exit Function
  ItemText = it.Text
End Function
```

其他常見進度顯示：

```vbs
' 每 N 筆刷新同一行（用 vbCr 蓋回行首）
If PROGRESS_EVERY > 0 Then
  If (scanned Mod PROGRESS_EVERY) = 0 Then _
    WScript.StdOut.Write vbCr & "  處理中… 已看過 " & scanned & " 筆"
End If
```

---

## Notes COM 陷阱（血淚累積）

### 1. `GetItemValue` 會回物件，`CStr` / 當參數傳都會炸

**症狀**：`類型不符: 'v'`（型別不符錯誤，通常指向存 `v(0)` 的那行）。

原因：Notes 的日期欄位、Names 欄位等，`GetItemValue` 回來可能是 `NotesDateTime` 這類**物件**。VBScript 對物件做 `CStr(v(0))` 或把 `v(0)` 當**函式參數**傳（觸發預設屬性求值）都會爆。

**修法**：讀欄位一律走 `NotesItem.Text`（Notes 直接給字串，不會有物件），如上方 `ItemText`。

要讀數字或日期時，也先拿字串再自己 parse：

```vbs
Function ItemNumber(d, fld)
  ItemNumber = 0
  Dim s : s = ItemText(d, fld)
  If Trim(s) = "" Then Exit Function
  s = Replace(s, ",", "")   ' 去千分位逗號
  On Error Resume Next
  Dim x : x = CDbl(s)
  If Err.Number = 0 Then ItemNumber = x
  Err.Clear : On Error Goto 0
End Function

Function ItemDate(d, fld)
  ItemDate = Empty
  Dim s : s = ItemText(d, fld)
  If Trim(s) = "" Then Exit Function
  On Error Resume Next
  Dim x : x = CDate(s)
  If Err.Number = 0 Then ItemDate = x
  Err.Clear : On Error Goto 0
End Function
```

`CDate` 依系統 locale 解析——如果 Notes 給 `MM/DD/YYYY` 但系統 locale 是 zh-TW（預設 `yyyy/MM/dd`），`CDate` 會 parse 失敗。這種情況要自己拆字串 + `DateSerial(y, m, d)`（locale-safe）。

### 2. `NotesViewEntry.ColumnValues` 每格都可能踩雷

分類檢視、某些奇怪 entry，`ColumnValues` 可能回不預期的東西。取值時用 `VarType` **inline** 判斷、別包 function 呼叫傳進去。整個 per-entry 處理最好抽成 Sub、外面 `On Error Resume Next` 包起來，怪 entry 直接算 `skipped`、繼續下一筆。

### 3. Notes item 名字不一定等於欄位標題

畫面上看到的中文欄位標題**不是** item name。要用 Notes 內文件「文件內容（藍色 i）→ Fields 分頁」查真正的 item name，或用 `ReadNotesDoc.vbs` 匯一筆已知文件的所有欄位。常見坑：`Station Time11`（看似有空白）實際是 `StationTime11`（沒空白）。

不確定就在腳本裡加短暫的 debug dump（列出前一筆保留文件所有名字含 keyword 的 items），確認完再拿掉。

### 4. `HasItem`：欄位不存在時的安全防呆

任何 `GetItemValue` / `GetFirstItem` 之前，先 `If d.HasItem(fld) Then`——欄位不存在直接回預設值，別靠例外處理。

### 5. 分類檢視 iterating 會有重複

分類檢視同一份文件可能在多個分類下出現，`GetFirstDocument`/`GetNextDocument` 可能重複走訪。解法：用 `Scripting.Dictionary` 以 `doc.NoteID` 去重。

### 6. 讀設計元素（Form / View design）通常要 Designer ACL

`CreateDXLExporter` 對 form 做 `Export` 時，**Reader 權限只會拿到 `<database>` 外殼**，內部 `<form>` / `<field>` / `<button>` 是空的（表面上看 43 個 note 匯出，但實際 XML 是空的）。這不是 XPath 或 namespace 問題，是純**權限問題**。判斷方法：用 `OUTFORMAT="dxl"` 存 XML 看有沒有 `<field` / `<button`。

### 7. 前端 UI 類別在外部 COM 拿不到

`Notes.NotesSession` 只給後端類別（`NotesDatabase`、`NotesDocument`…）。**`NotesUIWorkspace` / `NotesUIDocument` 這些前端類別**——負責「真的點按鈕、拉下拉選單、模擬打字」——**只能在 Notes client 內部的 LotusScript 用**，外部 VBScript 拿不到。要做「自動填單送出」只能走後端（`CreateDocument` → 設欄位 → `ComputeWithForm` → `Save` / `Send`），複製送出按鈕的邏輯，不是模擬點擊。

---

## Excel COM 陷阱

### 建立含多張工作表的 workbook

`Workbooks.Add.Worksheets.Add(Empty, wsF)` 這種**位置參數傳 `Empty`** 的寫法，某些 Excel 版本會爆 `無法取得類別 Sheets 的 Add 屬性`。**改用 `SheetsInNewWorkbook`**：

```vbs
Dim origSiNW : origSiNW = xl.SheetsInNewWorkbook
xl.SheetsInNewWorkbook = 2                   ' 一次建 2 張
Dim wb : Set wb = xl.Workbooks.Add()
xl.SheetsInNewWorkbook = origSiNW            ' 還原全域設定

' 防呆：把工作表數修成剛好 2
Do While wb.Worksheets.Count > 2
  wb.Worksheets(wb.Worksheets.Count).Delete
Loop
Do While wb.Worksheets.Count < 2
  wb.Worksheets.Add
Loop

Dim wsF : Set wsF = wb.Worksheets(1) : wsF.Name = "Sheet A"
Dim wsB : Set wsB = wb.Worksheets(2) : wsB.Name = "Sheet B"
```

### 寫入前把儲存格設文字格式，避免 `=` 開頭被當公式

寫程式碼、Notes 公式、以 `=` 開頭的字串到 Excel，若不先設文字格式，Excel 會嘗試當公式解析：

```vbs
ws.Range(ws.Cells(1,1), ws.Cells(nRows, nCols)).NumberFormat = "@"
' 然後再指派 2D 陣列
ws.Range(ws.Cells(1,1), ws.Cells(nRows, nCols)).Value = data
```

### 存 .xlsx 用 FileFormat = 51

```vbs
wb.SaveAs path, 51   ' 51 = xlOpenXMLWorkbook (.xlsx)
```

### 一次寫入 2D 陣列比一格一格快很多

用 `ReDim data(nRows-1, nCols-1)`（0-based）填好，再 `Range.Value = data` 一次寫入。

---

## DXL / MSXML 陷阱

- DXL 用預設 namespace `http://www.lotus.com/dxl`。XPath 前必須註冊 prefix：

  ```vbs
  Const DXL_NS = "xmlns:d='http://www.lotus.com/dxl'"
  dom.setProperty "SelectionLanguage", "XPath"
  dom.setProperty "SelectionNamespaces", DXL_NS
  ' 然後 XPath 用 d: 前綴
  Set nodes = dom.selectNodes("//d:form | //d:subform")
  ```

- 若擔心不同版本 DXL namespace 不一致，可退回 `//*[local-name()='form']` 這種 namespace 無關的寫法。

- **先移除 `<?xml ... ?>` 宣告再 `loadXML`**——MSXML 對 Unicode 字串加編碼宣告會報錯：

  ```vbs
  Function StripXmlDecl(x)
    Dim t : t = LTrim(x)
    If Left(t, 5) = "<?xml" Then
      Dim p : p = InStr(t, "?>")
      If p > 0 Then t = Mid(t, p + 2)
    End If
    StripXmlDecl = t
  End Function
  ```

- DXL 的 forms 需要 **Designer** ACL 才會匯出 body；Reader 只拿到 `<database>` 外殼（見上方「Notes COM 陷阱 #6」）。

---

## 用 Notes 寄 HTML 信（MIME）

固定 pattern，照抄 `SendNotesMail.vbs`：

```vbs
Const ENC_IDENTITY_BINARY = 1729     ' MIME encoding 常數

Dim mdb : Set mdb = ns.GetDatabase("", "")
Call mdb.OpenMail()
ns.ConvertMIME = False               ' 讓我們自己控制 MIME

Dim memo : Set memo = mdb.CreateDocument()
Call memo.ReplaceItemValue("Form", "Memo")
Call memo.ReplaceItemValue("SendTo", Split(mailTo, ";"))   ' 多人用 ; 分隔
Call memo.ReplaceItemValue("Subject", subject)

Dim body : Set body = memo.CreateMIMEEntity("Body")
Dim stream : Set stream = ns.CreateStream()
Call stream.WriteText(htmlBody)
Call body.SetContentFromText(stream, "text/html; charset=UTF-8", ENC_IDENTITY_BINARY)
Call stream.Close()

ns.ConvertMIME = True                ' 還原
Call memo.Send(False)
```

- 內嵌圖片：用 `multipart/related` + Content-ID，HTML 端 `<img src="cid:...">`。範例見 `SendNotesMail.vbs`。
- Notes 文件連結：`notes://<server>/<replicaID>/0/<UNID>`——點擊在 Notes 開回原文件（近似「Copy as Document Link」的效果，但不是真的 CD-record doclink）。
- Lotus Notes 內文的 HTML 引擎陽春：**沒有 JavaScript**、`:checked` 這類互動 CSS 也不支援。分頁 tab、動畫都做不了。要真互動就用附件 HTML 讓使用者用瀏覽器打開。

---

## VBScript 語法陷阱

- **`Const` 只能給單一字面值**——不能字串串接、不能陣列、不能跨行組合。多字串接續要用 `Dim`：
  ```vbs
  Dim MAIL_TO : MAIL_TO = "a@x.com;" & _
                          "b@x.com;" & _
                          "c@x.com"
  ```

- **沒有內建 `IIf`**——要自己寫：
  ```vbs
  Function IIf(cond, a, b)
    If cond Then IIf = a Else IIf = b
  End Function
  ```

- **`System.Collections.ArrayList` 用 `For Each` 迭代最安全**——用 `arr(i)` 索引存取有時因為預設屬性解析問題失敗，優先 `For Each` 或 `arr.Item(i)`。

- **`On Error Resume Next` 只在必要範圍打開**：碰到會炸的 COM 呼叫（`entry.ColumnValues`、`Excel.Application.CreateObject`、外部 API）就打開；讀完 Err.Number、`Err.Clear` 後立刻 `On Error Goto 0` 還原。範圍越窄越好，別包整個 loop。

- **單行 `If ... Then ... Else` 可以，但要注意 `_` 續行**：
  ```vbs
  If PROGRESS_EVERY > 0 Then _
    WScript.StdOut.Write vbCr & "  進度..."
  ```

- **`Split` 空字串會回一個含空元素的陣列**（`UBound=0`），別忘了 `Trim` 判空。

- **Wildcard / LIKE 比對**：VBScript 沒 SQL LIKE，用 `VBScript.RegExp` 轉：`*`→`.*`、`?`→`.`，其他 regex 特殊字元跳脫。

---

## Git 慣例

- Commit subject 一行動詞開頭、簡短表達 what；body 說明 why（尤其修 bug 時，把觸發的行為 / 錯誤訊息交代清楚）。
- 每次 push 前先 `git status` 確認沒把 log 檔、意外檔案吃進去。
- 使用者環境沒有 `gh` CLI；建 PR 用 GitHub MCP tools（`mcp__github__create_pull_request` 等）——本地 shell 開 PR 不會通。
- 使用者本機常用相對路徑 `.vbs` + `.bat` 雙擊執行，log 寫在同資料夾 `xxx_log.txt`——別把絕對路徑寫死進 commit（除非使用者明確要）。

---

## 什麼「不要做」

- **不要**用 `GetItemValue(fld)` + `CStr(v(0))` 這種寫法讀 Notes 欄位（見上方 Gotcha #1）。改用 `NotesItem.Text` 為基礎的 `ItemText`。
- **不要**嘗試從外部腳本模擬 Notes 前端 UI 操作（`NotesUIWorkspace` 之類拿不到）。要「填單送出」就走後端 document 寫入 + `Send`。
- **不要**猜 Notes 的 item name——請使用者用文件內容（藍色 i）→ Fields 分頁確認，或加暫時性 debug dump。
- **不要**寫 CSV 給 Excel 開卻沒加 BOM——中文會亂碼。UTF-8 + BOM 是預設。
- **不要**把 `.vbs` 存成 ANSI 或無 BOM UTF-8 給使用者跑——中文會爛。存 UTF-16 LE with BOM（「Unicode」）。
- **不要**在 CreateObject Excel 建立 workbook 時用 `Worksheets.Add(Empty, wsF)`——用 `SheetsInNewWorkbook`（見 Excel COM 陷阱）。
- **不要**只跑一次 `WScript.Echo "OK ..."` 就交差——腳本結尾要能讓使用者判斷「掃了 N 筆、命中 M 筆、輸出到哪」。
- **不要**在腳本裡把使用者的個人資訊（email、真實檔名路徑）寫死然後 commit 到 main——放到 branch/README 佔位字串即可，或使用者自己填在他機器上。

---

## 迭代工作模式

使用者通常會分步驟給需求、貼 error 訊息或截圖回報問題。回覆節奏：

1. **精準修改**：出錯時只改需要改的那幾行，別大改。
2. **每次改完 commit + push**：使用者拉最新版跑，貼結果回來繼續。
3. **不確定就在腳本裡加暫時性 debug dump**（列 item 名字、印 raw value），拿到答案後移除。
4. **語系**：使用者用繁體中文；回覆用繁體中文，code 內識別字（欄位名、變數）保持英文。
