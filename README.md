# notes_system_relative_tools — Cost Down Project 預估報表

`CostDownReport.vbs`：從 CDPAS（`UMCD46/UMC` `fab\cdpas.nsf`）檢視「6. 依分類 (Last 2 Years)」撈出 cost-down 項目，篩選後依預估月份整理成 HTML 表格，透過 Notes 寄信給收件人，並產生一份**互動版 HTML**（瀏覽器開啟，可依欄位排序 / 篩選）當附件與本機存檔。Project ID 是 Notes 文件連結（＝文件內容 → Meta → Identifier 的 `notes://` 網址），點了用 Notes 開回原文件。唯讀，不修改任何文件。

```
CDPAS 檢視 ──CostDownReport.vbs──▶ Notes 信件：靜態分月表格（Project ID 可點）
                                 ├─ 附件 cost_down_report.html：互動版（排序 / 篩選 / 合計）
                                 └─ 本機 cost_down_report.html：同一份互動版
```

AI agent 工作指南見 [`AGENTS.md`](AGENTS.md)；其他通用 Notes 工具在 `main` branch。

## 用法

前提：Notes client 已開啟並登入，對 `fab\cdpas.nsf` 有 Reader 權限。

```bat
cscript //nologo CostDownReport.vbs
```
或雙擊 / 排程 `cost_down_report.bat`（log 寫在同資料夾 `cost_down_report_log.txt`）。

## 篩選與欄位

| 條件 | 說明 |
|---|---|
| `Forecast_Rls_Date` 有值 | 預估完成日，同時用來分月 |
| `Actual_Date_Final` 空白 | 尚未實際結案 |
| `Status` 符合 `STATUS_LIKE` | 預設 `New Project%`，`%` 為萬用字元，留空＝不套用 |
| `Project_ID` 非空 | 空的略過 |

輸出欄：Project ID（連結）/ SEC / Cost Pool / EQP Type / Project Desc. / Method / Benefit Occur (Forecast) / Monthly Benefit / Total Benefit (Forecast) / Approve Date（`StationTime11` 內最早的 `MM/DD/YYYY`）。

### 信件內文（靜態）

每個預估月份一個區塊，月內先按 SEC 再按預估日排序。Notes 信件內文不能執行 JavaScript，所以內文沒有排序 / 篩選，頂端會提示「請開附件」。

### 互動版 HTML（附件 / 存檔）

全部資料放在一張平面表格，多一欄「預估月份」：

- **排序**：點欄位標題，再點一次反向；數字欄依數值、日期欄依日期、空值排最後。
- **篩選**：標題下方第二列逐欄篩選。預估月份 / SEC / Cost Pool / EQP Type / Method 是下拉選單；其他欄是文字框（包含即可，不分大小寫）；數字欄支援 `>1000`、`<=500`、`=0`、`500-2000`。左上角另有全文搜尋。
- **合計**：底列顯示目前篩選結果的筆數與 Monthly / Total 合計，隨篩選即時更新。
- **清除篩選 / 排序** 按鈕回到初始狀態。

附件以一般附件型式附上（雙擊用瀏覽器開）。

## 設定（都在 `CostDownReport.vbs` 頂端「設定區」）

| 設定 | 預設 | 說明 |
|---|---|---|
| `SERVER` / `DBFILE` / `VIEWNAME` | `UMCD46/UMC` / `fab\cdpas.nsf` / `6. 依分類 (Last 2 Years)` | 來源 |
| `F_*` | 見檔案 | 各欄位 item name |
| `STATUS_LIKE` | `New Project%` | Status 篩選樣式 |
| `MAIL_TO` | （分號分隔多人） | 收件人，用 `Dim` + `&` 續行，不能用 `Const` |
| `MAIL_SUBJECT` | `Cost Down Project 預估報表（未實際結案）` | 信件主旨 |
| `SEND_MAIL` | `True` | 是否寄信 |
| `ATTACH_HTML` | `True` | 寄信時是否附上互動版 HTML |
| `SAVE_HTML` | `True` | 是否把互動版 HTML 存到 `OUTFILE`（先存檔再寄信） |
| `OUTFILE` | `cost_down_report.html` | HTML 輸出路徑，相對路徑以 .vbs 所在資料夾為基準；附件檔名也用它 |
| `PROGRESS_EVERY` | `200` | 每 N 筆刷一次進度 |

三個開關可同時開（預設）。只想測版面不寄信就把 `SEND_MAIL` 改 `False`；`SAVE_HTML = False` 而 `ATTACH_HTML = True` 時會先寫到暫存資料夾、寄完刪掉。

## 結束代碼

| Code | 意義 |
|---|---|
| 0 | 成功（包含沒有符合資料的情況） |
| 3 | 找不到 / 無法開啟資料庫 |
| 4 | 無法建立 `Notes.NotesSession` |
| 5 | 找不到檢視，或無法開啟郵件資料庫 / 寄信失敗 |
| 8 | 無法開啟要附加的 HTML 檔 |

## 注意

- 檔案在 repo 內是 UTF-8；本機編輯後請存 **UTF-16 LE with BOM**（記事本「Unicode」），否則 cscript 讀中文會錯。
- 收件人目前直接寫在腳本裡，要給別人用前請改成佔位字串。
- 互動版的 JavaScript 放在 `InteractiveJs()`，字串只用單引號；要改功能先在瀏覽器測過，再逐行回填。
