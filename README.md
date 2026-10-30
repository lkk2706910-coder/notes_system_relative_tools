# notes_system_relative_tools — Cost Down Project 預估報表

`CostDownReport.vbs`：從 CDPAS（`UMCD46/UMC` `fab\cdpas.nsf`）檢視「6. 依分類 (Last 2 Years)」撈出 cost-down 項目，篩選後依預估月份整理成 HTML 表格，**同時**透過 Notes 寄信給收件人並在本機存一份 HTML。唯讀，不修改任何文件。

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

輸出欄：Project ID（可點回 Notes 文件）/ Section / CostPool / eqptype / 描述 / Method / 預估日 / 月效益 / Total Benefit / Approve Date（`StationTime11` 內最早的 `MM/DD/YYYY`）。每個月份一個區塊，月內先按 SEC 再按預估日排序。

## 設定（都在 `CostDownReport.vbs` 頂端「設定區」）

| 設定 | 預設 | 說明 |
|---|---|---|
| `SERVER` / `DBFILE` / `VIEWNAME` | `UMCD46/UMC` / `fab\cdpas.nsf` / `6. 依分類 (Last 2 Years)` | 來源 |
| `F_*` | 見檔案 | 各欄位 item name |
| `STATUS_LIKE` | `New Project%` | Status 篩選樣式 |
| `MAIL_TO` | （分號分隔多人） | 收件人，用 `Dim` + `&` 續行，不能用 `Const` |
| `MAIL_SUBJECT` | `Cost Down Project 預估報表（未實際結案）` | 信件主旨 |
| `SEND_MAIL` | `True` | 是否寄信 |
| `SAVE_HTML` | `True` | 是否把同一份 HTML 存到 `OUTFILE`（先存檔再寄信） |
| `OUTFILE` | `cost_down_report.html` | HTML 輸出路徑，相對路徑以 .vbs 所在資料夾為基準 |
| `PROGRESS_EVERY` | `200` | 每 N 筆刷一次進度 |

`SEND_MAIL` 與 `SAVE_HTML` 可同時開（預設）；只想測試版面就把 `SEND_MAIL` 改 `False`。

## 結束代碼

| Code | 意義 |
|---|---|
| 0 | 成功（包含沒有符合資料的情況） |
| 3 | 找不到 / 無法開啟資料庫 |
| 4 | 無法建立 `Notes.NotesSession` |
| 5 | 找不到檢視，或無法開啟郵件資料庫寄信 |

## 注意

- 檔案在 repo 內是 UTF-8；本機編輯後請存 **UTF-16 LE with BOM**（記事本「Unicode」），否則 cscript 讀中文會錯。
- 收件人目前直接寫在腳本裡，要給別人用前請改成佔位字串。
