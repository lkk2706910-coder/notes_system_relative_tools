# notes_system_relative_tools — PMS 文件比對（Excel key → Notes 資料 + 分類）

讀一份 Excel，用 `name` 欄（如 `U20251203756`）逐筆到 Notes `eri\pms.nsf` 以 `DOCNO` 查對應文件，抓 `A17 / A21 / A25 / A26`，依 `A17` 的「-」前綴分類（`CUSILPE-B12` → `CUSILPE`），並附上每筆文件的 Notes 連結（＝文件內容 → Meta → Identifier 的 `notes://…`），輸出 xlsx。**唯讀，不修改任何文件。**

```
輸入.xlsx（name 欄）──MatchPmsDocs.exe──▶ 輸入_notes.xlsx
                                          ├─ 明細：原欄位 + A17/A21/A25/A26 + 機台類別 + Link（可點）
                                          ├─ 分類統計：每類筆數、amt 合計、總計
                                          └─ 每個類別一張工作表
```

AI agent 工作指南見 [`AGENTS.md`](AGENTS.md)；其他通用 Notes 工具在 `main` branch。

---

## 用法

```bat
MatchPmsDocs.exe 檔案.xlsx
MatchPmsDocs.exe 檔案.xlsx --preview 5      # 順便在畫面印前 5 列結果
```
也可以直接把 xlsx **拖到 exe 上**。前提：Notes client 已開啟並登入。輸出預設為 `<輸入檔名>_notes.xlsx` 放同資料夾，log 在 exe 同資料夾 `match_pms_docs_log.txt`。

查詢方式：先用 `(ByDocNo)` 檢視 `GetDocumentByKey` 精準查（快）；查不到或檢視不存在時退回公式 `DOCNO = "…"` 搜尋。Excel 裡 key 空白或 Notes 查無的列，`機台類別` 標 `(查無)`、其餘新欄留空，log 會列出查無的 key。

### 輸出欄位

| 欄 | 內容 |
|---|---|
| 原表格所有欄位 | 原樣保留（含 `c_costctr, name, description, pum, ttl_qty, amt, 工號, 員工姓名`） |
| `A17` `A21` `A25` `A26` | Notes 文件欄位原值（`fetch_fields` 可改） |
| `機台類別` | `A17` 取 `-` 前的字串；沒有 `-` 就整串；查無＝`(查無)` |
| `Link` | 文件的 `notes://` URL，Excel 裡可直接點開文件 |

`分類統計` 工作表依筆數排序，`amt 合計` 用 `amt_column` 指定的欄加總（找不到該欄就只算筆數）。

### 設定：`match_pms_docs.ini`（放 exe 同資料夾，從 `.ini.example` 複製；沒有就用預設）

| key | 預設 | 說明 |
|---|---|---|
| `server` / `dbfile` | `UMCD20/UMC` / `eri\pms.nsf` | Notes 資料庫 |
| `docno_view` | `(ByDocNo)` | 第一欄為 DOCNO 的檢視 |
| `docno_field` | `DOCNO` | 與 Excel key 比對的 Notes 欄位 |
| `fetch_fields` | `A17,A21,A25,A26` | 要抓的欄位 |
| `class_field` | `A17` | 用哪個欄位做分類 |
| `input_xlsx` | （命令列給） | 固定輸入檔路徑 |
| `sheet` / `header_row` | 第一個工作表 / `1` | 工作表與標題列 |
| `key_column` | `name` | Excel 的 key 欄名（不分大小寫） |
| `amt_column` | `amt` | 分類統計加總欄 |
| `output_xlsx` | `<輸入>_notes.xlsx` | 輸出路徑 |
| `sheet_per_category` | `true` | 是否每類一張工作表 |

---

## 打包成 exe

```bat
build_exe.bat
```
或手動（打包用的 Python 位元數要能連 Notes，通常跟 Notes 同位元；多個 Python 時用 `py -3-32`）：
```bat
python -m pip install --upgrade pywin32 openpyxl pyinstaller
python -m PyInstaller --onefile --console --clean --name MatchPmsDocs --hidden-import pythoncom --hidden-import pywintypes --hidden-import win32com --hidden-import win32com.client --hidden-import win32com.client.dynamic --hidden-import win32com.client.gencache --hidden-import win32timezone --collect-submodules win32com --collect-submodules openpyxl match_pms_docs.py
```

## 結束代碼

| Code | 意義 |
|---|---|
| 0 | 成功 |
| 3 | 找不到 / 無法開啟資料庫 |
| 4 | 無法建立 `Notes.NotesSession`（Notes 未開 / 未登入 / pywin32 未裝 / 位元數不符） |
| 6 | 找不到輸入 Excel |
| 7 | Excel 讀取失敗，或標題列找不到 `key_column` |
| 8 | 寫入輸出 xlsx 失敗（檔案開著？） |

## 常見問題

- **全部 `(查無)`**：先確認 `(ByDocNo)` 檢視第一欄真的是 DOCNO、Excel 的 `name` 值與 Notes `DOCNO` 完全一致（大小寫/空白）。可暫時把 `docno_view` 留空強制走公式搜尋驗證。
- **Link 點了沒反應**：該台電腦要有 Notes client 並登入；`notes://` 協定由 Notes 處理。
- **ini 中文亂碼 / 讀不到**：程式會自動辨識 UTF-8 / Big5；仍有問題就把 ini 另存 UTF-8。
