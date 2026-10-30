# notes_system_relative_tools — PMS 文件比對（Excel key → Notes 資料 + 分類）

讀一份 Excel，用 `name` 欄逐筆到 Notes `eri\pms.nsf` 查對應文件——**搜尋依據可選 `DOCNO`（單號）、`A21`（material）、`A22`（原廠料號）**——抓 `A17 / A21 / A22 / A25 / A26 / VendorNameList`，依 `A17` 的「-」前綴分類（`CUSILPE-B12` → `CUSILPE`），附上文件的 Notes 連結（＝文件內容 → Meta → Identifier 的 `notes://…`）；可同時或單獨再到 `UMCD14/UMC` `eop\gr.nsf`（驗收系統）用同一個 key 對 `DCNUM` 找文件、把連結放在 `GR Link`，輸出 xlsx。**唯讀，不修改任何文件。**

```
輸入.xlsx（name 欄）──MatchPmsDocs.exe──▶ 輸入_notes.xlsx
                                          ├─ 明細：原欄位 + A17/A21/A22/A25/A26/VendorNameList + 機台類別 + ePMS Link + GR Link（可點）
                                          ├─ 分類統計：每類筆數、amt 合計、總計
                                          └─ 每個類別一張工作表
```

AI agent 工作指南見 [`AGENTS.md`](AGENTS.md)；其他通用 Notes 工具在 `main` branch。

---

## 用法

### 視窗模式（預設）

雙擊 `MatchPmsDocs.exe`（或把 xlsx 拖到 exe 上、或 `MatchPmsDocs.exe 檔案.xlsx`）會開一個視窗：

1. **輸入來源**二選一：
   - **Excel 檔** → 上方瀏覽選檔（拖檔進來的會自動帶入並開始查）。
   - **直接貼上清單** → 文字框貼 key（一行一個，或逗號 / 分號 / 空白分隔；從 Excel 複製整欄直接貼也行，會自動去空白、去重、忽略標題 `name`），下方即時顯示「已貼上 N 個 key」，**多筆一次查**。匯出時原欄位只有 `name` 一欄。
2. **第二列選項**：
   - **Excel key 欄對應**：`DOCNO（單號）` / `A21（material）` / `A22（原廠料號）`——Excel 的 `name` 欄就拿去對 ePMS 的這個欄位。
   - **比對 ePMS** / **比對驗收系統**：兩個勾選框，可只勾一個或都勾（至少一個）。
   - 不論用哪個搜尋，`A21 / A22 / VendorNameList` 一定帶出；用 A21/A22 搜尋時另外補 `DOCNO` 欄方便對照。
   - 用 A21/A22 搜尋同一料號對到多筆文件時，**每筆各一列**（Excel 原欄位重複），log 會提示有幾個 key 對到多筆；ini `multi_match = first` 可改成只取第一筆。
3. **查詢 Notes** → 背景逐筆查，進度條與狀態列顯示「n/總數，命中 m」；查完表格列出所有欄位。
4. **逐列確認**：
   - `✓` 欄點一下 → 切換該列**要不要匯出**（排除的列變灰）
   - 雙擊 **機台類別** 儲存格 → 直接改分類（Enter 確定、Esc 取消），右側統計即時更新
   - 雙擊 **Link** → 用 Notes 開啟該文件核對
   - **檢查欄位…** → 輸入一個 key，列出該文件所有 item 的名稱/型別/值（空的灰字），用來診斷「畫面有值、輸出沒有」
   - 「只看類別」下拉 → 只顯示某一類；查無的列紅字
   - 右側 **分類統計** 只計算勾選中的列
5. **匯出 Excel…** → 選存檔位置（預設 `<輸入>_notes.xlsx`），只匯出勾選的列；完成可直接開啟。
6. 最下方是 log，同時也寫到 exe 同資料夾 `match_pms_docs_log.txt`。

### 命令列模式（排程 / 不需確認）

```bat
MatchPmsDocs.exe 檔案.xlsx --cli
MatchPmsDocs.exe 檔案.xlsx --cli --preview 5          # 順便印前 5 列
MatchPmsDocs.exe 檔案.xlsx --cli --key A21            # 用 material 搜尋（DOCNO / A21 / A22）
MatchPmsDocs.exe 檔案.xlsx --cli --no-gr              # 只比對 ePMS
MatchPmsDocs.exe 檔案.xlsx --cli --no-epms            # 只比對驗收系統
MatchPmsDocs.exe --cli --keys "U20251203756,V20260202498"   # 不用 Excel，直接給清單（也可給一個文字檔路徑，一行一個）
MatchPmsDocs.exe --dump U20251203756                  # 列出該單所有 item（名稱/型別/值），診斷欄位帶不出來
MatchPmsDocs.exe --dump S00213358 --key A21           # 依 material 找到的那張單
```
不開視窗，直接查詢並匯出到 `<輸入>_notes.xlsx`。exe 是 `--noconsole` 打包，此模式的訊息看 log 檔。

前提：Notes client 已開啟並登入。

查詢方式：先用 `(ByDocNo)` 檢視 `GetDocumentByKey` 精準查（快）；查不到或檢視不存在時退回公式 `DOCNO = "…"` 搜尋。Excel 裡 key 空白或 Notes 查無的列，`機台類別` 標 `(查無)`、其餘新欄留空，log 會列出查無的 key。

### 輸出欄位

| 欄 | 內容 |
|---|---|
| 原表格所有欄位 | 原樣保留（含 `c_costctr, name, description, pum, ttl_qty, amt, 工號, 員工姓名`） |
| `A17` `A21` `A22` `A25` `A26` `VendorNameList` | ePMS 文件欄位原值（`fetch_fields` 可改；`A21 / A22 / VendorNameList` 固定帶出）。用 A21/A22 搜尋時另有 `DOCNO` 欄 |
| `機台類別` | `A17` 取 `-` 前的字串；沒有 `-` 就整串；查無＝`(查無)` |
| `ePMS Link` | `eri\pms.nsf` 文件的 `notes://` URL，Excel 裡可直接點開 |
| `GR Link` | 再到 `UMCD14/UMC` 的 `eop\gr.nsf` 用同一個 `name` 對 `DCNUM` 找到的文件連結；查無則空白（`gr_fetch_fields` 有填時會多出 `GR <欄名>` 欄） |

`分類統計` 工作表依筆數排序，`amt 合計` 用 `amt_column` 指定的欄加總（找不到該欄就只算筆數）。

### 設定：`match_pms_docs.ini`（放 exe 同資料夾，從 `.ini.example` 複製；沒有就用預設）

| key | 預設 | 說明 |
|---|---|---|
| `server` / `dbfile` | `UMCD20/UMC` / `eri\pms.nsf` | Notes 資料庫 |
| `docno_view` | `(ByDocNo)` | 第一欄為 DOCNO 的檢視 |
| `docno_field` | `DOCNO` | 單號欄位名 |
| `epms_key_field` | `DOCNO` | Excel key 預設對 ePMS 哪個欄位（`DOCNO` / `A21` / `A22`），視窗可臨時切換 |
| `epms_key_options` | `DOCNO,A21,A22` | 視窗下拉可選的欄位 |
| `epms_enabled` | `true` | 預設是否比對 ePMS（視窗可臨時勾選） |
| `multi_match` | `all` | A21/A22 對到多筆時 `all`＝每筆一列、`first`＝只取第一筆 |
| `field_fallbacks` | `A21:C11,A22:C12,A26:C15` | 備援欄位：另一種表單版本把資料存在 C 欄，主欄位空時改讀備援欄；用 A21/A22 搜尋時也一併搜備援欄（`主:備援`，多個備援用 `/`） |
| `fetch_fields` | `A17,A21,A22,A25,A26,VendorNameList` | 要抓的欄位（A21/A22/VendorNameList 會自動補上） |
| `class_field` | `A17` | 用哪個欄位做分類 |
| `input_xlsx` | （命令列給） | 固定輸入檔路徑 |
| `sheet` / `header_row` | 第一個工作表 / `1` | 工作表與標題列 |
| `key_column` | `name` | Excel 的 key 欄名（不分大小寫） |
| `amt_column` | `amt` | 分類統計加總欄 |
| `output_xlsx` | `<輸入>_notes.xlsx` | 輸出路徑 |
| `sheet_per_category` | `true` | 是否每類一張工作表 |
| `gr_enabled` | `true` | 預設是否比對驗收系統 GR（視窗可臨時勾選） |
| `gr_server` / `gr_dbfile` | `UMCD14/UMC` / `eop\gr.nsf` | GR 資料庫 |
| `gr_view` | `(ByDocNo);(AllByDocNoOnly);(DocNo);(AllByGRNo);(AllByGRNoV2)` | 候選檢視，分號分隔。啟動時用前幾個 key 逐一驗證（查到的文件 `DCNUM` 要等於 key），第一個通過的採用，log 會印 `採用檢視 …`；都不行才退回公式搜尋（慢） |
| `gr_key_field` | `DCNUM` | 與 `name` 比對的 GR 欄位 |
| `gr_fetch_fields` | （空） | 順便抓的 GR 欄位，逗號分隔 |

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
| 2 | ePMS 與驗收系統都沒選（至少要一個） |
| 3 | 找不到 / 無法開啟資料庫 |
| 4 | 無法建立 `Notes.NotesSession`（Notes 未開 / 未登入 / pywin32 未裝 / 位元數不符） |
| 6 | 找不到輸入 Excel |
| 7 | Excel 讀取失敗，或標題列找不到 `key_column` |
| 8 | 寫入輸出 xlsx 失敗（檔案開著？） |

## 常見問題

- **全部 `(查無)`**：先確認 `(ByDocNo)` 檢視第一欄真的是 DOCNO、Excel 的 `name` 值與 Notes `DOCNO` 完全一致（大小寫/空白）。可暫時把 `docno_view` 留空強制走公式搜尋驗證。
- **GR 查詢很慢**：log 若出現 `候選檢視都不適用，改用公式搜尋`，表示 `gr_view` 的候選都不是依 DCNUM 排序的檢視；用 `ListNotesViews.vbs` 找出正確的填到 `gr_view` 最前面。
- **Notes 畫面上 A21/A26 有值，輸出卻空白**：按視窗「檢查欄位…」（或 `--dump <key>`）看該單所有 item。
  - 同名 item 有多個、前面是空的 → 程式已改成取第一個非空的，重新打包即可。
  - 清單裡 `A21` 是空的、值在 `C11` 這類別的 item → 另一種表單版本，把對應加進 `field_fallbacks`（預設已有 `A21:C11,A22:C12,A26:C15`）。
  - 清單裡**完全沒有**那個 item → 表單上看到的是「計算顯示」欄位，文件沒存，程式讀不到；要看它是由哪個 item 算出來的（問系統管理者或用 `main` branch 的 `ExportFormDesign.vbs` 看表單設計），改抓那個來源欄位。
- **Link 點了沒反應**：該台電腦要有 Notes client 並登入；`notes://` 協定由 Notes 處理。
- **ini 中文亂碼 / 讀不到**：程式會自動辨識 UTF-8 / Big5；仍有問題就把 ini 另存 UTF-8。
