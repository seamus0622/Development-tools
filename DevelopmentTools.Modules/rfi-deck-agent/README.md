# rfi-deck-agent（釋疑簡報）

負責人：@seamus0622　ToolId：`DT_RfiDeckAgent`　Ribbon：RTS › 檢討與管理 › 釋疑簡報

在 Revit 擷取視圖並記錄釋疑問題，累積在收件匣，再一鍵產出釋疑簡報 PPT；可選擇同步到 Google 共用簡報。屬於「文書與協作 Agent」的 Revit 端。

## 類型判斷

本模組使用 Revit API（截取視圖、讀取樓層／柱線／選取元件），所以是 **Revit Ribbon 功能模組**（`module.json`），不是可熱替換 Agent Worker（Worker 不得引用 RevitAPI）。

## 使用方式

按「釋疑簡報」後選擇：

| 選項 | 功能 |
|---|---|
| 釋疑擷取 | 截下目前視圖（3D／平面／剖面／立面），帶入樓層、柱線、選取元件，填寫問題、建議與圖說編號後存入收件匣。可「保留草稿」換視圖再擷取，多張截圖累積在同一筆 |
| 產生簡報 | 勾選收件匣項目，呼叫 Python 產生釋疑簡報 PPT，存到 AI AGENT 資料夾的 `04_成果\簡報PPT`；可同時寫入釋疑問題庫（R### 追蹤編號） |
| 共用簡報 | 開啟 Google 共用簡報（需先在設定填 Apps Script 網址與密鑰） |
| 收件匣、成果資料夾與設定 | 開啟資料夾，或設定 AI AGENT 資料夾、Python、Google 連線 |

第一次使用會先要求指定 **AI AGENT 資料夾**（內含 `03_資料庫` 的收件匣、問題庫、圖說索引，以及輸出用的 `04_成果`）。產生簡報與對照圖說用的 Python 程式**已內嵌在外掛中**，隨外掛更新，不再需要放在 AI AGENT 資料夾。

## 輸入、輸出與副作用

| 項目 | 內容 |
|---|---|
| 輸入 | 目前視圖、選取元件、使用者輸入的問題與建議 |
| 輸出 | 收件匣 `item.json` 與截圖；釋疑簡報 `.pptx` |
| 模型 | **不修改**。不開 Transaction，只讀取元素資訊與匯出視圖影像 |
| 本機檔案 | 使用者指定的 AI AGENT 資料夾；`%APPDATA%\DT_AIAgent`（設定）；`%LOCALAPPDATA%\DT_AIAgent`（草稿） |
| 外部程式 | 本機 Python 3。內嵌的 Python 程式解壓到 `%LOCALAPPDATA%\DT_AIAgent\python\<內容雜湊>\` 執行（同版本只解壓一次、舊版保留）。第一次產生簡報時以 `pip install --user` 安裝 `python-pptx`、`openpyxl`；對照圖說另需 `pymupdf`（見 `python/requirements.txt`） |
| 網路（選用） | HTTPS 呼叫使用者自設的 Google Apps Script 網址。密鑰存在 `%APPDATA%\DT_AIAgent\gslides.json`，**不在程式碼內** |
| Bridge／MCP | 不使用。沒有建立第二套 Bridge、Ribbon 啟動器、授權或 Updater |

## 結構

| 路徑 | 內容 |
|---|---|
| `module.json` | 模組描述檔 |
| `RfiDeckAgent.csproj` | SDK 格式 net48 類別庫，沿用 `../Directory.Build.props`，另加 WinForms 與 `System.Web.Extensions` |
| `OpenCommand.cs` | 唯一入口：選單式 TaskDialog，分派到功能本體 |
| `src/*.cs` | 功能本體（擷取、收件匣、簡報、Google 同步、設定），命名空間 `RfiDeckAgent.Core` |
| `python/*.py` | 簡報產生（`revit_deck.py`、`gen_rfi_deck.py`、`pptx_tools.py`）、問題庫（`rfi_db.py`）、對照圖說（`seek_lookup.py`）、路徑與專案設定（`agent_paths.py`）。以 `EmbeddedResource` 內嵌進 `RfiDeckAgent.dll` |

## 來源與修改

來源：`簡報製作及文書處理 AI AGENT\01_AI運行本體\Revit外掛\src\Core`（原獨立外掛 DT AI AGENT）。

- 不帶入原外掛的 `IExternalApplication`、獨立「DT Tools」頁籤、熱重載 Loader、`build_install.ps1`；6 個按鈕改為一個入口加選單（平台每個模組一個 Ribbon 按鈕）。
- 拿掉寫死的 `D:\` 路徑：未設定資料夾時先引導到設定。
- 預設文字去掉特定專案名稱；命名空間由 `DT_AIAgent.Core` 改為 `RfiDeckAgent.Core`。
- 設定與草稿資料夾名稱仍為 `DT_AIAgent`，與舊版資料相容。
- Python 程式來源：`01_AI運行本體\程式`（2026-10-20 版）。改為從環境變數 `DT_AIAGENT_ROOT` 取得 AI AGENT 資料夾（外掛自動設定；單獨執行時仍以程式所在位置往上兩層推算）；問題庫檔名、圖說 PDF 資料夾、柱線對應棟別改由設定決定，程式內不含專案名稱。

## 專案設定（`%APPDATA%\DT_AIAgent\seek.json`，選用）

| 欄位 | 用途 | 省略時 |
|---|---|---|
| `rfi_db` | 釋疑問題庫 xlsx 路徑 | 在 `03_資料庫\釋疑問題庫` 找 `釋疑問題庫.xlsx`，或唯一的 `*釋疑問題庫.xlsx` |
| `index` | 圖說查找索引 | `03_資料庫\圖說\圖說查找索引.db` |
| `pdf_dir` | 圖說 PDF 資料夾 | 依索引中的檔名，在 `03_資料庫\圖說` 或其下一層自動尋找 |
| `building_by_grid` | 數字柱線對應棟別，例如 `[[7, 17, "實驗棟"], [1, 6, "行政棟"]]` | 不由柱線推測棟別（樓層名稱前綴 `LB_`／`OB_` 仍會帶入棟別） |

## 開發者：測試未發版的 Python 程式

在 `%APPDATA%\DT_AIAgent\scripts_dir.txt` 寫入程式資料夾完整路徑（例如本機 clone 的 `DevelopmentTools.Modules\rfi-deck-agent\python`），外掛就會改用該資料夾，不必重新建置 DLL。刪除該檔即恢復使用內嵌版本。

## 已知限制

- 棟別由樓層名稱前綴推得（`LB_` 實驗棟、`OB_` 行政棟），屬事務所命名慣例，寫在 `src/ModelContext.cs`；其他專案辨識不到棟別時欄位留空。後續可比照 DrawingFactory 改為命名規則檔。
- 柱線、樓層與圖說對照依台灣常用圖名用語；「對照圖說」與「同位置釋疑歷史」需要 AI AGENT 資料夾內的圖說索引與 `seek_lookup.py`。

## 不放進 Repo

專案圖資、DWG、建議書、產出成果、含真實人名的範例（客戶資料）；安裝檔、wheels、打包 zip、編譯結果；第三方反向工程資料。會議紀錄、簽到表、雙週會簡報等其他文書程式與範本（會議紀錄、簽到表、雙週會簡報、釋疑問題清單）仍在各使用者的 AI AGENT 資料夾，另案處理；本模組只內嵌 Revit 按鈕會呼叫的 Python 程式。

## 測試

```powershell
python module-platform/validate.py
python -m unittest discover -s module-platform -p "test_*.py" -v
dotnet build DevelopmentTools.sln -c Debug -clp:ErrorsOnly
```

Revit 2024 人工驗收項目見 `module.json` 的 `tests`。新 ToolId `DT_RfiDeckAgent` 需由主管在授權系統核准；未核准前按鈕可見但無法執行。

## 已安裝舊版獨立外掛的使用者

請刪除 `%APPDATA%\Autodesk\Revit\Addins\2024\DT_AIAgent.addin` 與 `DT_AIAgent\` 資料夾，否則會同時看到舊的「DT Tools」頁籤。

## 回退

刪除 `DevelopmentTools.Modules/rfi-deck-agent/` 後重新建置。不影響既有功能、授權或存檔格式。
