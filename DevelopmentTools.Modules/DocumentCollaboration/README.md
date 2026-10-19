# DocumentCollaboration（文書與協作 Agent）

> 狀態：**匯入準備中（Draft）**。本檔是第一個 commit，用來確認分支可以正常上傳；程式碼尚未移入。

## 模組介面（依 docs/MODULE_CONTRACT.md）

```yaml
module: DocumentCollaboration
owner: seamus0622
tool_id: DT_DocumentCollaboration   # 已全 Repo 搜尋，無重複
inputs: [會議資料 JSON, 釋疑問題清單 xlsx, Revit 視圖截圖]
outputs: [會議紀錄 docx, 簽到表 docx, 雙週會簡報 pptx, 釋疑簡報 pptx]
side_effects: [none]                # 目前只讀模型、只寫輸出檔
authorized_calls: [待確認]
revit_versions: [2024]
tests:
  static: required
  unit: when feasible
  build: Windows environment required
  live_revit: manual signoff required
```

## 預計匯入內容（待主管確認）

| 來源 | 去處 | 狀態 |
|---|---|---|
| Python 產生器（會議紀錄、簽到表、雙週會簡報、釋疑簡報） | `python/` | 待移入 |
| 工作流程說明、用語規範 | `docs/` | 待移入 |
| 互動 3D 簡報引擎 | `web3d/` | 待移入 |
| Google 簡報同步（Apps Script） | `gas/` | 待移入 |
| 圖面查找 SEEKER | `seeker/` | 待移入 |
| Revit 外掛原始碼 | — | **NEEDS_OWNER_APPROVAL**：自帶 IExternalApplication 與 Ribbon，需改接既有外掛 |

## 不放進 Repo 的檔案

- 專案圖資、DWG、建議書、產出成果（客戶資料）
- 安裝檔、wheels、打包 zip（二進位檔）
- 第三方反向工程資料（授權不明）
- 含真實人名、專案名的範例資料：移入前改為假資料

## 需要的範本（不放進 Repo，由主管決定存放位置）

| 範本 | 用途 |
|---|---|
| 會議紀錄範本.docx | 會議紀錄產生器 |
| 簽到表範本.docx | 簽到表產生器 |
| BIM雙週會簡報範本.pptx | 雙週會簡報產生器 |
| 釋疑問題清單_空白範本.xlsx | 釋疑簡報的輸入格式 |

## 回復方式

本模組目前沒有改動共用核心；刪除 `DevelopmentTools.Modules/DocumentCollaboration/` 即可完整回復。
