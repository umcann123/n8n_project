# Excel AI 分類工作流 V5 - 詳細說明文檔

## 📋 目錄
1. [工作流概述](#工作流概述)
2. [整體流程圖](#整體流程圖)
3. [節點詳細說明](#節點詳細說明)
4. [使用範例](#使用範例)
5. [參數說明](#參數說明)

---

## 工作流概述

這個工作流實現了**關鍵詞快速匹配 + AI 智能分類**的混合分類系統，能夠：
- 🚀 使用關鍵詞快速分類（節省 API 費用）
- 🤖 對未匹配的資料使用 AI 進行智能分類
- 📊 批量處理資料（可調整批次大小）
- 📁 輸出包含分類結果的 Excel 檔案

---

## 整體流程圖

```
┌─────────────────────────────────────────────────────────────────────────┐
│                        工作流程示意圖                                    │
└─────────────────────────────────────────────────────────────────────────┘

使用者提交表單
    │
    ├─► [1] On form submission ──► 接收檔案、欄位、類別關鍵詞、批次大小
    │
    ├─► [2] Extract Excel Data ──► 解析 Excel 檔案為 JSON 格式
    │
    ├─► [3] Parse Categories ──► 解析使用者輸入的類別和關鍵詞
    │                             輸出：所有資料 + 類別關鍵詞字典
    │
    ├─► [4] Keyword Classification ──► 使用關鍵詞快速匹配
    │                                   │
    │                                   ├─► keywordMatched: 已匹配的資料
    │                                   └─► unmatched: 未匹配的資料
    │
    ├─► [5] Prepare AI Classification ──► 將未匹配資料分批（預設100筆/批）
    │                                      │
    │                                      └─► 為每批準備 AI 提示詞
    │
    ├─► [6] AI Classify Rows ──► 批量調用 AI 進行分類
    │                             （使用 OpenAI GPT-5-mini 模型）
    │
    ├─► [7] Merge AI Results ──► 解析 AI 回應並合併結果
    │                            輸出：AI 分類的資料列表
    │
    ├─► [8] Merge All Results ──► 合併關鍵詞匹配 + AI 分類的結果
    │                             │
    │                             └─► 按原始順序排序
    │
    ├─► [9] Clean Output ──► 清理輸出，只保留原始欄位 + classification_result
    │
    └─► [10] Export to Excel ──► 輸出為 Excel 檔案

數據流向：
[表單] → [提取] → [解析] → [關鍵詞匹配] ──┐
                                          │
                                          ├─► [合併結果] → [清理] → [輸出]
                                          │
[未匹配資料] → [分批] → [AI分類] ─────────┘
```

---

## 節點詳細說明

### 📝 [1] On form submission (表單觸發器)
**節點類型**: `n8n-nodes-base.formTrigger` (typeVersion: 2.2)

**功能**: 接收使用者通過網頁表單提交的資料

**輸入**: 無（這是工作流的起點）

**輸出**:
```json
{
  "file": "<檔案二進位數據>",
  "fields": "title,content,description",
  "categories": "專案管理:專案,項目\n設備維修:故障,壞掉",
  "batch_size": "100"
}
```

**參數說明**:
- `file`: Excel 檔案（必須）
- `fields`: 要分析的欄位，用逗號分隔（必須）
- `categories`: 類別和關鍵詞，每行一個類別（必須）
- `batch_size`: 每批處理筆數（選填，預設100）

---

### 📄 [2] Extract Excel Data (提取 Excel 資料)
**節點類型**: `n8n-nodes-base.extractFromFile` (typeVersion: 1)

**功能**: 從上傳的 Excel 檔案中提取資料並轉換為 JSON 格式

**輸入**: 
- 來自 [1] 的檔案二進位數據

**輸出**: 
多個 items，每個 item 代表 Excel 的一行資料
```json
[
  { "json": { "title": "申請新專案", "content": "...", "description": "..." } },
  { "json": { "title": "電腦故障", "content": "...", "description": "..." } },
  ...
]
```

**處理邏輯**:
- 讀取 Excel 檔案
- 將每一行轉換為一個 JSON 物件
- 保持原始欄位名稱

---

### 🔍 [3] Parse Categories and Keywords (解析類別和關鍵詞)
**節點類型**: `n8n-nodes-base.code` (typeVersion: 2)

**功能**: 解析使用者輸入的類別關鍵詞，並為每筆資料準備分析文字

**輸入**: 
- 來自 [2] 的所有 Excel 資料行
- 從表單獲取的 `fields`、`categories`、`batch_size`

**輸出**: 
單一 item，包含所有資料和解析後的類別資訊
```json
{
  "fieldsToAnalyze": ["title", "content", "description"],
  "categoryKeywords": {
    "專案管理": ["專案", "項目", "project"],
    "設備維修": ["故障", "壞掉", "維修"]
  },
  "categoryList": ["專案管理", "設備維修"],
  "totalRows": 250,
  "allRowsData": [
    {
      "rowIndex": 0,
      "originalData": { "title": "...", "content": "..." },
      "analysisText": "title: ...\ncontent: ...",
      "combinedText": "title: ...\ncontent: ...",  // 小寫，用於匹配
      "fieldsUsed": ["title", "content", "description"]
    },
    ...
  ],
  "batchSize": 100
}
```

**處理邏輯**:
```javascript
1. 解析 categories 字串（格式：類別:關鍵詞1,關鍵詞2\n類別2:關鍵詞3...）
2. 建立 categoryKeywords 字典和 categoryList 陣列
3. 為每筆資料提取指定欄位並組合分析文字
4. 生成 combinedText（小寫）用於關鍵詞匹配
```

---

### ⚡ [4] Keyword Classification (關鍵詞快速分類)
**節點類型**: `n8n-nodes-base.code` (typeVersion: 2)

**功能**: 使用關鍵詞進行快速匹配分類，節省 AI API 費用

**輸入**: 
- 來自 [3] 的資料和類別關鍵詞

**輸出**: 
單一 item，包含已匹配和未匹配的資料
```json
{
  "keywordMatched": [
    {
      "rowIndex": 0,
      "originalData": { ... },
      "category": "專案管理",
      "classification_method": "keyword_match",
      "matched_keywords": ["專案", "項目"]
    },
    ...
  ],
  "unmatched": [
    {
      "rowIndex": 5,
      "originalData": { ... },
      ...
    },
    ...
  ],
  "keywordMatchedCount": 150,
  "unmatchedCount": 100,
  ...
}
```

**處理邏輯**:
```javascript
對每筆資料：
1. 檢查 combinedText（小寫）是否包含任何類別的關鍵詞
2. 如果找到匹配：
   - 加入 keywordMatched 陣列
   - 標記為 'keyword_match'
   - 記錄匹配的關鍵詞
3. 如果沒找到匹配：
   - 加入 unmatched 陣列，等待 AI 分類
```

**匹配邏輯**:
- 使用 `includes()` 進行字串包含檢查
- 不區分大小寫
- 只要包含任一關鍵詞即匹配（第一個匹配的類別）

---

### 🤖 [5] Prepare AI Classification (準備 AI 分類)
**節點類型**: `n8n-nodes-base.code` (typeVersion: 2)

**功能**: 將未匹配的資料分批，並為每批準備 AI 提示詞

**輸入**: 
- 來自 [4] 的 `unmatched` 資料

**輸出**: 
多個 items（每批一個），每個包含：
```json
{
  "batchIndex": 0,
  "batchData": [ /* 100筆資料 */ ],
  "batchSize": 100,
  "prompt": "你是一個專業的數據分類系統...\n數據列表：...\n請以JSON格式返回...",
  "categoryKeywords": { ... },
  "categoryList": [ ... ]
}
```

**處理邏輯**:
```javascript
1. 從表單獲取 batch_size（預設100）
2. 將 unmatched 資料分成批次：
   - 第1批：索引 0-99
   - 第2批：索引 100-199
   - ...
3. 為每批建立表格格式的資料展示
4. 構建 AI 提示詞，包含：
   - 類別和關鍵詞參考
   - 批次資料表格（編號 + 內容摘要）
   - 要求 JSON 格式回應
```

**批次表格格式**:
```
編號 | 數據內容
--- | ---
1 | title: 申請新專案 content: 需要建立...
2 | title: 電腦故障 content: 無法開機...
...
```

---

### 🧠 [6] AI Classify Rows (AI 分類節點)
**節點類型**: `@n8n/n8n-nodes-langchain.agent` (typeVersion: 2.2)

**功能**: 調用 AI 模型對批量資料進行分類

**連接的模型**: `@n8n/n8n-nodes-langchain.lmChatOpenAi` (OpenAI GPT-5-mini)

**輸入**: 
- 來自 [5] 的每個批次 item（每批包含 prompt）

**輸出**: 
每批一個回應，包含 AI 的分類結果：
```json
{
  "output": "{\n  \"results\": [\n    { \"編號\": 1, \"類別\": \"專案管理\" },\n    { \"編號\": 2, \"類別\": \"設備維修\" },\n    ...\n  ]\n}"
}
```

**處理邏輯**:
- AI 接收包含多筆資料的提示詞
- 分析每筆資料的內容
- 參考提供的類別和關鍵詞
- 返回 JSON 格式的分類結果

**優勢**:
- 一次處理多筆（預設100筆），大幅節省 API 調用
- AI 可以參考關鍵詞資訊，提高準確度
- 可以創建新類別（如果都不符合）

---

### 🔄 [7] Merge AI Classification Results (合併 AI 分類結果)
**節點類型**: `n8n-nodes-base.code` (typeVersion: 2)

**功能**: 解析 AI 的回應，將分類結果映射回原始資料

**輸入**: 
- 來自 [6] 的 AI 回應（多個批次）
- 來自 [5] 的批次原始資料（用於映射）

**輸出**: 
單一 item，包含所有 AI 分類的結果
```json
{
  "aiClassified": [
    {
      ...originalData,
      "classification_result": "專案管理",
      "classification_method": "ai",
      "classified_at": "2024-01-01T12:00:00.000Z",
      "row_number": 6
    },
    ...
  ],
  "skipAI": false
}
```

**處理邏輯**:
```javascript
對每個批次：
1. 提取 AI 回應文字（從 output/text/message.content）
2. 解析 JSON（移除 markdown 代碼塊）
3. 將結果映射回原始資料：
   - 使用"編號"找到對應的資料
   - 提取"類別"名稱
   - 合併到 originalData
4. 如果解析失敗：
   - 嘗試逐行解析（降級處理）
   - 或設置默認類別"其他"
```

**錯誤處理**:
- JSON 解析失敗 → 嘗試逐行解析
- 找不到對應資料 → 記錄錯誤並跳過
- 完全失敗 → 設置為"其他"類別

---

### 🔀 [8] Merge All Results (合併所有結果)
**節點類型**: `n8n-nodes-base.code` (typeVersion: 2)

**功能**: 合併關鍵詞匹配和 AI 分類的所有結果，按原始順序排列

**輸入**: 
- 來自 [4] 的 `keywordMatched`（通過關鍵詞引用）
- 來自 [7] 的 `aiClassified`

**輸出**: 
多個 items，每個代表一筆最終結果
```json
[
  {
    "json": {
      ...originalData,  // 原始 Excel 欄位
      "classification_result": "專案管理",
      "classification_method": "keyword_match",
      "matched_keywords": ["專案"],
      "classified_at": "...",
      "row_number": 1
    }
  },
  {
    "json": {
      ...originalData,
      "classification_result": "設備維修",
      "classification_method": "ai",
      "classified_at": "...",
      "row_number": 6
    }
  },
  ...
]
```

**處理邏輯**:
```javascript
1. 從 Keyword Classification 節點獲取關鍵詞匹配結果
2. 將關鍵詞結果轉換為標準格式（加上 classification_method）
3. 獲取 AI 分類結果
4. 合併兩個陣列：[...keywordResults, ...aiClassified]
5. 按 row_number 排序（確保與原始 Excel 順序一致）
```

---

### 🧹 [9] Clean Output (清理輸出)
**節點類型**: `n8n-nodes-base.code` (typeVersion: 2)

**功能**: 移除中間處理欄位，只保留原始欄位和分類結果

**輸入**: 
- 來自 [8] 的所有結果 items

**輸出**: 
清理後的 items，只包含原始欄位和 `classification_result`
```json
[
  {
    "json": {
      "title": "申請新專案",
      "content": "需要建立一個新專案...",
      "description": "...",
      "classification_result": "專案管理"  // 唯一新增的欄位
    }
  },
  ...
]
```

**處理邏輯**:
```javascript
排除的欄位：
- classification_method
- matched_keywords
- classified_at
- row_number
- category
- error

保留：
- 所有原始 Excel 欄位（title, content, description 等）
- classification_result（分類結果）
```

**目的**: 
輸出乾淨的 Excel 檔案，只包含原始資料和分類結果，不包含中間處理資訊

---

### 📊 [10] Export to Excel (導出 Excel)
**節點類型**: `n8n-nodes-base.convertToFile` (typeVersion: 1.1)

**功能**: 將清理後的資料轉換為 Excel 檔案

**輸入**: 
- 來自 [9] 的清理後資料 items

**輸出**: 
Excel 檔案（.xlsx 格式）

**檔案內容**:
- 所有原始欄位（title, content, description 等）
- 新增 `classification_result` 欄位
- 保持原始資料順序

---

## 使用範例

### 表單輸入範例

**fields**:
```
title,content,description
```

**categories**:
```
專案管理:專案,項目,project,建立專案
設備維修:故障,壞掉,維修,當機,電腦
儲存空間:空間不足,容量,擴充,備份,硬碟
軟體需求:安裝,申請軟體,授權,軟體安裝
網路問題:連線,網路,斷線,無法上網
帳號管理:密碼,帳號,登入,權限,忘記密碼
```

**batch_size**:
```
100
```
（可選，不填則使用預設值 100）

### 執行流程範例

假設有 250 筆資料：

```
1. 關鍵詞匹配：150 筆直接匹配 ✅
2. 未匹配：100 筆
3. AI 分類：100 筆分成 1 批處理
4. 合併：150 + 100 = 250 筆
5. 輸出：250 筆帶分類結果的 Excel
```

**API 調用次數**:
- 舊方式（逐筆）：250 次
- 新方式（批量）：1 次
- **節省：99.6% 的 API 調用**

---

## 參數說明

### batch_size（批次大小）

**建議值**:
- **50 筆**: 適合資料較長或需要較高準確度
- **100 筆**: 預設值，平衡效率與品質 ⭐
- **150-200 筆**: 適合資料較短或追求最大效率
- **超過 200 筆**: 不建議，可能超過 token 限制

**影響**:
- 批次越大 → API 調用越少 → 費用越低
- 批次太大 → Prompt 過長 → 可能超過 token 限制或影響準確度

### categories（類別格式）

**格式**:
```
類別名稱:關鍵詞1,關鍵詞2,關鍵詞3
類別名稱2:關鍵詞4,關鍵詞5
```

**規則**:
- 每行一個類別
- 用 `:` 分隔類別名稱和關鍵詞
- 用 `,` 分隔多個關鍵詞
- 關鍵詞匹配是**不區分大小寫**的

**範例**:
```
專案管理:專案,項目,project,建立專案,創建項目
設備維修:故障,壞掉,維修,repair,fix,無法使用
```

---

## 資料流程總結

```
[表單提交]
    ↓
[Excel 資料] (250筆)
    ↓
[解析類別關鍵詞] → 建立關鍵詞字典
    ↓
[關鍵詞匹配]
    ├─► 150筆匹配成功 → keywordMatched
    └─► 100筆未匹配 → unmatched
    ↓
[分批處理] (batch_size=100)
    └─► 1批（100筆）
    ↓
[AI 分類] → 1次 API 調用
    ↓
[解析 AI 回應]
    └─► 100筆分類結果
    ↓
[合併結果]
    ├─► 150筆（關鍵詞）
    └─► 100筆（AI）
    = 250筆總計
    ↓
[清理輸出] → 移除中間欄位
    ↓
[Excel 輸出] → 250筆 + classification_result
```

---

## 故障排除

### 1. 節點消失
- **問題**: Form Trigger 或 Langchain Agent 節點顯示為 "unknown node"
- **解決**: 
  - 檢查 n8n 版本（建議 1.88.0+）
  - 確認已安裝 Langchain 套件
  - 檢查節點版本號是否匹配

### 2. AI 分類失敗
- **問題**: Merge AI Results 無法解析回應
- **解決**: 
  - 檢查批次大小是否過大（降低 batch_size）
  - 查看錯誤日誌了解具體原因
  - 確認 AI 模型回應格式正確

### 3. 分類結果不正確
- **問題**: 分類結果與預期不符
- **解決**:
  - 檢查關鍵詞是否完整
  - 調整批次大小
  - 確認類別格式正確

---

## 技術細節

### 關鍵詞匹配算法
```javascript
// 簡單的字串包含匹配（不區分大小寫）
combinedText.toLowerCase().includes(keyword.toLowerCase())
```

### AI 批次處理
- 使用表格格式展示批次資料
- 要求 JSON 格式回應，便於解析
- 有降級處理機制（逐行解析）

### 資料追蹤
- 使用 `rowIndex` 追蹤原始資料順序
- 使用 `row_number` 確保最終排序正確
- 使用 `classification_method` 標記分類方式

---

## 版本資訊

- **工作流版本**: V5
- **n8n 最低版本**: 1.0.0
- **建議 n8n 版本**: 1.88.0+
- **最後更新**: 2024

---

## 授權與支援

如有問題或建議，請聯繫開發團隊。

