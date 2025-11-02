# Excel AI 分類工作流說明

這個 n8n 工作流可以讓使用者上傳 Excel 檔案，指定需要分析的欄位，然後透過 AI 自動對資料進行分類，並將分類結果填入新的欄位。

## 功能特點

- ✅ 接收 Excel 檔案上傳
- ✅ 動態指定要分析的欄位
- ✅ 使用 AI（OpenAI GPT）自動分類
- ✅ AI 自動決定分類類別數量（3-10類）
- ✅ 將分類結果添加到新欄位（`category` 和 `classification_result`）
- ✅ 生成包含分類結果的新 Excel 檔案
- ✅ 提供分類統計資訊

## 工作流程結構

```
1. Webhook 接收請求
   ↓
2. 提取請求參數（Excel檔案 + 要分析的欄位）
   ↓
3. 解析 Excel 檔案
   ↓
4. 提取指定欄位的數據
   ↓
5. 為每筆資料準備 AI 提示詞
   ↓
6. 呼叫 OpenAI API 進行分類
   ↓
7. 合併分類結果到原始數據
   ↓
8. 彙總所有結果並統計
   ↓
9. 匯出為新的 Excel 檔案
   ↓
10. 返回結果給使用者
```

## 設定步驟

### 1. 匯入工作流

1. 開啟 n8n
2. 點擊左上角的「工作流」→「匯入」
3. 選擇 `excel-ai-classification-workflow-v2.json` 檔案
4. 點擊「匯入」

### 2. 設定 OpenAI API 憑證

1. 在工作流中找到「Call OpenAI API」節點
2. 點擊節點，然後點擊「憑證」旁的「新增憑證」
3. 選擇「Header Auth」類型
4. 設定憑證：
   - **名稱**: `Authorization`
   - **值**: `Bearer YOUR_OPENAI_API_KEY`
   - 將 `YOUR_OPENAI_API_KEY` 替換為您的實際 OpenAI API Key

### 3. 啟動工作流

1. 點擊右上角的「啟動」開關，啟用工作流
2. 複製 Webhook URL（格式類似：`https://your-n8n-instance.com/webhook/classify-excel`）

## 使用方法

### API 呼叫方式

#### 使用 cURL

```bash
curl -X POST "YOUR_WEBHOOK_URL" \
  -H "Content-Type: multipart/form-data" \
  -F "file=@user_form_data.xlsx" \
  -F "fields=title,content,description"
```

#### 使用 Python

```python
import requests

url = "YOUR_WEBHOOK_URL"
files = {
    'file': ('user_form_data.xlsx', open('user_form_data.xlsx', 'rb'), 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
}
data = {
    'fields': 'title,content,description'  # 指定要分析的欄位，用逗號分隔
}

response = requests.post(url, files=files, data=data)
result = response.json()

print(f"分類完成！")
print(f"總共處理: {result['totalRecords']} 筆資料")
print(f"分類類別: {result['categories']}")
print(f"統計: {result['categoryStatistics']}")
```

#### 使用 JavaScript/Node.js

```javascript
const FormData = require('form-data');
const fs = require('fs');
const axios = require('axios');

const form = new FormData();
form.append('file', fs.createReadStream('user_form_data.xlsx'));
form.append('fields', 'title,content,description');

axios.post('YOUR_WEBHOOK_URL', form, {
  headers: form.getHeaders()
})
.then(response => {
  console.log('分類完成！');
  console.log('結果:', response.data);
})
.catch(error => {
  console.error('錯誤:', error);
});
```

#### 使用 Postman

1. 選擇 POST 方法
2. 輸入 Webhook URL
3. 在「Body」標籤選擇「form-data」
4. 添加兩個欄位：
   - `file` (類型: File) - 選擇您的 Excel 檔案
   - `fields` (類型: Text) - 輸入 `title,content,description`

### 參數說明

- **file** (必需): 要上傳的 Excel 檔案（.xlsx 格式）
- **fields** (可選): 要分析的欄位名稱，用逗號分隔。預設為 `title,content,description`

### 回應格式

成功時會返回 JSON：

```json
{
  "success": true,
  "summary": "成功分類 1000 筆資料，共 8 個類別",
  "totalRecords": 1000,
  "totalCategories": 8,
  "categories": [
    "專案管理",
    "設備維修",
    "儲存空間",
    "軟體需求",
    "網路問題",
    "帳號管理",
    "資料備份",
    "其他"
  ],
  "categoryStatistics": {
    "專案管理": 250,
    "設備維修": 180,
    "儲存空間": 150,
    "軟體需求": 120,
    "網路問題": 100,
    "帳號管理": 90,
    "資料備份": 70,
    "其他": 40
  },
  "message": "分類完成！"
}
```

## 輸出檔案

工作流會自動生成包含分類結果的 Excel 檔案，新增的欄位包括：

- **category**: 分類類別名稱
- **classification_result**: 分類結果（與 category 相同）
- **classified_at**: 分類時間（ISO 格式）
- **row_number**: 原始行號

## 自訂化

### 調整 AI 模型

在「Call OpenAI API」節點中，可以修改：

```json
{
  "model": "gpt-4o-mini",  // 可以改為 gpt-4, gpt-3.5-turbo 等
  "temperature": 0.3,      // 調整創造性（0-2）
  "max_tokens": 50         // 最大回應長度
}
```

### 修改分類提示詞

在「Prepare AI Prompt」節點中修改 `ai_prompt` 的值，可以調整分類的規則和邏輯。

### 批量處理限制

目前設定為逐筆處理資料。如果資料量很大（>100筆），建議：

1. 使用「Split in Batches」節點分批處理
2. 調整 OpenAI API 的 rate limit 設定
3. 考慮使用批量 API 呼叫（需要修改工作流）

## 常見問題

### Q: 為什麼分類結果不一致？

A: AI 分類有一定的隨機性。可以透過降低 `temperature` 參數（例如改為 0.1）來增加一致性。

### Q: 如何讓 AI 使用固定的分類類別？

A: 修改「Prepare AI Prompt」節點中的提示詞，明確列出您想要的分類類別。

### Q: 處理大量資料時很慢怎麼辦？

A: 
- 考慮使用更快的模型（如 `gpt-4o-mini`）
- 實施分批處理機制
- 增加並發處理數量

### Q: 如何處理 Excel 格式錯誤？

A: 確保上傳的檔案是有效的 .xlsx 格式，且包含指定的欄位名稱。

## 測試資料

可以使用專案中的 `user_form_data.xlsx` 作為測試資料，這個檔案包含 1000 筆表單資料，欄位包括：
- `title`
- `content`
- `description`

## 授權

本工作流使用 MIT 授權。

