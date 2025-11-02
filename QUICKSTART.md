# 快速開始指南

## 📋 概述

這個專案包含一個完整的 n8n 工作流，可以：
1. 接收使用者上傳的 Excel 檔案
2. 讓使用者指定要分析的欄位
3. 使用 AI 自動對資料進行分類
4. 將分類結果填入新的欄位並輸出

## 🚀 5 分鐘快速設定

### 步驟 1: 匯入工作流

1. 開啟您的 n8n 實例
2. 點擊「工作流」→「匯入」
3. 選擇 `excel-ai-classification-workflow-v2.json`
4. 點擊「匯入」

### 步驟 2: 設定 OpenAI API

1. 在工作流中點擊「Call OpenAI API」節點
2. 點擊「憑證」→「新增憑證」
3. 選擇「Header Auth」
4. 設定：
   - 名稱: `Authorization`
   - 值: `Bearer sk-your-api-key-here`
5. 儲存

### 步驟 3: 啟動工作流

1. 點擊右上角的「啟動」開關
2. 複製 Webhook URL（例如：`https://your-n8n.com/webhook/classify-excel`）

### 步驟 4: 測試

使用以下任一方式測試：

#### 選項 A: 使用 Python 腳本

```bash
# 安裝依賴
pip install requests

# 執行測試
python simple_test.py
```

記得先修改 `simple_test.py` 中的 `WEBHOOK_URL`

#### 選項 B: 使用 cURL

```bash
curl -X POST "YOUR_WEBHOOK_URL" \
  -F "file=@user_form_data.xlsx" \
  -F "fields=title,content,description"
```

#### 選項 C: 使用 Postman

1. 匯入 `Postman_Collection.json`
2. 設定環境變數 `webhook_url`
3. 選擇「Upload Excel and Classify」請求
4. 選擇 `user_form_data.xlsx` 檔案
5. 點擊「Send」

## 📁 檔案說明

| 檔案 | 說明 |
|------|------|
| `excel-ai-classification-workflow-v2.json` | 主要工作流檔案（推薦使用） |
| `excel-ai-classification-workflow.json` | 備用工作流檔案 |
| `README_Workflow.md` | 完整使用說明文件 |
| `QUICKSTART.md` | 快速開始指南（本檔案） |
| `simple_test.py` | 簡單測試腳本 |
| `test_workflow.py` | 完整測試腳本（含錯誤處理） |
| `Postman_Collection.json` | Postman 測試集合 |
| `user_form_data.xlsx` | 測試資料（1000 筆） |

## 🎯 使用範例

### 基本使用

```python
import requests

url = "YOUR_WEBHOOK_URL"
files = {'file': open('user_form_data.xlsx', 'rb')}
data = {'fields': 'title,content,description'}

response = requests.post(url, files=files, data=data)
result = response.json()

print(f"分類完成！共 {result['totalRecords']} 筆")
print(f"分類類別: {result['categories']}")
```

### 只分析特定欄位

```python
data = {'fields': 'title,description'}  # 只分析 title 和 description
```

## ⚙️ 自訂化

### 修改 AI 模型

在「Call OpenAI API」節點中修改：

- `gpt-4o-mini` (推薦，快速且便宜)
- `gpt-4` (更準確，但較慢較貴)
- `gpt-3.5-turbo` (快速，準確度中等)

### 調整分類行為

在「Prepare AI Prompt」節點中修改提示詞，可以：
- 指定固定的分類類別
- 調整分類的細緻度
- 改變分類邏輯

## 📊 輸出結果

工作流會在 Excel 檔案中新增以下欄位：

- `category`: 分類類別名稱
- `classification_result`: 分類結果
- `classified_at`: 分類時間
- `row_number`: 原始行號

## ❓ 常見問題

**Q: 處理大量資料很慢？**
A: 預設是逐筆處理。可以修改工作流使用批量處理。

**Q: 分類結果不穩定？**
A: 將 `temperature` 參數降低到 0.1-0.2。

**Q: 如何指定固定分類類別？**
A: 修改「Prepare AI Prompt」節點的提示詞，明確列出類別。

## 🔗 相關資源

- [n8n 官方文件](https://docs.n8n.io/)
- [OpenAI API 文件](https://platform.openai.com/docs/api-reference)
- [完整說明文件](README_Workflow.md)

## 📝 授權

MIT License

