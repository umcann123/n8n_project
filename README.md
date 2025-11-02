# n8n Excel AI 分類工作流

這是一個使用 n8n 和 AI 對 Excel 資料進行自動分類的完整解決方案。

## 🎯 功能特點

- 📤 透過表單上傳 Excel 檔案
- 🔍 動態指定要分析的欄位
- 🤖 使用 AI（OpenAI GPT）自動分類
- 📊 AI 自動決定分類類別數量
- 📝 將分類結果添加到新欄位
- 📥 匯出包含分類結果的 Excel 檔案
- 📈 提供分類統計資訊

## 📁 專案結構

```
.
├── README.md                              # 本文件
├── README_Workflow.md                     # 完整工作流使用說明
├── QUICKSTART.md                          # 快速開始指南
├── excel-ai-classification-workflow-v3.json   # 工作流檔案（推薦）
├── excel-ai-classification-workflow-v4.json   # 兩階段分類工作流
├── excel-ai-classification-workflow-v2.json   # 備用版本
├── excel-ai-classification-workflow.json      # 原始版本
├── generate_excel_data.py                 # 生成測試資料腳本
├── test_workflow.py                       # 測試腳本
├── simple_test.py                         # 簡單測試腳本
├── verify_data.py                         # 資料驗證腳本
├── Postman_Collection.json                 # Postman 測試集合
└── user_form_data.xlsx                    # 測試資料（1000筆）
```

## 🚀 快速開始

### 1. 匯入工作流

1. 開啟您的 n8n 實例
2. 點擊「工作流」→「匯入」
3. 選擇 `excel-ai-classification-workflow-v3.json`
4. 點擊「匯入」

### 2. 設定 OpenAI API

1. 在工作流中找到 AI 節點（Classify with AI Agent）
2. 點擊「憑證」→「新增憑證」
3. 選擇「OpenAI API」
4. 輸入您的 API Key

### 3. 啟動工作流

1. 點擊右上角的「啟動」開關
2. 複製表單 URL
3. 在瀏覽器中開啟表單
4. 上傳 Excel 檔案並指定欄位

## 📋 工作流版本說明

### V3 - 單階段分類（推薦）
- 逐行分類，每筆資料獨立分析
- 適合需要彈性分類的場景
- **檔案**: `excel-ai-classification-workflow-v3.json`

### V4 - 兩階段分類
- 第一階段：分析樣本建立分類體系
- 第二階段：基於統一標準進行分類
- 適合需要統一分類體系的場景
- 避免產生過多類別
- **檔案**: `excel-ai-classification-workflow-v4.json`

## 🔧 使用方式

### 表單提交

工作流提供了一個表單介面，您可以直接：
1. 上傳 Excel 檔案（.xlsx）
2. 輸入要分析的欄位（用逗號分隔，例如：`title,content,description`）
3. 提交後等待處理完成

### API 呼叫（V2版本）

如果您使用 V2 版本（Webhook），可以使用：

```bash
curl -X POST "YOUR_WEBHOOK_URL" \
  -F "file=@user_form_data.xlsx" \
  -F "fields=title,content,description"
```

或使用 Python：

```python
import requests

url = "YOUR_WEBHOOK_URL"
files = {'file': open('user_form_data.xlsx', 'rb')}
data = {'fields': 'title,content,description'}

response = requests.post(url, files=files, data=data)
result = response.json()
print(result)
```

## 📊 輸出結果

工作流會在原始 Excel 資料中新增以下欄位：

- `category`: 分類類別名稱
- `classification_result`: 分類結果（與 category 相同）
- `classified_at`: 分類時間（ISO 格式）
- `row_number`: 原始行號

## 🧪 測試

使用提供的測試腳本：

```bash
# 簡單測試
python simple_test.py

# 完整測試（含錯誤處理）
python test_workflow.py YOUR_WEBHOOK_URL user_form_data.xlsx "title,content,description"
```

## 📝 生成測試資料

如果需要生成更多測試資料：

```bash
python generate_excel_data.py
```

這會生成包含 1000 筆不重複資料的 Excel 檔案。

## 🔍 詳細文件

- **完整使用說明**: 請參考 [README_Workflow.md](README_Workflow.md)
- **快速開始指南**: 請參考 [QUICKSTART.md](QUICKSTART.md)

## ⚙️ 技術架構

- **n8n**: 工作流自動化平台
- **OpenAI GPT-4o-mini**: AI 分類模型
- **Python**: 測試腳本和資料生成工具

## 📄 授權

MIT License

## 🤝 貢獻

歡迎提交 Issue 或 Pull Request！

