"""
簡單測試腳本：快速測試 n8n Excel AI 分類工作流
"""
import requests

# 設定您的 Webhook URL
WEBHOOK_URL = "https://your-n8n-instance.com/webhook/classify-excel"

# 設定 Excel 檔案路徑
EXCEL_FILE = "user_form_data.xlsx"

# 設定要分析的欄位
FIELDS = "title,content,description"

def main():
    print("🚀 開始測試 n8n Excel AI 分類工作流\n")
    
    # 讀取 Excel 檔案
    try:
        with open(EXCEL_FILE, 'rb') as file:
            files = {'file': (EXCEL_FILE, file, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')}
            data = {'fields': FIELDS}
            
            print(f"📤 上傳檔案: {EXCEL_FILE}")
            print(f"📋 分析欄位: {FIELDS}\n")
            
            # 發送請求
            response = requests.post(WEBHOOK_URL, files=files, data=data, timeout=600)
            
            if response.status_code == 200:
                result = response.json()
                
                print("✅ 分類完成！\n")
                print("=" * 60)
                print(f"📊 總記錄數: {result.get('totalRecords', 0)}")
                print(f"🏷️  分類數量: {result.get('totalCategories', 0)}")
                print("=" * 60)
                
                print("\n📈 分類統計:")
                stats = result.get('categoryStatistics', {})
                for cat, count in sorted(stats.items(), key=lambda x: x[1], reverse=True):
                    percentage = (count / result.get('totalRecords', 1)) * 100
                    bar = "█" * int(percentage / 2)
                    print(f"  {cat:20s} {count:4d} 筆 ({percentage:5.1f}%) {bar}")
                
                print(f"\n💡 {result.get('summary', '')}\n")
            else:
                print(f"❌ 錯誤: HTTP {response.status_code}")
                print(f"回應: {response.text}")
                
    except FileNotFoundError:
        print(f"❌ 找不到檔案: {EXCEL_FILE}")
        print("請確認檔案路徑正確")
    except Exception as e:
        print(f"❌ 發生錯誤: {str(e)}")

if __name__ == "__main__":
    main()

