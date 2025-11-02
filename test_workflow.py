"""
測試腳本：用於測試 n8n Excel AI 分類工作流
"""
import requests
import json
import sys

def test_classification_workflow(webhook_url, excel_file_path, fields='title,content,description'):
    """
    測試分類工作流
    
    Args:
        webhook_url: n8n Webhook URL
        excel_file_path: Excel 檔案路徑
        fields: 要分析的欄位，用逗號分隔
    """
    print(f"開始測試分類工作流...")
    print(f"Webhook URL: {webhook_url}")
    print(f"Excel 檔案: {excel_file_path}")
    print(f"分析欄位: {fields}")
    print("-" * 50)
    
    try:
        # 準備檔案和資料
        with open(excel_file_path, 'rb') as f:
            files = {
                'file': (excel_file_path.split('/')[-1], f, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            }
            data = {
                'fields': fields
            }
            
            # 發送請求
            print("正在發送請求到 n8n...")
            response = requests.post(webhook_url, files=files, data=data, timeout=300)
            
            # 檢查回應
            if response.status_code == 200:
                result = response.json()
                print("\n✅ 分類成功！")
                print("\n分類摘要:")
                print(f"  總記錄數: {result.get('totalRecords', 'N/A')}")
                print(f"  分類數量: {result.get('totalCategories', 'N/A')}")
                print(f"\n分類類別:")
                for category in result.get('categories', []):
                    print(f"  - {category}")
                
                print(f"\n分類統計:")
                stats = result.get('categoryStatistics', {})
                for category, count in sorted(stats.items(), key=lambda x: x[1], reverse=True):
                    print(f"  {category}: {count} 筆")
                
                print(f"\n摘要: {result.get('summary', 'N/A')}")
                return True
            else:
                print(f"\n❌ 錯誤: HTTP {response.status_code}")
                print(f"回應內容: {response.text}")
                return False
                
    except FileNotFoundError:
        print(f"❌ 錯誤: 找不到檔案 {excel_file_path}")
        return False
    except requests.exceptions.Timeout:
        print("❌ 錯誤: 請求超時（可能需要較長時間處理大量資料）")
        return False
    except requests.exceptions.RequestException as e:
        print(f"❌ 錯誤: {str(e)}")
        return False
    except Exception as e:
        print(f"❌ 未預期的錯誤: {str(e)}")
        return False

if __name__ == "__main__":
    # 設定參數
    if len(sys.argv) < 2:
        print("使用方法: python test_workflow.py <WEBHOOK_URL> [EXCEL_FILE] [FIELDS]")
        print("\n範例:")
        print("  python test_workflow.py https://your-n8n.com/webhook/classify-excel")
        print("  python test_workflow.py https://your-n8n.com/webhook/classify-excel user_form_data.xlsx")
        print("  python test_workflow.py https://your-n8n.com/webhook/classify-excel user_form_data.xlsx 'title,description'")
        sys.exit(1)
    
    webhook_url = sys.argv[1]
    excel_file = sys.argv[2] if len(sys.argv) > 2 else "user_form_data.xlsx"
    fields = sys.argv[3] if len(sys.argv) > 3 else "title,content,description"
    
    # 執行測試
    success = test_classification_workflow(webhook_url, excel_file, fields)
    
    sys.exit(0 if success else 1)

