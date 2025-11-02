import pandas as pd
import random
import string


# 生成随机的删除原因（增加变化）
def generate_reason():
    """生成删除原因"""
    reasons = [
        '不再需要', '已完成', '預算不足', '客戶取消', '時程變更', '資源不足',
        '需求變更', '技術問題', '管理決策', '合併專案', '優先級調整', '合約終止',
        '公司政策調整', '市場變化', '人員異動', '時機不對', '風險過高',
        '找到更好的替代方案', '上級指示', '不符合效益', '暫緩執行'
    ]
    
    reason = random.choice(reasons)
    
    # 有时候加入一些修饰词，让表达更自然（但不总是）
    if random.random() < 0.3:
        modifiers = ['由於', '因為', '鑑於']
        modifier = random.choice(modifiers)
        return f'{modifier}{reason}'
    
    return reason

# 定义多种表单类型和对应的生成函数
def generate_title_and_content():
    """生成多样化的title和对应的content、description"""
    
    form_types = [
        # 删除專案相关
        {
            'type': 'delete_project',
            'titles': lambda project: [
                f'刪除專案諜{project}',
                f'刪除專案{project}',
                f'請刪除專案諜{project}',
                f'申請刪除專案諜{project}',
                f'需要刪除專案諜{project}',
                f'麻煩刪除專案諜{project}',
            ],
            'descriptions': lambda project, reason: [
                f'因{reason}，請幫我刪除專案諜{project}',
                f'因為{reason}，麻煩刪除專案諜{project}',
                f'由於{reason}，請協助刪除專案諜{project}',
            ],
            'contents': lambda project, reason: [
                f'申請刪除專案{project}，原因：{reason}',
                f'請協助處理專案{project}的刪除作業，原因為{reason}',
            ]
        },
        
        # 個人電腦故障維修
        {
            'type': 'pc_repair',
            'titles': lambda project: [
                '個人電腦故障維修',
                '電腦故障需要維修',
                '申請個人電腦維修',
                '電腦無法開機需要協助',
                '筆電故障申請維修',
                '個人電腦當機問題',
                '電腦維修申請',
                '筆記型電腦故障',
                '個人電腦無法開機',
                '電腦藍屏需要維修',
                '筆電無法充電',
                '電腦硬碟故障',
                '申請IT協助維修電腦',
                '電腦系統異常',
                '個人筆電螢幕故障',
            ],
            'descriptions': lambda project, reason: [
                '個人電腦出現故障，需要申請維修服務',
                '電腦無法正常運作，麻煩協助維修',
                '筆電出現問題，需要維修處理',
                '個人電腦故障，請幫忙安排維修',
                '電腦當機無法使用，需要技術支援',
                '電腦突然無法開機，麻煩協助檢查',
                '筆電充電有問題，需要維修',
                '電腦螢幕出現異常，申請維修',
                '個人電腦系統當機，無法正常使用',
                '電腦硬碟有異音，擔心資料遺失',
            ],
            'contents': lambda project, reason: [
                '我的個人電腦最近出現故障，無法正常開機，希望能安排維修服務',
                '筆記型電腦在使用時突然當機，重開機後仍然無法正常運作，麻煩協助處理',
                '個人電腦故障，螢幕無法顯示，需要申請維修',
                '電腦無法連接網路，且經常自動重開機，需要技術人員協助檢查',
                '昨天電腦突然藍屏，重新開機後還是無法正常使用，麻煩協助維修',
                '筆電無法充電，電池也無法蓄電，需要檢查維修',
                '電腦開機後會自動關機，螢幕有時候會閃爍，需要技術人員協助',
                '個人電腦硬碟運轉時有異常聲音，擔心會損壞，希望可以檢查維修',
                '電腦系統異常緩慢，經常當機，影響工作進度，需要協助處理',
            ]
        },
        
        # 個人諜空間放大
        {
            'type': 'storage_expand',
            'titles': lambda project: [
                '須個人諜空間放大，以供備份資料',
                '申請增加個人儲存空間',
                '個人雲端空間不足，需要擴充',
                '申請擴充個人諜空間',
                '需要增加個人備份空間',
                '個人儲存空間已滿，申請擴大',
                '申請個人雲端空間升級',
                '個人諜空間不足需要擴大',
                '申請擴大個人雲端儲存空間',
                '個人備份空間已滿',
                '需要擴充個人諜空間容量',
                '申請增加個人資料備份空間',
                '個人儲存空間快滿了',
                '申請個人雲端空間擴充',
                '個人諜空間需要升級',
            ],
            'descriptions': lambda project, reason: [
                '個人諜空間不足，需要擴大容量以備份重要資料',
                '儲存空間已滿，申請增加個人諜空間',
                '因備份需求，需要擴大個人雲端儲存空間',
                '個人資料備份空間不足，麻煩協助擴充',
                '個人諜空間快滿了，需要擴大以備份專案資料',
                '儲存空間不足，無法備份重要工作檔案',
                '個人雲端空間即將用完，申請擴充',
            ],
            'contents': lambda project, reason: [
                '目前個人諜空間已接近容量上限，需要擴大空間以備份重要工作資料',
                '申請增加個人儲存空間，因為有大量專案資料需要備份',
                '個人雲端空間不足，希望可以擴充容量以儲存更多備份資料',
                '因工作需要備份大量資料，現有個人諜空間不足，申請擴大容量',
                '最近專案資料量增加很多，個人諜空間已經使用超過80%，需要擴大以備份資料',
                '個人諜空間已經滿了，無法上傳新的備份檔案，麻煩協助擴大空間',
                '因為需要備份多個專案的資料，目前的個人儲存空間不夠用，申請擴大容量',
                '個人雲端空間只剩10%可用，擔心資料無法備份，希望可以盡快擴充',
            ]
        },
        
        # 其他常見類型
        {
            'type': 'project_other',
            'titles': lambda project: [
                f'申請新增專案諜{project}',
                f'專案諜{project}權限申請',
                f'申請修改專案諜{project}設定',
                f'專案諜{project}資料恢復申請',
                f'申請專案諜{project}權限調整',
                f'專案諜{project}名稱修改申請',
                f'申請轉移專案諜{project}',
                f'專案諜{project}成員新增申請',
            ],
            'descriptions': lambda project, reason: [
                f'需要新增專案諜{project}，麻煩協助處理',
                f'申請新增專案{project}的建立權限',
                f'需要修改專案諜{project}的相關設定',
            ],
            'contents': lambda project, reason: [
                f'因業務需要，申請新增專案諜{project}',
                f'需要建立新的專案{project}，麻煩協助開通權限',
                f'申請修改專案諜{project}的名稱和設定',
            ]
        },
        
        {
            'type': 'software',
            'titles': lambda project: [
                '軟體安裝申請',
                '申請安裝特定軟體',
                '需要安裝開發工具',
                '申請軟體授權',
                '軟體使用權限申請',
            ],
            'descriptions': lambda project, reason: [
                '需要安裝特定軟體，申請安裝權限',
                '因工作需求，申請安裝開發工具',
                '需要申請軟體使用授權',
            ],
            'contents': lambda project, reason: [
                '因專案開發需求，需要安裝特定開發工具，麻煩協助處理',
                '申請安裝專業軟體，需要管理員權限協助',
                '工作需要特定軟體授權，麻煩協助申請',
            ]
        },
        
        {
            'type': 'network',
            'titles': lambda project: [
                '網路問題反映',
                '無法連接網路',
                '網路速度緩慢',
                '申請網路權限調整',
                '網路連線異常',
            ],
            'descriptions': lambda project, reason: [
                '辦公室網路連線異常，無法正常上網',
                '網路速度明顯變慢，影響工作效率',
                '申請調整網路使用權限',
            ],
            'contents': lambda project, reason: [
                '辦公室區域網路連線不穩定，經常斷線，麻煩協助檢查',
                '網路速度緩慢，影響日常作業，希望可以改善',
                '需要申請特定網站的訪問權限，麻煩協助調整',
            ]
        },
        
        {
            'type': 'account',
            'titles': lambda project: [
                '帳號密碼重設',
                '忘記密碼申請重設',
                '帳號被鎖定需要解鎖',
                '申請修改登入密碼',
                '帳號權限問題',
            ],
            'descriptions': lambda project, reason: [
                '忘記登入密碼，申請重設',
                '帳號被鎖定，需要協助解鎖',
                '申請修改帳號密碼',
            ],
            'contents': lambda project, reason: [
                '忘記系統登入密碼，麻煩協助重設',
                '帳號因多次錯誤登入被鎖定，需要協助解鎖',
                '因安全考量，申請修改帳號密碼',
            ]
        },
    ]
    
    # 隨機選擇一種表單類型
    form_type = random.choice(form_types)
    
    # 生成project和reason（即使某些類型不需要，也保持接口一致性）
    project = generate_project_ref()
    reason = generate_reason()
    
    # 生成三個字段
    title = random.choice(form_type['titles'](project))
    description = random.choice(form_type['descriptions'](project, reason))
    content = random.choice(form_type['contents'](project, reason))
    
    return title, content, description

def generate_project_ref():
    """生成項目參考編號"""
    formats = [
        lambda: f'P{random.randint(1, 9999)}',
        lambda: f'PROJ-{2020 + random.randint(0, 5)}-{random.randint(1, 999)}',
        lambda: f'專案-{random.randint(1, 999)}',
        lambda: f'PJ{random.choice(string.ascii_uppercase)}{random.randint(100, 999)}',
        lambda: f'A{random.randint(1000, 9999)}',
    ]
    return random.choice(formats)()


# 生成随机数据
def generate_data(num_records):
    """生成num_records笔不重复的数据"""
    data = []
    used_records = set()  # 存储完整记录的哈希，确保完全不重复
    
    for i in range(1, num_records + 1):
        attempts = 0
        max_attempts = 300
        
        while attempts < max_attempts:
            # 使用新的生成函数，生成多样化的title、content和description
            title, content, description = generate_title_and_content()
            
            # 创建完整记录的唯一标识（基于所有三个字段）
            record_key = f"{title}|{content}|{description}"
            
            # 如果这个组合是新的，就添加
            if record_key not in used_records:
                used_records.add(record_key)
                data.append({
                    'title': title,
                    'content': content,
                    'description': description
                })
                break
            
            attempts += 1
            
            # 如果尝试太多次还是重复，添加序号确保唯一
            if attempts >= max_attempts:
                # 在content末尾添加唯一标识，确保不重复
                unique_id = f" [申請編號:{i}-{random.randint(10000, 99999)}]"
                title, base_content, description = generate_title_and_content()
                content = base_content + unique_id
                
                record_key = f"{title}|{content}|{description}"
                used_records.add(record_key)
                data.append({
                    'title': title,
                    'content': content,
                    'description': description
                })
                break
    
    return data

# 生成數據（1000筆不同的資料）
num_records = 1000
print(f'正在生成 {num_records} 筆不重複的資料...')

data = generate_data(num_records)

# 創建 DataFrame
df = pd.DataFrame(data)

# 生成 Excel 文件
output_file = 'user_form_data.xlsx'
# 如果文件被占用，尝试使用临时文件名
import os
try:
    if os.path.exists(output_file):
        os.remove(output_file)
except:
    pass

df.to_excel(output_file, index=False, engine='openpyxl')

print(f'Excel 文件已生成：{output_file}')
print(f'總共 {len(df)} 筆資料')

# 驗證數據不重複
unique_count = len(df.drop_duplicates())
print(f'唯一記錄數：{unique_count} 筆')

# 顯示不同類型的title統計
print('\nTitle 類型統計（前20個最常見的）：')
title_counts = df['title'].value_counts().head(20)
for title, count in title_counts.items():
    print(f'  {title}: {count} 筆')

print('\n前10筆資料預覽：')
for idx, row in df.head(10).iterrows():
    print(f'\n資料 {idx + 1}:')
    print(f'  Title: {row["title"]}')
    print(f'  Description: {row["description"]}')
    print(f'  Content: {row["content"][:80]}...' if len(row["content"]) > 80 else f'  Content: {row["content"]}')

