import pandas as pd

df = pd.read_excel('user_form_data.xlsx')

print(f'總共 {len(df)} 筆資料')
print(f'\n唯一記錄數：{len(df.drop_duplicates())} 筆')

print('\n=== Title類型統計（前15個）===')
print(df['title'].value_counts().head(15))

print('\n\n=== 前10筆資料示例 ===')
for idx, row in df.head(10).iterrows():
    print(f"\n第 {idx+1} 筆:")
    print(f"  Title: {row['title']}")
    print(f"  Description: {row['description']}")
    print(f"  Content: {row['content'][:50]}...")

print('\n\n=== 檢查重複記錄 ===')
duplicates = df[df.duplicated()]
if len(duplicates) > 0:
    print(f'發現 {len(duplicates)} 筆重複記錄')
else:
    print('[OK] 沒有重複記錄')

