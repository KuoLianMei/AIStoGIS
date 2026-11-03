Skip to content
 
Search Gists
Search...
All gists
Back to GitHub
@KuoLianMei
KuoLianMei/gist:baa05a0a35051c4f97097f8dc2cef2da
Last active 4 minutes ago
Code
Revisions
2
Clone this repository at &lt;script src=&quot;https://gist.github.com/KuoLianMei/baa05a0a35051c4f97097f8dc2cef2da.js&quot;&gt;&lt;/script&gt;
<script src="https://gist.github.com/KuoLianMei/baa05a0a35051c4f97097f8dc2cef2da.js"></script>
ExcelToCsv.py
gistfile1.txt
"""
資料夾內整批excel 轉存為csv (解決中文亂碼問題)
張耘馨
2024.05.25
"""
import os
import pandas as pd

# 指定包含Excel文件的文件夹路径
folder_path = r'c:\Users\Students\Desktop\數量觀測\D&TMOK HWLN_2025_04'

# 取文件夾内所有文件的列表
file_list = os.listdir(folder_path)

# 迴圈處理每個文件
for file_name in file_list:
    if file_name.endswith('.xlsx') or file_name.endswith('.xls'):
        # 建完整的文件路径
        file_path = os.path.join(folder_path, file_name)

        # 讀取Excel文件
        df = pd.read_excel(file_path)

        # 建CSV文件的文件名
        csv_file_name = file_name.replace('.xlsx', '.csv').replace('.xls', '.csv')
        csv_file_path = os.path.join(folder_path, csv_file_name)

        # 将數據保存為CSV文件 (帶 BOM，Excel 開啟不會亂碼)
        df.to_csv(csv_file_path, index=False, encoding="utf-8-sig")

        print(f"已將 '{file_name}' 另存為 '{csv_file_name}'")
@KuoLianMei
Comment
 

Leave a comment
Footer
© 2026 GitHub, Inc.
Footer navigation
Terms
Privacy
Security
Status
Community
Docs
Contact
Manage cookies
Do not share my personal information
