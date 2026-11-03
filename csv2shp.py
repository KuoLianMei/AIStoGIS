Skip to content
 
Search Gists
Search...
All gists
Back to GitHub
@KuoLianMei
KuoLianMei/csv2shp.py Secret
Last active 2 years ago
Code
Revisions
5
Clone this repository at &lt;script src=&quot;https://gist.github.com/KuoLianMei/ca36e835434cc23effb13d9b1d760910.js&quot;&gt;&lt;/script&gt;
<script src="https://gist.github.com/KuoLianMei/ca36e835434cc23effb13d9b1d760910.js"></script>
csv2shp
csv2shp.py
"""
將資料夾內CSV檔轉檔為SHP
含BUG沒有坐標系
使用前請安裝pip install pyshp套件
更新於20240501
坐標系仍於GIS呈現?
新增AIS年月日時分秒
更新於20240506
"""
import os
import shapefile as shp
import csv

def csv_to_shp(csv_folder, shp_folder):
    # Loop through each CSV file in the folder
    for filename in os.listdir(csv_folder):
        if filename.endswith('.csv'):
            csv_file = os.path.join(csv_folder, filename)
            shp_file = os.path.join(shp_folder, filename.split('.')[0] + '.shp')

            # Create a new Shapefile Writer
            output_shp = shp.Writer(shp_file, shp.POINT)
            output_shp.epsg = 104137  # TWD97的EPSG編號


            # Set field names and data types
            output_shp.field('LONGITUDE')
            output_shp.field('LATITUDE')
            output_shp.field('SOURCE_ID')
            output_shp.field('RMC_DATE_YEAR')
            output_shp.field('RMC_DATE_MON')
            output_shp.field('RMC_DATE_DAY')
            output_shp.field('RMC_TAKEN_AT_HOUR')
            output_shp.field('RMC_TAKEN_AT_MIN')
            output_shp.field('RMC_TAKEN_AT_SEC')
            

            # Read data from CSV and write to Shapefile
            with open(csv_file, 'r', encoding='utf-8') as csvfile:
                reader = csv.DictReader(csvfile)
                for row in reader:
                    # Check if LONGITUDE and LATITUDE are valid floats
                    try:
                        longitude = float(row['LONGITUDE'])
                        latitude = float(row['LATITUDE'])
                    except ValueError:
                        # Skip this row if LONGITUDE or LATITUDE is not a valid float
                        continue

                    source_id = row['SOURCE_ID']
                    RMC_DATE_YEAR= row['RMC_DATE_YEAR']
                    RMC_DATE_MON = row['RMC_DATE_MON']
                    RMC_DATE_DAY = row['RMC_DATE_DAY']
                    RMC_TAKEN_AT_HOUR = row['RMC_TAKEN_AT_HOUR']
                    RMC_TAKEN_AT_MIN = row['RMC_TAKEN_AT_MIN']
                    RMC_TAKEN_AT_SEC = row['RMC_TAKEN_AT_SEC']
                    output_shp.point(longitude, latitude)
                    output_shp.record(longitude, latitude, source_id,RMC_DATE_YEAR,
                                      RMC_DATE_MON,RMC_DATE_DAY,RMC_TAKEN_AT_HOUR,RMC_TAKEN_AT_MIN
                                      ,RMC_TAKEN_AT_SEC)

            # Save the Shapefile
            output_shp.close()

# 輸入的CSV資料夾路徑和輸出的SHP資料夾路徑
csv_folder = r'C:\Users\604\Desktop\0417shiptrack'
shp_folder = r'C:\Users\604\Desktop\0417shiptrack\point2'

csv_to_shp(csv_folder, shp_folder)
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
