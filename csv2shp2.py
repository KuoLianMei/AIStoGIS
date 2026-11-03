
# -*- coding: utf-8 -*-

"""
程式名稱：CSV 批次轉 SHP（Field Mapping）
功能：
1. 批次讀取資料夾內 CSV 檔案
2. 利用 Mapping 自動辨識經緯度與識別欄位
3. 轉換為 POINT Shapefile
4. 自動建立輸出資料夾
5. 輸出 WGS84 (EPSG:4326) .prj 座標定義檔

使用前安裝：
python -m pip install pyshp

原始作者：張耘馨
原始日期：20240501
修改：加入 Field Mapping 與 WGS84 座標定義
原始日期：20261103
"""

import os
import csv
import shapefile as shp


# =====================================
# 1. 欄位 Mapping 設定
# =====================================

FIELD_MAPPING = {
    'LONGITUDE': ['LONGITUDE', 'lon', 'LON', 'long', 'X'],
    'LATITUDE': ['LATITUDE', 'lat', 'LAT', 'Y'],
    'SOURCE_ID': ['SOURCE_ID', 'source_id', 'mmsi', 'MMSI']
}


# =====================================
# 2. 自動辨識 CSV 欄位
# =====================================

def find_field(fieldnames, candidates):

    field_lookup = {
        name.strip().lower(): name
        for name in fieldnames
    }

    for candidate in candidates:
        key = candidate.lower()

        if key in field_lookup:
            return field_lookup[key]

    return None


# =====================================
# 3. WGS84 座標系統
# =====================================

WGS84_PRJ = (
    'GEOGCS["WGS 84",'
    'DATUM["WGS_1984",'
    'SPHEROID["WGS 84",6378137,298.257223563]],'
    'PRIMEM["Greenwich",0],'
    'UNIT["degree",0.0174532925199433]]'
)


# =====================================
# 4. CSV 批次轉 SHP
# =====================================

def csv_to_shp(csv_folder, shp_folder):

    os.makedirs(shp_folder, exist_ok=True)

    for filename in os.listdir(csv_folder):

        if not filename.lower().endswith('.csv'):
            continue

        csv_file = os.path.join(csv_folder, filename)

        base_name = os.path.splitext(filename)[0]

        shp_file = os.path.join(
            shp_folder,
            base_name + '.shp'
        )

        print(f"\n正在處理：{filename}")

        with open(
            csv_file,
            'r',
            encoding='utf-8-sig',
            newline=''
        ) as csvfile:

            reader = csv.DictReader(csvfile)

            if not reader.fieldnames:
                print("CSV 無欄位，跳過")
                continue

            # 使用 Mapping 尋找實際欄位
            lon_field = find_field(
                reader.fieldnames,
                FIELD_MAPPING['LONGITUDE']
            )

            lat_field = find_field(
                reader.fieldnames,
                FIELD_MAPPING['LATITUDE']
            )

            source_field = find_field(
                reader.fieldnames,
                FIELD_MAPPING['SOURCE_ID']
            )

            if not lon_field or not lat_field:
                print("找不到經緯度欄位，跳過")
                continue

            print(f"LONGITUDE → {lon_field}")
            print(f"LATITUDE  → {lat_field}")
            print(f"SOURCE_ID → {source_field}")

            count = 0
            skipped = 0

            with shp.Writer(
                shp_file,
                shapeType=shp.POINT
            ) as output_shp:

                output_shp.field(
                    'LONGITUDE', 'F', 18, 8
                )

                output_shp.field(
                    'LATITUDE', 'F', 18, 8
                )

                output_shp.field(
                    'SOURCE_ID', 'C', 100
                )

                for row in reader:

                    try:
                        longitude = float(
                            row[lon_field]
                        )

                        latitude = float(
                            row[lat_field]
                        )

                        # 經緯度範圍檢查
                        if not (
                            -180 <= longitude <= 180
                            and -90 <= latitude <= 90
                        ):
                            skipped += 1
                            continue

                    except (ValueError, TypeError):
                        skipped += 1
                        continue

                    source_id = (
                        str(row.get(source_field, ''))
                        if source_field else ''
                    )

                    output_shp.point(
                        longitude,
                        latitude
                    )

                    output_shp.record(
                        longitude,
                        latitude,
                        source_id
                    )

                    count += 1

        # 建立 WGS84 PRJ
        prj_file = os.path.join(
            shp_folder,
            base_name + '.prj'
        )

        with open(
            prj_file,
            'w',
            encoding='utf-8'
        ) as prj:
            prj.write(WGS84_PRJ)

        print(f"成功轉換：{count} 筆")
        print(f"跳過資料：{skipped} 筆")

    print("\n全部 CSV 轉 SHP 完成")


# =====================================
# 5. 設定資料夾路徑
# =====================================

csv_folder = (
    r'C:\Users\Students\Desktop'
    r'\D&TMOK KLNG_2026_04'
    r'\filter_result_SOB-Decode'
    r'\16型檔\test'
)

shp_folder = (
    r'C:\Users\Students\Desktop'
    r'\D&TMOK KLNG_2026_04'
    r'\filter_result_SOB-Decode'
    r'\16型檔\test\shp'
)


# =====================================
# 6. 執行
# =====================================

csv_to_shp(csv_folder, shp_folder)
