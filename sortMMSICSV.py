# -*- coding: utf-8 -*-

import pandas as pd
import os
import re
import tkinter as tk
from tkinter import filedialog

# =====================================================
# 欄位名稱兼容處理
# =====================================================

column_mapping = {
    "msg_type": "msg_type",
    "MESSAGE_ID": "msg_type",

    "Year": "GPS_year",
    "RMC_DATE_YEAR": "GPS_year",
    "Month": "GPS_month",
    "RMC_DATE_MON": "GPS_month",
    "Day": "GPS_day",
    "RMC_DATE_DAY": "GPS_day",

    "ship_type": "ship_type",
    "TYPE_OF_SHIP_AND_CARGO_TYPE": "ship_type",

    "MMSI": "mmsi",
    "SOURCE_ID": "mmsi",

    "NAME": "shipname",
    "shipname": "shipname",

    "CALL_SIGN": "callsign",
    "callsign": "callsign",

    "IMO_NUMBER": "imo",
    "imo": "imo",

    "LENGTH": "VESSEL_LENGTH",
    "DESTINATION": "destination",
    "destination": "destination",
}


# =====================================================
# 1. 讀取 Excel / CSV
# =====================================================

def read_ais_file(file_path):

    ext = os.path.splitext(file_path)[1].lower()

    # ---------- Excel ----------
    if ext in [".xlsx", ".xls"]:

        df = pd.read_excel(
            file_path,
            dtype=str
        )

        return df

    # ---------- CSV ----------
    elif ext == ".csv":

        # 嘗試不同文字編碼
        encodings = [
            "utf-8-sig",
            "utf-8",
            "cp950",
            "big5",
            "gb18030"
        ]

        last_error = None

        for encoding in encodings:

            try:

                df = pd.read_csv(
                    file_path,
                    encoding=encoding,
                    dtype=str,
                    low_memory=False
                )

                print(f"   ✔ CSV 編碼：{encoding}")

                return df

            except UnicodeDecodeError as e:
                last_error = e
                continue

        raise ValueError(
            f"CSV 無法使用支援的編碼讀取：{last_error}"
        )

    else:
        raise ValueError("不支援的檔案格式")


# =====================================================
# 2. MMSI 欄位清理
# =====================================================

def clean_mmsi(value):

    if pd.isna(value):
        return None

    value = str(value).strip()

    # 處理 Excel 可能產生的 .0
    if re.fullmatch(r"\d+\.0", value):
        value = value[:-2]

    # 空值排除
    if value.lower() in [
        "",
        "nan",
        "none",
        "null"
    ]:
        return None

    # MMSI 應為 9 位數字
    if not re.fullmatch(r"\d{9}", value):
        return None

    return value


# =====================================================
# 3. 選擇資料夾
# =====================================================

root = tk.Tk()
root.withdraw()

folder_path = filedialog.askdirectory(
    title="請選擇 AIS Excel / CSV 資料夾"
)

root.destroy()

if not folder_path:
    print("❌ 未選擇資料夾，程式結束")
    raise SystemExit


# =====================================================
# 4. 建立輸出資料夾
# =====================================================

output_base = os.path.join(
    folder_path,
    "shiptrack"
)

os.makedirs(
    output_base,
    exist_ok=True
)


# =====================================================
# 5. 搜尋 Excel 與 CSV 檔案
# =====================================================

supported_extensions = (
    ".xlsx",
    ".xls",
    ".csv"
)

data_files = sorted([
    f for f in os.listdir(folder_path)
    if (
        os.path.isfile(os.path.join(folder_path, f))
        and f.lower().endswith(supported_extensions)
        and not f.startswith("~$")
    )
])

if not data_files:
    print("❌ 資料夾內沒有 Excel 或 CSV 檔案")
    raise SystemExit


print("=" * 60)
print("🚢 AIS 船舶資料 MMSI 分類工具")
print("=" * 60)

print(f"📁 來源資料夾：{folder_path}")
print(f"📁 輸出資料夾：{output_base}")
print(f"📄 待處理檔案數：{len(data_files)}")

print("\n🚢 開始處理...\n")


# =====================================================
# 6. 統計變數
# =====================================================

success_files = 0
failed_files = 0
total_output = 0


# =====================================================
# 7. 逐一處理 Excel / CSV
# =====================================================

for file in data_files:

    file_path = os.path.join(
        folder_path,
        file
    )

    print("-" * 60)
    print(f"📄 讀取：{file}")

    # -------------------------------------------------
    # 7-1. 擷取檔名中間標籤
    # -------------------------------------------------

    # 範例：
    # D&TMOK HWLN_20250508_23.xlsx
    # D&TMOK HWLN_20250508_23.csv
    #
    # 取得：20250508

    filename_no_ext = os.path.splitext(file)[0]

    name_parts = filename_no_ext.split("_")

    if len(name_parts) >= 2:

        file_tag = name_parts[1]

    else:

        file_tag = filename_no_ext

        print(
            f"   ⚠ 檔名 {file} 無底線，"
            "使用原始檔名作為標籤"
        )

    # -------------------------------------------------
    # 7-2. 讀取資料
    # -------------------------------------------------

    try:

        df = read_ais_file(file_path)

        print(f"   ✔ 原始資料筆數：{len(df):,}")

    except Exception as e:

        print(f"   ❌ 讀取失敗：{e}")

        failed_files += 1

        continue

    # -------------------------------------------------
    # 7-3. 清理欄位名稱
    # -------------------------------------------------

    df.columns = (
        df.columns
        .astype(str)
        .str.strip()
        .str.replace("\ufeff", "", regex=False)
    )

    # -------------------------------------------------
    # 7-4. 欄位名稱標準化
    # -------------------------------------------------

    df = df.rename(
        columns={
            k: v
            for k, v in column_mapping.items()
            if k in df.columns
        }
    )

    # 若標準化後出現重複欄位，優先保留第一欄
    if df.columns.duplicated().any():

        print("   ⚠ 偵測到重複欄位，保留第一欄")

        df = df.loc[
            :,
            ~df.columns.duplicated()
        ].copy()

    # -------------------------------------------------
    # 7-5. 檢查 MMSI 欄位
    # -------------------------------------------------

    if "mmsi" not in df.columns:

        print(
            f"   ⚠ {file} 沒有 MMSI / SOURCE_ID 欄位"
        )

        failed_files += 1

        continue

    # -------------------------------------------------
    # 7-6. MMSI 資料清理
    # -------------------------------------------------

    df["mmsi"] = df["mmsi"].apply(
        clean_mmsi
    )

    original_count = len(df)

    df = df.dropna(
        subset=["mmsi"]
    ).copy()

    valid_count = len(df)

    invalid_count = original_count - valid_count

    print(f"   ✔ 有效 MMSI 資料：{valid_count:,}")
    print(f"   ✔ 排除無效資料：{invalid_count:,}")

    if df.empty:

        print("   ⚠ 沒有有效 MMSI，跳過")

        failed_files += 1

        continue


# -------------------------------------------------
# 7-7. MMSI 分類輸出（不建立 MMSI 子資料夾）
# -------------------------------------------------

mmsi_count = df["mmsi"].nunique()

print(f"   🚢 船舶 MMSI 數量：{mmsi_count:,}")

output_count = 0

for mmsi, df_single in df.groupby(
    "mmsi",
    sort=False
):

    # 輸出檔名：MMSI_日期.xlsx
    new_filename = f"{mmsi}_{file_tag}.xlsx"

    # 所有檔案直接存放在 shiptrack 資料夾
    output_file = os.path.join(
        output_base,
        new_filename
    )

    try:

        df_single.to_excel(
            output_file,
            index=False,
            engine="openpyxl"
        )

        output_count += 1
        total_output += 1

        print(
            f"   → {new_filename} "
            f"({len(df_single):,} 筆)"
        )

    except Exception as e:

        print(
            f"   ❌ MMSI {mmsi} 輸出失敗：{e}"
        )

if output_count == mmsi_count:
    success_files += 1
else:
    failed_files += 1


# =====================================================
# 8. 作業結果統計
# =====================================================

print("\n" + "=" * 60)
print("✅ AIS 船舶資料處理完成")
print("=" * 60)

print(f"📄 來源檔案總數：{len(data_files):,}")
print(f"✔ 完整處理檔案：{success_files:,}")
print(f"⚠ 異常或未完整處理：{failed_files:,}")
print(f"🚢 成功輸出檔案：{total_output:,}")

print(f"\n📁 輸出位置：{output_base}")

print("=" * 60)

input("\n按 Enter 結束程式...")
