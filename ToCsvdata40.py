
# -*- coding: utf-8 -*-

r"""
AIS 解碼資料 40 欄位標準化工具

功能：
1. 視窗選取資料夾
2. 批次讀取 Excel (.xlsx/.xls) 與 CSV
3. 自動對應新舊欄位名稱
4. 固定輸出 40 個欄位
5. 保留 accuracy 數字、文字與空白
6. 缺少欄位自動補空白
7. CSV 自動辨識常見編碼及分隔符
8. 不修改原始檔案
9. 輸出 UTF-8 BOM CSV
10. 產生批次處理報告

張耘馨
原版：2025.08.24
修正版：2026.11.04

執行方式：
python AIS_40Columns_ExcelCSV_Converter.py

需安裝：
pip install openpyxl pandas xlrd

DB資料類型
CREATE TABLE ship_logTT ( 

  

    GPS_year   VARCHAR(8), 

    GPS_month VARCHAR(4), 

    GPS_day VARCHAR(4), 

    GPS_hour VARCHAR(4), 

    GPS_minute VARCHAR(4), 

    GPS_second VARCHAR(4),  

  PACKET_TYPE VARCHAR(10), 

    CHANNEL VARCHAR(2),  

    msg_type TEXT, 

    mmsi TEXT, 

    shipname VARCHAR(100), 

    callsign VARCHAR(20), 

    imo TEXT, 

    ship_type TEXT,  

    TO_BOW DOUBLE PRECISION, 

    TO_STERN DOUBLE PRECISION, 

    TO_PORT DOUBLE PRECISION, 

    TO_STARBOARD DOUBLE PRECISION,  

    VESSEL_LENGTH DOUBLE PRECISION, 

    BEAM DOUBLE PRECISION,  

    lon DOUBLE PRECISION, 

    LONGITUDE_DESC VARCHAR(20), 

    lat DOUBLE PRECISION, 

    LATITUDE_DESC VARCHAR(20),  

    accuracy TEXT, 

    course DOUBLE PRECISION, 

    speed DOUBLE PRECISION, 

    heading TEXT, 

    ROT TEXT,  

    eta_datetime TEXT, 

    draught DOUBLE PRECISION, 

    destination VARCHAR(100),  

    SENTENCE VARCHAR(20), 

    TALKER VARCHAR(20),  

    status TEXT,  

    bearing DOUBLE PRECISION, 

    distance DOUBLE PRECISION, 

    "distance in nautical miles" DOUBLE PRECISION,  

    "基點lon" DOUBLE PRECISION DEFAULT 121.7776, 

    "基點lat" DOUBLE PRECISION DEFAULT 25.14993  

); 

 

"""

import csv
import math
import re
from pathlib import Path
from collections import Counter
from datetime import date, datetime

import tkinter as tk
from tkinter import filedialog, messagebox

from openpyxl import load_workbook


# ==================================================
# 1. 基本設定
# ==================================================

OUTPUT_FOLDER = "CSV_40欄位_標準化"

# True：只要某檔缺少欄位，就不匯出該檔
# False：缺少欄位自動補空白
STRICT_COLUMNS = False

# CSV 資料列欄位數不符合表頭時：
# True：不匯出該檔案
# False：跳過異常列並記錄
STRICT_ROWS = True

# 基點座標預設留空。
# 若整批資料確定使用同一基點，再填入實際數值。
BASE_LON = ""
BASE_LAT = ""

# ==================================================
# 2. 定義標準 40 欄位
# ==================================================

COLUMNS = [
    "GPS_year",
    "GPS_month",
    "GPS_day",
    "GPS_hour",
    "GPS_minute",
    "GPS_second",
    "PACKET_TYPE",
    "CHANNEL",
    "msg_type",
    "mmsi",
    "shipname",
    "callsign",
    "imo",
    "ship_type",
    "TO_BOW",
    "TO_STERN",
    "TO_PORT",
    "TO_STARBOARD",
    "VESSEL_LENGTH",
    "BEAM",
    "lon",
    "LONGITUDE_DESC",
    "lat",
    "LATITUDE_DESC",
    "accuracy",
    "course",
    "speed",
    "heading",
    "ROT",
    "eta_datetime",
    "draught",
    "destination",
    "SENTENCE|SENTENCE",
    "TALKER|TALKER",
    "status",
    "bearing",
    "distance",
    "distance in nautical miles",
    "基點lon",
    "基點lat"
]

assert len(COLUMNS) == 40


# ==================================================
# 3. 新舊欄位名稱對照
# ==================================================

ALIASES = {
    "GPS_year": ["Year", "rmc_date_year"],
    "GPS_month": ["Month", "rmc_date_mon"],
    "GPS_day": ["Day", "rmc_date_day"],
    "GPS_hour": ["Hour", "rmc_taken_at_hour"],
    "GPS_minute": ["Minute", "rmc_taken_at_min"],
    "GPS_second": ["Second", "rmc_taken_at_sec"],

    "PACKET_TYPE": ["packet_type"],
    "CHANNEL": ["channel"],

    "msg_type": ["message_id"],
    "mmsi": ["source_id"],
    "shipname": ["name"],
    "callsign": ["call_sign"],
    "imo": ["imo_number"],
    "ship_type": ["type_of_ship_and_cargo_type"],

    "TO_BOW": [
        "to_bow",
        "dimension_reference_for_position_a"
    ],
    "TO_STERN": [
        "to_stern",
        "dimension_reference_for_position_b"
    ],
    "TO_PORT": [
        "to_port",
        "dimension_reference_for_position_c"
    ],
    "TO_STARBOARD": [
        "to_starboard",
        "dimension_reference_for_position_d"
    ],

    "VESSEL_LENGTH": ["LENGTH", "vessel_length"],
    "BEAM": ["beam"],

    "lon": ["longitude"],
    "LONGITUDE_DESC": ["longitude_desc"],
    "lat": ["latitude"],
    "LATITUDE_DESC": ["latitude_desc"],

    "accuracy": ["position_accuracy"],
    "course": ["course_over_ground"],
    "speed": ["speed_over_ground"],
    "heading": ["true_heading"],
    "ROT": ["rate_of_turn"],

    "eta_datetime": ["eta"],
    "draught": ["maximum_present_static_draught"],
    "destination": [],

    "SENTENCE|SENTENCE": ["SENTENCE", "sentence"],
    "TALKER|TALKER": ["TALKER", "talker"],

    "status": ["navigational_status"],
    "bearing": [],
    "distance": [],
    "distance in nautical miles": [
        "distance_in_nautical_miles",
        "distance_nm"
    ],

    "基點lon": ["base_lon"],
    "基點lat": ["base_lat"]
}


# ==================================================
# 4. 欄位名稱標準化
# ==================================================

def normalize_name(value):
    if value is None:
        return ""

    name = str(value).strip()

    # 移除 UTF-8 BOM
    name = name.replace("\ufeff", "")

    # 空白、底線、大小寫統一
    name = re.sub(r"[\s_]+", "", name)

    return name.casefold()


def build_alias_map():
    result = {}

    for standard_name in COLUMNS:
        names = [standard_name]
        names.extend(ALIASES.get(standard_name, []))

        for name in names:
            key = normalize_name(name)
            result[key] = standard_name

    return result


ALIAS_MAP = build_alias_map()


def create_mapping(headers):
    mapping = {}
    extra = []
    duplicate = []

    for index, header in enumerate(headers):
        key = normalize_name(header)
        standard = ALIAS_MAP.get(key)

        if standard is None:
            extra.append(str(header))
            continue

        if standard in mapping:
            duplicate.append(str(header))
            continue

        mapping[standard] = index

    missing = [
        col for col in COLUMNS
        if col not in mapping
    ]

    return mapping, missing, extra, duplicate


# ==================================================
# 5. 數值轉換為 CSV 文字
# ==================================================

def to_text(value):

    if value is None:
        return ""

    if isinstance(value, str):
        # 保留文字與空白內容
        return value

    if isinstance(value, float):
        if math.isnan(value):
            return ""

        if math.isfinite(value) and value.is_integer():
            return str(int(value))

        return str(value)

    if isinstance(value, (datetime, date)):
        return value.isoformat(sep=" ") if isinstance(
            value, datetime
        ) else value.isoformat()

    return str(value)


# ==================================================
# 6. CSV 編碼偵測
# ==================================================

def detect_encoding(file_path):

    with open(file_path, "rb") as f:
        sample = f.read(65536)

    if sample.startswith(b"\xff\xfe"):
        return "utf-16"

    if sample.startswith(b"\xfe\xff"):
        return "utf-16"

    if sample.startswith(b"\xef\xbb\xbf"):
        return "utf-8-sig"

    for encoding in ["utf-8", "cp950", "big5"]:
        try:
            sample.decode(encoding)
            return encoding
        except UnicodeDecodeError as e:
            # 取樣可能剛好切到多位元組字元
            if (
                encoding == "utf-8"
                and e.start >= len(sample) - 4
            ):
                return "utf-8"
            continue

    raise ValueError("無法辨識 CSV 文字編碼")


# ==================================================
# 7. CSV 分隔符號偵測
# ==================================================

def detect_delimiter(file_path, encoding):

    with open(
        file_path,
        "r",
        encoding=encoding,
        newline=""
    ) as f:
        sample = f.read(8192)

    try:
        dialect = csv.Sniffer().sniff(
            sample,
            delimiters=",;\t|"
        )
        return dialect.delimiter
    except csv.Error:
        return ","


# ==================================================
# 8. CSV 讀取
# ==================================================

def read_csv_rows(file_path):

    encoding = detect_encoding(file_path)
    delimiter = detect_delimiter(
        file_path, encoding
    )

    with open(
        file_path,
        "r",
        encoding=encoding,
        newline=""
    ) as f:

        reader = csv.reader(
            f,
            delimiter=delimiter
        )

        headers = next(reader)

        # 將檔案保持開啟，逐列提供資料
        yield headers, None

        for row in reader:
            yield None, row


# ==================================================
# 9. Excel 讀取
# ==================================================

def read_excel_rows(file_path):

    suffix = file_path.suffix.lower()

    if suffix == ".xlsx":

        wb = load_workbook(
            file_path,
            read_only=True,
            data_only=True
        )

        try:
            ws = wb.active
            rows = ws.iter_rows(values_only=True)

            headers = next(rows)

            yield list(headers), None

            for row in rows:
                yield None, list(row)

        finally:
            wb.close()

    elif suffix == ".xls":

        import pandas as pd

        df = pd.read_excel(
            file_path,
            dtype=object,
            keep_default_na=False,
            engine="xlrd"
        )

        yield list(df.columns), None

        for row in df.itertuples(
            index=False,
            name=None
        ):
            yield None, list(row)


# ==================================================
# 10. 將資料列轉成標準 40 欄位
# ==================================================

def standardize_row(row, mapping):

    result = []

    for column in COLUMNS:

        if column in mapping:
            index = mapping[column]
            value = (
                row[index]
                if index < len(row)
                else ""
            )
        else:
            value = ""

        value = to_text(value)

        # 基點座標：只補缺少的值
        if column == "基點lon" and value == "":
            value = BASE_LON

        if column == "基點lat" and value == "":
            value = BASE_LAT

        result.append(value)

    return result


# ==================================================
# 11. 處理單一檔案
# ==================================================

def process_file(file_path, output_dir):

    output_path = output_dir / (
        file_path.stem + "_40cols.csv"
    )

    # 防止同名 Excel / CSV 互相覆蓋
    if output_path.exists():
        output_path = output_dir / (
            file_path.stem
            + "_"
            + file_path.suffix[1:].lower()
            + "_40cols.csv"
        )

    if output_path.exists():
        raise FileExistsError(
            f"輸出檔案已存在：{output_path.name}"
        )

    print("\n" + "=" * 60)
    print("處理檔案：", file_path.name)

    if file_path.suffix.lower() == ".csv":
        data = read_csv_rows(file_path)
    else:
        data = read_excel_rows(file_path)

    headers, _ = next(data)

    mapping, missing, extra, duplicate = (
        create_mapping(headers)
    )

    print("原始欄位數：", len(headers))
    print("標準欄位數：", len(COLUMNS))

    if missing:
        print("缺少欄位：", missing)

    if extra:
        print("多餘欄位：", extra)

    if duplicate:
        print("重複對應欄位：", duplicate)

    if STRICT_COLUMNS and missing:
        raise ValueError(
            "缺少必要欄位：" + ", ".join(missing)
        )

    # 建立輸出 CSV
    total = 0
    written = 0
    abnormal = 0
    accuracy_types = Counter()
    errors = []

    try:

        with open(
            output_path,
            "w",
            encoding="utf-8-sig",
            newline=""
        ) as f:

            writer = csv.writer(f)

            # 固定輸出標準 40 欄表頭
            writer.writerow(COLUMNS)

            for line_number, (_, row) in enumerate(
                data, start=2
            ):

                total += 1

                # 空列不匯入
                if all(
                    to_text(v).strip() == ""
                    for v in row
                ):
                    continue

                if len(row) != len(headers):

                    abnormal += 1

                    errors.append(
                        f"第{line_number}列："
                        f"預期{len(headers)}欄，"
                        f"實際{len(row)}欄"
                    )

                    if STRICT_ROWS:
                        continue

                    # 非嚴格模式也跳過異常列
                    continue

                new_row = standardize_row(
                    row, mapping
                )

                # 檢查 accuracy 原始內容
                accuracy_index = COLUMNS.index(
                    "accuracy"
                )

                accuracy_value = new_row[
                    accuracy_index
                ]

                if accuracy_value.strip() == "":
                    accuracy_types["空白"] += 1

                else:
                    try:
                        float(accuracy_value.strip())
                        accuracy_types["數字"] += 1
                    except ValueError:
                        accuracy_types["文字"] += 1

                writer.writerow(new_row)
                written += 1

        if STRICT_ROWS and abnormal > 0:
            output_path.unlink(missing_ok=True)

            raise ValueError(
                f"發現 {abnormal} 列欄位數異常，"
                "已取消此檔案匯出。"
                + "；".join(errors[:5])
            )

    except Exception:
        output_path.unlink(missing_ok=True)
        raise

    print("讀取資料列：", total)
    print("成功輸出列：", written)
    print("異常資料列：", abnormal)
    print("accuracy 統計：", dict(accuracy_types))
    print("輸出檔案：", output_path.name)

    return {
        "file": file_path.name,
        "output": output_path.name,
        "rows": written,
        "missing": missing,
        "extra": extra,
        "abnormal": abnormal
    }


# ==================================================
# 12. 主程式：視窗選擇資料夾
# ==================================================

def main():

    root = tk.Tk()
    root.withdraw()

    folder = filedialog.askdirectory(
        title="選擇 AIS Excel / CSV 資料夾"
    )

    if not folder:
        print("未選擇資料夾，程式結束")
        root.destroy()
        return

    folder_path = Path(folder)

    output_dir = (
        folder_path / OUTPUT_FOLDER
    )

    output_dir.mkdir(
        parents=True,
        exist_ok=True
    )

    files = sorted([
        p for p in folder_path.iterdir()
        if p.is_file()
        and p.suffix.lower()
        in [".xlsx", ".xls", ".csv"]
        and not p.name.startswith("~$")
    ])

    if not files:
        messagebox.showwarning(
            "沒有檔案",
            "找不到 Excel 或 CSV 檔案"
        )
        root.destroy()
        return

    success = 0
    failed = 0
    reports = []

    for file_path in files:

        try:
            result = process_file(
                file_path, output_dir
            )

            success += 1

            reports.append([
                file_path.name,
                "成功",
                result["rows"],
                result["abnormal"],
                ",".join(result["missing"]),
                ",".join(result["extra"]),
                ""
            ])

        except Exception as e:

            failed += 1

            print(
                "[FAILED]",
                file_path.name,
                str(e)
            )

            reports.append([
                file_path.name,
                "失敗",
                0,
                0,
                "",
                "",
                str(e)
            ])

    # 匯出處理報告
    report_path = (
        output_dir / "processing_report.csv"
    )

    with open(
        report_path,
        "w",
        encoding="utf-8-sig",
        newline=""
    ) as f:

        writer = csv.writer(f)

        writer.writerow([
            "原始檔名",
            "處理狀態",
            "輸出筆數",
            "異常列數",
            "缺少欄位",
            "多餘欄位",
            "錯誤訊息"
        ])

        writer.writerows(reports)

    print("\n" + "=" * 60)
    print("AIS 40欄位轉換完成")
    print("成功檔案數：", success)
    print("失敗檔案數：", failed)
    print("輸出資料夾：", output_dir)
    print("=" * 60)

    messagebox.showinfo(
        "AIS 資料轉換完成",
        f"成功：{success} 個檔案\n"
        f"失敗：{failed} 個檔案\n\n"
        f"輸出資料夾：\n{output_dir}"
    )

    root.destroy()


if __name__ == "__main__":
    main()
