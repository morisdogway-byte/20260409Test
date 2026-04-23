import csv
import os
import re
from collections import defaultdict


ERA_OFFSETS = {"明治": 1867, "大正": 1911, "昭和": 1925, "平成": 1988, "令和": 2018}

JAPANESE_DATE_RE = re.compile(r"([^\d]+)(\d+)年(\d+)月")


def detect_encoding(file_path: str) -> str:
    for enc in ("utf-8-sig", "cp932", "utf-8"):
        try:
            with open(file_path, encoding=enc) as f:
                f.read()
            return enc
        except UnicodeDecodeError:
            continue
    return "utf-8"


def month_key_from_value(value: str) -> str | None:
    """日付文字列から 'YYYY-MM' キーを返す。和暦・西暦両対応。"""
    m = JAPANESE_DATE_RE.search(value)
    if m:
        era, year, month = m.group(1), int(m.group(2)), int(m.group(3))
        offset = ERA_OFFSETS.get(era)
        if offset is None:
            return None
        return f"{offset + year:04d}-{month:02d}"
    # YYYY-MM-DD 形式
    if re.match(r"\d{4}-\d{2}", value):
        return value[:7]
    return None


def split_csv_by_month(input_file: str, date_column: str = "日付", output_dir: str = "output"):
    encoding = detect_encoding(input_file)
    os.makedirs(output_dir, exist_ok=True)

    rows_by_month: dict[str, list] = defaultdict(list)
    fieldnames = None

    with open(input_file, newline="", encoding=encoding) as f:
        reader = csv.DictReader(f)
        fieldnames = reader.fieldnames
        if date_column not in fieldnames:
            raise ValueError(f"列 '{date_column}' が見つかりません。利用可能な列: {fieldnames}")
        for row in reader:
            key = month_key_from_value(row[date_column])
            if key:
                rows_by_month[key].append(row)

    created_files = []
    for key, rows in sorted(rows_by_month.items()):
        output_file = os.path.join(output_dir, f"{key}.csv")
        with open(output_file, "w", newline="", encoding=encoding) as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(rows)
        created_files.append((output_file, len(rows)))

    return created_files


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="CSVファイルを月別に分割します（和暦・西暦対応）")
    parser.add_argument("input", help="入力CSVファイルのパス")
    parser.add_argument("--date-column", default="日付", help="日付列の名前 (デフォルト: 日付)")
    parser.add_argument("--output-dir", default="output", help="出力ディレクトリ (デフォルト: output)")
    args = parser.parse_args()

    files = split_csv_by_month(args.input, args.date_column, args.output_dir)
    for path, count in files:
        print(f"{path}: {count} 件")
