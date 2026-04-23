import csv
import os
from collections import defaultdict


def split_csv_by_month(input_file: str, date_column: str = "date", output_dir: str = "output"):
    """
    CSVファイルを月別に分割して保存する。
    date_column で指定した列（YYYY-MM-DD 形式）を基準にする。
    """
    os.makedirs(output_dir, exist_ok=True)

    with open(input_file, newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        if date_column not in reader.fieldnames:
            raise ValueError(f"列 '{date_column}' が見つかりません。利用可能な列: {reader.fieldnames}")

        rows_by_month = defaultdict(list)
        for row in reader:
            month_key = row[date_column][:7]  # "YYYY-MM"
            rows_by_month[month_key].append(row)

    fieldnames = None
    with open(input_file, newline="", encoding="utf-8") as f:
        fieldnames = csv.DictReader(f).fieldnames

    created_files = []
    for month_key, rows in sorted(rows_by_month.items()):
        output_file = os.path.join(output_dir, f"{month_key}.csv")
        with open(output_file, "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(rows)
        created_files.append((output_file, len(rows)))

    return created_files


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="CSVファイルを月別に分割します")
    parser.add_argument("input", help="入力CSVファイルのパス")
    parser.add_argument("--date-column", default="date", help="日付列の名前 (デフォルト: date)")
    parser.add_argument("--output-dir", default="output", help="出力ディレクトリ (デフォルト: output)")
    args = parser.parse_args()

    files = split_csv_by_month(args.input, args.date_column, args.output_dir)
    for path, count in files:
        print(f"{path}: {count} 件")
