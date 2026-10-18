import argparse
import csv
from pathlib import Path

import xlsxwriter


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("input_file", nargs="?", default="checks/4/input.csv")
    parser.add_argument("output_file", nargs="?", default="checks/4/output.xlsx")
    parser.add_argument("-d", "--delimiter", default=";")
    args = parser.parse_args()

    Path(args.output_file).parent.mkdir(parents=True, exist_ok=True)
    workbook = xlsxwriter.Workbook(args.output_file, {"strings_to_numbers": True})
    worksheet = workbook.add_worksheet()

    with open(args.input_file, encoding="utf-8-sig", newline="") as source:
        rows = list(csv.reader(source, delimiter=args.delimiter))
        for row_number, row in enumerate(rows):
            worksheet.write_row(row_number, 0, row)

    for column in range(max(map(len, rows), default=0)):
        width = max((len(row[column]) for row in rows if column < len(row)), default=0)
        worksheet.set_column(column, column, min(width + 2, 22))
    worksheet.set_landscape()
    worksheet.fit_to_pages(1, 1)

    workbook.close()


if __name__ == "__main__":
    main()
