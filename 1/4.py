import argparse
import csv

import xlsxwriter


parser = argparse.ArgumentParser()
parser.add_argument("input_file")
parser.add_argument("output_file")
parser.add_argument("-d", "--delimiter", default=";")
args = parser.parse_args()

with open(args.input_file, encoding="utf-8", newline="") as file:
    rows = csv.reader(file, delimiter=args.delimiter)

    workbook = xlsxwriter.Workbook(args.output_file)
    worksheet = workbook.add_worksheet()

    for row_number, row in enumerate(rows):
        worksheet.write_row(row_number, 0, row)

    workbook.close()
