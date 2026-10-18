import argparse
import csv
import re
from collections import Counter, defaultdict
from pathlib import Path

import xlsxwriter


def to_number(value):
    try:
        return float(value.strip().replace(",", "."))
    except ValueError:
        return 0.0


def final_grade(rating):
    if rating >= 80:
        return 5
    if rating >= 70:
        return 4
    if rating >= 50:
        return 3
    return 2


def sheet_name(group, used_names):
    name = re.sub(r"[\[\]:*?/\\]", "_", group)[:31] or "Группа"
    original = name
    number = 2
    while name in used_names:
        suffix = f"_{number}"
        name = original[: 31 - len(suffix)] + suffix
        number += 1
    used_names.add(name)
    return name


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("input_file", nargs="?", default="checks/5/students.csv")
    parser.add_argument("output_file", nargs="?", default="checks/5/grades.xlsx")
    parser.add_argument("-d", "--delimiter", default=";")
    args = parser.parse_args()

    with open(args.input_file, encoding="utf-8-sig", newline="") as source:
        reader = csv.reader(source, delimiter=args.delimiter)
        header = next(reader)
        groups = defaultdict(list)
        for row in reader:
            if len(row) >= 2:
                row += [""] * (len(header) - len(row))
                groups[row[1]].append(row[: len(header)])

    Path(args.output_file).parent.mkdir(parents=True, exist_ok=True)
    workbook = xlsxwriter.Workbook(args.output_file)
    header_format = workbook.add_format({"bold": True, "bg_color": "#D9EAF7", "border": 1})
    percent_format = workbook.add_format({"num_format": "0.00", "border": 1})
    cell_format = workbook.add_format({"border": 1})
    used_names = set()

    for group, students in groups.items():
        worksheet = workbook.add_worksheet(sheet_name(group, used_names))
        assignment_names = header[2:]
        result_header = header + ["Средний рейтинг", "Оценка"]
        worksheet.write_row(0, 0, result_header, header_format)

        grades = []
        assignment_totals = [0.0] * len(assignment_names)

        for row_index, row in enumerate(students, start=1):
            scores = [to_number(value) for value in row[2:]]
            rating = sum(scores) / len(scores) if scores else 0.0
            grade = final_grade(rating)
            grades.append(grade)
            for index, score in enumerate(scores):
                assignment_totals[index] += score

            worksheet.write(row_index, 0, row[0], cell_format)
            worksheet.write(row_index, 1, row[1], cell_format)
            for column, score in enumerate(scores, start=2):
                worksheet.write_number(row_index, column, score, cell_format)
            worksheet.write_number(row_index, len(header), rating, percent_format)
            worksheet.write_number(row_index, len(header) + 1, grade, cell_format)

        summary_row = len(students) + 3
        worksheet.write_row(summary_row, 0, ["Оценка", "Количество"], header_format)
        grade_counts = Counter(grades)
        for offset, grade in enumerate((2, 3, 4, 5), start=1):
            worksheet.write_number(summary_row + offset, 0, grade, cell_format)
            worksheet.write_number(summary_row + offset, 1, grade_counts[grade], cell_format)

        assignment_row = summary_row + 7
        worksheet.write_row(assignment_row, 0, ["Задание", "Средний результат"], header_format)
        for offset, (name, total) in enumerate(zip(assignment_names, assignment_totals), start=1):
            worksheet.write(assignment_row + offset, 0, name, cell_format)
            average = total / len(students) if students else 0.0
            worksheet.write_number(assignment_row + offset, 1, average, percent_format)

        pie_chart = workbook.add_chart({"type": "pie"})
        pie_chart.add_series(
            {
                "name": "Распределение оценок",
                "categories": [worksheet.name, summary_row + 1, 0, summary_row + 4, 0],
                "values": [worksheet.name, summary_row + 1, 1, summary_row + 4, 1],
                "data_labels": {"percentage": True},
            }
        )
        pie_chart.set_title({"name": "Оценки группы"})

        column_chart = workbook.add_chart({"type": "column"})
        if assignment_names:
            column_chart.add_series(
                {
                    "name": "Средняя успеваемость",
                    "categories": [worksheet.name, assignment_row + 1, 0, assignment_row + len(assignment_names), 0],
                    "values": [worksheet.name, assignment_row + 1, 1, assignment_row + len(assignment_names), 1],
                }
            )
        column_chart.set_title({"name": "Успеваемость по заданиям"})
        column_chart.set_y_axis({"min": 0, "max": 100})

        worksheet.insert_chart(1, len(header) + 3, pie_chart)
        worksheet.insert_chart(17, len(header) + 3, column_chart)
        worksheet.set_column(0, 0, 30)
        worksheet.set_column(1, len(result_header) - 1, 16)
        worksheet.set_landscape()
        worksheet.fit_to_pages(1, 1)

    workbook.close()


if __name__ == "__main__":
    main()
