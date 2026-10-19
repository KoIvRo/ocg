import csv
from collections import Counter, defaultdict

import xlsxwriter


def get_grade(rating):
    if rating >= 80:
        return 5
    if rating >= 70:
        return 4
    if rating >= 50:
        return 3
    return 2


groups = defaultdict(list)

with open("checks/5/test_table.csv", encoding="utf-8", newline="") as file:
    for row in csv.reader(file, delimiter=","):
        group = row[0]
        name = row[1]
        results = [float(value.replace(",", ".")) for value in row[2:]]
        groups[group].append((name, results))

workbook = xlsxwriter.Workbook("checks/5/grades.xlsx")

for group, students in groups.items():
    worksheet = workbook.add_worksheet(group)
    block_count = len(students[0][1])
    headings = ["Группа", "ФИО"] + [f"Блок {number}" for number in range(1, block_count + 1)]
    headings += ["Средний рейтинг", "Оценка"]
    worksheet.write_row(0, 0, headings)

    grades = []

    for row_number, (name, results) in enumerate(students, start=1):
        rating = sum(results) / len(results)
        grade = get_grade(rating)
        grades.append(grade)

        worksheet.write(row_number, 0, group)
        worksheet.write(row_number, 1, name)
        worksheet.write_row(row_number, 2, results)
        worksheet.write(row_number, block_count + 2, rating)
        worksheet.write(row_number, block_count + 3, grade)

    summary_row = len(students) + 2
    worksheet.write_row(summary_row, 0, ["Оценка", "Количество"])
    grade_counts = Counter(grades)

    for number, grade in enumerate([2, 3, 4, 5], start=1):
        worksheet.write_row(summary_row + number, 0, [grade, grade_counts[grade]])

    averages_row = summary_row + 6
    worksheet.write_row(averages_row, 0, ["Задание", "Средняя успеваемость"])

    for block in range(block_count):
        average = sum(student[1][block] for student in students) / len(students)
        worksheet.write_row(averages_row + block + 1, 0, [f"Блок {block + 1}", average])

    pie_chart = workbook.add_chart({"type": "pie"})
    pie_chart.add_series({
        "categories": [group, summary_row + 1, 0, summary_row + 4, 0],
        "values": [group, summary_row + 1, 1, summary_row + 4, 1],
    })
    worksheet.insert_chart(1, block_count + 5, pie_chart)

    column_chart = workbook.add_chart({"type": "column"})
    column_chart.add_series({
        "categories": [group, averages_row + 1, 0, averages_row + block_count, 0],
        "values": [group, averages_row + 1, 1, averages_row + block_count, 1],
    })
    worksheet.insert_chart(17, block_count + 5, column_chart)

workbook.close()
