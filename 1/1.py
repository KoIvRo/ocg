import random
from datetime import date
from pathlib import Path

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.shared import Mm
from docxtpl import DocxTemplate


COMPANIES = ["ООО Альфа", "ООО Вектор", "ИП Иванов И.И.", "ООО Меридиан"]
SELLERS = ["Иванов Иван Иванович", "Петров Петр Петрович", "Сидорова Анна Олеговна"]
ADDRESSES = [
    "г. Владивосток, ул. Светланская, 15",
    "г. Владивосток, ул. Русская, 42",
    "г. Артем, ул. Кирова, 8",
]
PRODUCTS = [
    ("Бумага офисная", "A001", "пач."),
    ("Ручка шариковая", "A002", "шт."),
    ("Карандаш", "A003", "шт."),
    ("Папка", "A004", "шт."),
    ("Блокнот", "A005", "шт."),
]
MONTHS = [
    "января", "февраля", "марта", "апреля", "мая", "июня",
    "июля", "августа", "сентября", "октября", "ноября", "декабря",
]


def generate_data(count=15):
    documents = []
    today = date.today()

    for _ in range(max(15, count)):
        product_rows = []
        for title, code, unit in random.sample(PRODUCTS, random.randint(2, 5)):
            amount = random.randint(1, 10)
            price = random.randint(50, 2000)
            product_rows.append(
                {
                    "title": title,
                    "code": code,
                    "unit": unit,
                    "amount": amount,
                    "price": f"{price:.2f}",
                    "sum": f"{amount * price:.2f}",
                }
            )

        documents.append(
            (
                random.choice(COMPANIES),
                random.randint(10000, 99999),
                today.day,
                MONTHS[today.month - 1],
                str(today.year)[-2:],
                random.choice(SELLERS),
                random.choice(ADDRESSES),
                random.randint(10**12, 10**13 - 1),
                product_rows,
                f"{sum(float(item['sum']) for item in product_rows):.2f}",
            )
        )

    return documents


def main():
    output_dir = Path("checks/1")
    output_dir.mkdir(parents=True, exist_ok=True)

    for index, item in enumerate(generate_data(), start=1):
        company, number, day, month, year, seller, address, ogrn, products, total = item
        document = DocxTemplate("tmp.docx")
        document.render(
            {
                "company": company,
                "check_number": number,
                "day": day,
                "month": month,
                "year": year,
                "seller": seller,
                "address": address,
                "ORGN": ogrn,
                "products": products,
                "general_sum": total,
            }
        )
        output_path = output_dir / f"check_{index}.docx"
        document.save(output_path)


if __name__ == "__main__":
    main()
