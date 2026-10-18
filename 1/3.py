import argparse
import random
import re
from datetime import date, timedelta
from pathlib import Path

from docxtpl import DocxTemplate


LAST_NAMES = ["Иванов", "Петров", "Сидоров", "Смирнов", "Кузнецов", "Попов", "Волков", "Орлов"]
FIRST_NAMES = ["Иван", "Петр", "Алексей", "Дмитрий", "Анна", "Мария", "Елена", "Ольга"]
PATRONYMICS = ["Иванович", "Петрович", "Алексеевич", "Сергеевич", "Ивановна", "Петровна", "Алексеевна", "Сергеевна"]
ADDRESSES = [
    "г. Владивосток, ул. Светланская, 10",
    "г. Владивосток, ул. Русская, 25",
    "г. Артем, ул. Кирова, 17",
]
PLACES = ["ресторан Морской", "банкетный зал Восток", "усадьба Белый сад"]


def generate_guests(n):
    guests = []
    used_names = set()

    while len(guests) < n:
        full_name = f"{random.choice(LAST_NAMES)} {random.choice(FIRST_NAMES)} {random.choice(PATRONYMICS)}"
        if full_name in used_names:
            continue
        used_names.add(full_name)
        event_date = date.today() + timedelta(days=random.randint(30, 180))
        guests.append(
            (
                full_name,
                random.choice(ADDRESSES),
                event_date.strftime("%d.%m.%Y"),
                f"{random.randint(15, 19):02d}:00",
                random.choice(PLACES),
            )
        )

    return guests


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("n", type=int)
    args = parser.parse_args()
    if not 1 <= args.n <= len(LAST_NAMES) * len(FIRST_NAMES) * len(PATRONYMICS):
        parser.error("n должно быть от 1 до 512")

    output_dir = Path("checks/3")
    output_dir.mkdir(parents=True, exist_ok=True)

    for full_name, address, event_date, event_time, place in generate_guests(args.n):
        document = DocxTemplate("checks/2/invitation_template.docx")
        document.render(
            {
                "full_name": full_name,
                "address": address,
                "date": event_date,
                "time": event_time,
                "place": place,
            }
        )
        filename = re.sub(r'[\\/:*?"<>|]', "_", full_name)
        document.save(output_dir / f"{filename}.docx")


if __name__ == "__main__":
    main()
