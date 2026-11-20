# -*- coding: utf-8 -*-
"""Реестр автомобилей «Автомобили.xlsx»: валидация и безопасная дозапись строк.

Используется скриптом scan_cars.py, но может вызываться и отдельно:

    from car_registry import append_vehicle, validate_vehicle
"""
from __future__ import annotations

import os
import re
from copy import copy
from datetime import date, datetime

try:  # боту Excel не нужен: без openpyxl импорт модуля всё равно проходит
    from openpyxl import load_workbook
except ImportError:  # pragma: no cover
    load_workbook = None


def _require_openpyxl():
    if load_workbook is None:
        raise RuntimeError(
            "Для работы с Excel нужен пакет openpyxl: py -m pip install openpyxl"
        )

HEADERS = [
    "№",
    "Гос. номер",
    "VIN",
    "Марка",
    "Модель",
    "Тип ТС",
    "Год выпуска",
    "Владелец (по СТС)",
    "Заказчик",
    "СТС №",
    "Дата выдачи СТС",
    "Примечание",
]

VIN_ALLOWED = re.compile(r"^[A-HJ-NPR-Z0-9]{17}$")

# Похожие друг на друга буквы: всё приводим к латинице (решение пользователя).
CYR_TO_LAT = str.maketrans({
    "А": "A", "В": "B", "Е": "E", "К": "K", "М": "M", "Н": "H", "О": "O",
    "Р": "P", "С": "C", "Т": "T", "У": "Y", "Х": "X", "І": "I", "Ї": "I",
    "а": "a", "в": "b", "е": "e", "к": "k", "м": "m", "н": "h", "о": "o",
    "р": "p", "с": "c", "т": "t", "у": "y", "х": "x", "і": "i", "ї": "i",
})


def to_latin(text: str | None) -> str | None:
    """Переводит кириллические двойники латинских букв в латиницу."""
    if text is None:
        return None
    return str(text).translate(CYR_TO_LAT)


# Российские и белорусские номерные знаки — проверка мягкая, только предупреждение.
# Сравнение идёт по латинице: А123ВС77 -> A123BC77.
RU_PLATE = re.compile(
    r"^(?:[ABEKMHOPCTYX]\d{3}[ABEKMHOPCTYX]{2}\d{2,3}"     # A123BC77
    r"|[ABEKMHOPCTYX]{2}\d{4,5}\d{2,3}"                     # BP2655 77 (прицеп)
    r"|[ABEKMHOPCTYX]{1,2}\d{3,5})$"                        # прочие/транзит
)
BY_PLATE = re.compile(
    r"^(?:\d{4}\s?[ABEKMHOPCTYX]{2}-\d"                     # 1234 AB-5
    r"|[ABEKMHOPCTYX]\s?\d{4}\s?[ABEKMHOPCTYX]-\d)$"        # A 8003 B-3
)

# 10-й знак VIN -> модельный год. Кодировка повторяется каждые 30 лет,
# поэтому возвращаем оба варианта (1980-2009 и 2010-2039).
LETTER_YEARS = "ABCDEFGHJKLMNPRSTVWXY"  # без I, O, Q, U, Z


def model_year_candidates(vin: str) -> list[int]:
    if not vin or len(vin) != 17:
        return []
    ch = vin[9].upper()
    if ch in LETTER_YEARS:
        base = 1980 + LETTER_YEARS.index(ch)
    elif ch.isdigit() and ch != "0":
        base = 2000 + int(ch)
    else:
        return []
    return [base, base + 30]


def model_year_from_vin(vin: str, near: int | None = None):
    """Год по 10-му знаку VIN; при known `near` — ближайший прошедший кандидат."""
    candidates = model_year_candidates(vin)
    if not candidates:
        return None
    if near is None:
        return candidates[0]
    past = [y for y in candidates if y <= near + 1]
    if past:
        return max(past)
    return min(candidates, key=lambda y: abs(y - near))


class RegistryLocked(RuntimeError):
    """Файл реестра занят другим процессом (обычно открыт в Excel)."""


def ensure_unlocked(path: str) -> None:
    if not os.path.exists(path):
        return
    try:
        with open(path, "r+b"):
            pass
    except PermissionError as exc:
        raise RegistryLocked(
            f"Файл {os.path.basename(path)} занят (открыт в Excel или синхронизируется). "
            "Закройте его и повторите запуск."
        ) from exc


def normalize_plate(plate: str | None) -> str:
    """Сравнительный вид номера: латиница, без пробелов и дефисов."""
    if not plate:
        return ""
    return re.sub(r"[\s\-]", "", to_latin(str(plate))).upper()


def display_plate(plate: str | None, enabled: bool = False) -> str | None:
    """Номер для записи в таблицу: всегда латиница, пробелы/дефисы как в документе.

    С enabled=True дополнительно добавляет пробелы в прицепных номерах
    («BP265577» -> «BP 2655 77»).
    """
    text = clean_text(plate)
    if not text:
        return text
    text = to_latin(text)
    if not enabled:
        return text
    compact = normalize_plate(text)
    m = re.match(r"^([ABEKMHOPCTYX]{2})(\d{4})(\d{2,3})$", compact)   # прицеп: BP2655 77
    if m:
        return f"{m.group(1)} {m.group(2)} {m.group(3)}"
    return text


def clean_text(value):
    if value is None:
        return None
    text = str(value).strip()
    return text or None


def as_date(value):
    if value in (None, ""):
        return None
    if isinstance(value, (date, datetime)):
        return value.date() if isinstance(value, datetime) else value
    text = str(value).strip()
    for fmt in ("%d.%m.%Y", "%Y-%m-%d", "%d/%m/%Y"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return None


def normalize_caps(text):
    """«ПОЛУПРИЦЕП ЦИСТЕРНА» -> «Полуприцеп цистерна»."""
    text = clean_text(text)
    if not text:
        return None
    return text.capitalize() if text.isupper() else text


def validate_vehicle(vehicle: dict, existing: list[dict]) -> tuple[list[str], list[str]]:
    """Возвращает (ошибки, предупреждения). Ошибки блокируют запись."""
    errors, warnings = [], []

    vin = (to_latin(vehicle.get("vin")) or "").strip().upper()
    if not vin:
        errors.append("VIN не распознан")
    elif len(vin) != 17:
        errors.append(f"VIN {vin}: {len(vin)} знаков вместо 17")
    elif not VIN_ALLOWED.match(vin):
        bad = sorted({c for c in vin if not re.match(r"[A-HJ-NPR-Z0-9]", c)})
        errors.append(f"VIN {vin}: недопустимые символы {bad or 'I/O/Q'}")

    plate = clean_text(vehicle.get("plate"))
    norm_plate = normalize_plate(plate)
    if not plate:
        errors.append("гос. номер не распознан")
    elif norm_plate and not (RU_PLATE.match(norm_plate)
                             or BY_PLATE.match(to_latin(plate) or "")):
        warnings.append(f"номер «{to_latin(plate)}» не похож на стандартный РФ/РБ")

    year = vehicle.get("year")
    if year is None:
        warnings.append("год выпуска не распознан")
    elif not (1950 <= int(year) <= date.today().year + 1):
        errors.append(f"год выпуска {year} вне разумного диапазона")

    vin_year = model_year_from_vin(vin, near=int(year) if year else None)
    if year and vin_year and abs(int(year) - vin_year) > 1:
        warnings.append(f"год в документе {year}, а 10-й знак VIN даёт {vin_year}")

    for old in existing:
        if vin and normalize_plate(old.get("vin")) == normalize_plate(vin):
            errors.append(f"дубль VIN — уже есть строка №{old.get('row_no')} ({old.get('plate')})")
        elif norm_plate and normalize_plate(old.get("plate")) == norm_plate:
            errors.append(f"дубль номера — уже есть строка №{old.get('row_no')} (VIN {old.get('vin')})")

    if not vehicle.get("brand"):
        warnings.append("марка не распознана")
    if not vehicle.get("sts"):
        warnings.append("№ свидетельства не распознан")

    return errors, warnings


def read_registry(path: str) -> list[dict]:
    """Читает существующие строки реестра; занятый файл — ошибка RegistryLocked."""
    if not os.path.exists(path):
        return []
    _require_openpyxl()
    ensure_unlocked(path)
    wb = load_workbook(path, data_only=True)
    ws = wb[wb.sheetnames[0]]
    out = []
    for idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if not row or all(v is None for v in row):
            continue
        out.append(
            {
                "row_no": row[0],
                "plate": row[1],
                "vin": row[2],
                "brand": row[3],
                "model": row[4],
                "row_index": idx,
            }
        )
    return out


def build_note(vehicle: dict) -> str | None:
    parts = []
    if vehicle.get("doc_type"):
        parts.append(str(vehicle["doc_type"]))
    if vehicle.get("pts"):
        parts.append(f"ПТС {vehicle['pts']}")
    if vehicle.get("category") and str(vehicle["category"]).strip().lower() not in ("", "—"):
        parts.append(f"категория {vehicle['category']}")
    if vehicle.get("lessor"):
        parts.append(f"лизинг: лизингополучатель, лизингодатель {vehicle['lessor']}")
    if vehicle.get("max_mass"):
        parts.append(f"разреш. max масса {vehicle['max_mass']} кг")
    if vehicle.get("curb_mass"):
        parts.append(f"масса без нагрузки {vehicle['curb_mass']} кг")
    if vehicle.get("color"):
        parts.append(f"цвет {str(vehicle['color']).lower()}")
    if vehicle.get("notes"):
        parts.append(str(vehicle["notes"]))
    return "; ".join(parts) or None


def append_vehicle(path: str, vehicle: dict, dry_run: bool = False,
                   format_plates: bool = False) -> dict:
    """Добавляет строку в конец листа, копируя оформление последней строки.

    Занятый файл — ошибка RegistryLocked, никаких обходных копий не создаётся.
    """
    _require_openpyxl()
    ensure_unlocked(path)
    wb = load_workbook(path)
    ws = wb[wb.sheetnames[0]]

    template_row = max(ws.max_row, 2)
    new_row = ws.max_row + 1

    sts_date = as_date(vehicle.get("sts_date"))
    values = [
        new_row - 1,
        display_plate(vehicle.get("plate"), format_plates),
        clean_text((to_latin(vehicle.get("vin")) or "").upper()),
        clean_text(vehicle.get("brand")),
        clean_text(vehicle.get("model")),
        normalize_caps(vehicle.get("vehicle_type")),
        int(vehicle["year"]) if vehicle.get("year") else None,
        clean_text(vehicle.get("owner")),
        clean_text(vehicle.get("customer")),
        clean_text(vehicle.get("sts")),
        sts_date,
        build_note(vehicle),
    ]

    for col, val in enumerate(values, start=1):
        src = ws.cell(row=template_row, column=col)
        dst = ws.cell(row=new_row, column=col, value=val)
        dst.font = copy(src.font)
        dst.border = copy(src.border)
        dst.alignment = copy(src.alignment)
        dst.number_format = src.number_format

    ws.row_dimensions[new_row].height = 60
    ws.auto_filter.ref = f"A1:L{new_row}"

    if not dry_run:
        wb.save(path)
    return {"row": new_row, "values": values}
