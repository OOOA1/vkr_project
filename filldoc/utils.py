from typing import Any, Dict, Optional
import re
from dateutil import parser as dateparser
import datetime as _dt

# --- утилиты текста/якорей ---

def normalize(s: str) -> str:
    """Сжать повторяющиеся пробелы и обрезать края."""
    if s is None:
        return ""
    return re.sub(r"\s+", " ", str(s)).strip()

def parse_target_step(target: Optional[str]) -> tuple[str, int]:
    """'right(1)' -> ('right', 1)"""
    if not target:
        return ("right", 1)
    m = re.match(r"([a-zA-Z]+)(?:\((\d+)\))?$", target.strip())
    if not m:
        return ("right", 1)
    return (m.group(1).lower(), int(m.group(2) or "1"))

# --- трансформации значений ---

def fio_to_initials_surname(s: str) -> str:
    """
    'Иванов Иван Иванович' -> 'И.И. Иванов'
    'Иванов Иван'          -> 'И. Иванов'
    """
    parts = [p for p in re.split(r"\s+", s.strip()) if p]
    if not parts:
        return s
    if len(parts) == 1:
        return parts[0]
    if len(parts) == 2:
        surname, name = parts[0], parts[1]
        ini = (name[0] + ".") if name else ""
        return f"{ini} {surname}".strip()
    surname, name, patron = parts[0], parts[1], parts[2]
    ini = ""
    if name:
        ini += name[0] + "."
    if patron:
        ini += patron[0] + "."
    return f"{ini} {surname}".strip()

def transform_value(value: Any, transform: Optional[str]) -> str:
    # всё пустое и "none"/"nan" считаем пустым
    if value is None:
        return ""
    s = str(value)
    if s.strip().lower() in {"none", "nan"}:
        return ""

    if not transform:
        return s
    t = transform.strip()
    if t == "trim":
        return s.strip()
    if t == "UPPER":
        return s.upper()
    if t == "lower":
        return s.lower()
    if t == "Title":
        return s.title()
    if t == "FIO_INITIALS_SURNAME":
        return fio_to_initials_surname(s)
    if t.startswith("date:"):
        fmt = t.split(":", 1)[1]
        # дата из Excel приходит как datetime.date/datetime.datetime -> форматируем напрямую
        if isinstance(value, (_dt.date, _dt.datetime)):
            return value.strftime(fmt)
        try:
            dt = dateparser.parse(s, dayfirst=True, fuzzy=True)
            return dt.strftime(fmt)
        except Exception:
            return s
    return s

# --- политика длины для подчёркиваний ---

def apply_length_policy(src: str, replacement: str, policy: str) -> str:
    """
    src — исходный фрагмент (например, '__________'), replacement — что пишем.
    policy:
      - 'underline_and_keep_line' — вписываем в начало, оставшиеся '_' сохраняем.
      - 'replace_line' — заменяем полностью на replacement.
    """
    if policy == "replace_line":
        return replacement
    # underline_and_keep_line
    if len(replacement) >= len(src):
        re
