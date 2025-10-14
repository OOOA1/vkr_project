from typing import Any, Dict, Optional
import re
from dateutil import parser as dateparser
import datetime as _dt
import os, shutil, subprocess, tempfile, pathlib, sys

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

def _which(*names):
    for n in names:
        if not n: 
            continue
        p = shutil.which(n)
        if p:
            return p
    return None

def normalize(s: str) -> str:
    if s is None: return ""
    return re.sub(r"\s+", " ", str(s)).strip()

def find_soffice() -> str | None:
    # приоритет: переменные окружения → типовые пути → PATH
    env = os.environ.get("SOFFICE_PATH") or os.environ.get("LIBREOFFICE_PATH")
    candidates = [
        env,
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        "/usr/bin/soffice",
        "/usr/lib/libreoffice/program/soffice",
        "/snap/bin/libreoffice",
    ]
    for c in candidates:
        if c and os.path.isfile(c):
            return c
    return _which("soffice") or _which("libreoffice")

def convert_doc_to_docx(doc_path: str, outdir: str | None = None) -> str:
    """
    Возвращает путь к .docx. Если вход уже .docx — возвращает исходный.
    Сначала пробуем LibreOffice, на Windows без него — MS Word COM.
    """
    if not doc_path.lower().endswith(".doc"):
        return doc_path

    src = os.path.abspath(doc_path)  # <-- ВАЖНО: абсолютный путь
    if not os.path.exists(src):
        raise FileNotFoundError(f"Файл не найден: {src}")

    # 1) LibreOffice (если установлен)
    soffice = find_soffice()
    if soffice:
        outdir = outdir or tempfile.mkdtemp(prefix="filldoc_")
        cmd = [soffice, "--headless", "--nologo", "--convert-to", "docx", "--outdir", outdir, src]
        proc = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
        if proc.returncode != 0:
            raise RuntimeError(f"LibreOffice: ошибка конвертации .doc→.docx:\n{proc.stderr or proc.stdout}")
        out = os.path.join(outdir, pathlib.Path(src).with_suffix(".docx").name)
        if not os.path.exists(out):
            # запасной поиск (иногда имя/регистры меняются)
            stem = pathlib.Path(src).stem
            for p in pathlib.Path(outdir).glob(f"{stem}*.docx"):
                out = str(p); break
        if not os.path.exists(out):
            raise RuntimeError("LibreOffice отработал, но .docx не найден.")
        return out

    # 2) Windows COM (MS Word)
    if os.name == "nt":
        try:
            import win32com.client  # pip install pywin32
            wdFormatXMLDocument = 16
            outdir = outdir or tempfile.mkdtemp(prefix="filldoc_")
            out_path = os.path.join(outdir, pathlib.Path(src).with_suffix(".docx").name)

            word = win32com.client.gencache.EnsureDispatch("Word.Application")
            word.Visible = False
            try:
                # Нормализованный абсолютный путь
                doc = word.Documents.Open(os.path.normpath(src))
                doc.SaveAs(os.path.normpath(out_path), FileFormat=wdFormatXMLDocument)
                doc.Close(False)
            finally:
                word.Quit()
            return out_path
        except Exception as e:
            raise RuntimeError(
                "Нет LibreOffice и не удалось конвертировать через MS Word COM. "
                "Поставьте LibreOffice (soffice) или установите MS Word + pywin32."
            ) from e

    raise RuntimeError("LibreOffice (soffice) не найден. Установите LibreOffice или сконвертируйте .doc вручную.")
    