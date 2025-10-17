# pip install pandas openpyxl docxtpl
import os
import re
from pathlib import Path
import pandas as pd
from docxtpl import DocxTemplate

# ---------- Настройки ----------
EXCEL_PATH = "main.xlsx"
SHEET_NAME = 0  # или имя листа
OUT_DIR = Path(r"C:\Users\artem\Desktop\vkr_project-main\output")

# Описываем каждый шаблон:
#   path   — путь к .docx шаблону
#   out    — паттерн имени итогового файла (можно подставлять колонки из Excel)
#   fields — сопоставление {переменная_в_шаблоне: колонка_в_Excel}
TEMPLATES = [
    {
        "path": "input/дневник.docx",
        "out":  "Дневник_{ФИО}_{Группа}.docx",
        "fields": {
            "org_name": "БазаПрактики",   # {{ org_name }} в шаблоне ← колонка БазаПрактики
            "student":  "ФИО",
            "group": "Группа",
            "pr_type": "ТипПрактики",
            "vidPractiki": "ВидПрактика",
            "kurs": "Курс",
            "kafedra": "Кафедра",
            "fio": "ФИО",
            "startPracticaDate": "НачалоПрактики",
            "endPracticaDate": "КонецПрактики",
            "initialStudent": "РасшифровкаСтудента",
            "RukProfOrg": "РукПрофОрг",
            "RukOrg": "РукВУЗ"
        }
    },
    
    {
        "path": "input/договор новый тз.docx",
        "out":  "Договор_новый_{ФИО}_{Группа}.docx",
        "fields": {
            "org_name": "БазаПрактики",   # {{ org_name }} в шаблоне ← колонка БазаПрактики
            "RukOrgFIO": "РукВУЗФИО",
            "burnOrgDate": "ДатаСозданияОрганизации",
            "UrAdrVUZ": "ЮрАдресПрофОрг",
            "INN": "ОргИНН",
            "dolj": "Должность",
            "org_adress": "АдресОрганизации",
            "strukPodr": "СтруктурноеПодразделение",
            "kab": "КабинетНомер",
            "group": "Группа",
            "kafedra": "Кафедра",
            "fio": "ФИО",
            "RukProfOrg": "РукПрофОрг",
            "startPracticaDate": "НачалоПрактики",
            "endPracticaDate": "КонецПрактики",
            "pr_type": "ТипПрактики",
            "vidPractiki": "ВидПрактика",
        }
    },
    
        {
        "path": "input/договор старый тз.docx",
        "out":  "Договор_старый_{ФИО}_{Группа}.docx",
        "fields": {
            "org_name": "БазаПрактики",   # {{ org_name }} в шаблоне ← колонка БазаПрактики
            "RukOrgFIO": "РукВУЗФИО",
            "burnOrgDate": "ДатаСозданияОрганизации",
            "UrAdrVUZ": "ЮрАдресПрофОрг",
            "INN": "ОргИНН",
            "dolj": "Должность",
            "org_adress": "АдресОрганизации",
            "strukPodr": "СтруктурноеПодразделение",
            "kab": "КабинетНомер",
            "group": "Группа",
            "kafedra": "Кафедра",
            "fio": "ФИО",
            "RukProfOrg": "РукПрофОрг",
            "startPracticaDate": "НачалоПрактики",
            "endPracticaDate": "КонецПрактики",
            "pr_type": "ТипПрактики",
            "vidPractiki": "ВидПрактика",
        }
    },

            {
                # ДОП СВЕДЕНИЯ
        "path": "input/Доп.сведения.docx",
        "out":  "Доп.сведения_{ФИО}_{Группа}.docx",
        "fields": {
            "org_name": "БазаПрактики",   # {{ org_name }} в шаблоне ← колонка БазаПрактики
            "RukOrgFIO": "РукВУЗФИО",
            "burnOrgDate": "ДатаСозданияОрганизации",
            "UrAdrVUZ": "ЮрАдресПрофОрг",
            "INN": "ОргИНН",
            "dolj": "Должность",
            "org_adress": "АдресОрганизации",
            "strukPodr": "СтруктурноеПодразделение",
            "kab": "КабинетНомер",
            "group": "Группа",
            "kafedra": "Кафедра",
            "fio": "ФИО",
            "fioRP": "ФИОвРП",
            "RukProfOrg": "РукПрофОрг",
            "startPracticaDate": "НачалоПрактики",
            "endPracticaDate": "КонецПрактики",
            "pr_type": "ТипПрактики",
            "vidPractiki": "ВидПрактика",
            "kurs": "Курс",
            "studyForm": "формыОбучения",
            "naprPodg": "направлениеПодготовки",
            "profile": "ПрофильОбучения",
            "SvedOFormStudy": "СведенияОФормеОбучения",
            "SvedOProhObuch": "СведенияОПрохЧастиОбрПрог",
            "SvedOProhUskObuch": "СведОПрохУскорОбуч",
            "initialStudent": "РасшифровкаСтудента",
        }
    },
    
            {
                # задание на ВКР 1
        "path": "input/Задание на ВКР - 1.docx",
        "out":  "Задание_на_ВКР - 1_{ФИО}_{Группа}.docx",
        "fields": {
            "naprPodg": "направлениеПодготовки",
            "kurs": "Курс",
            "group": "Группа",
            "studyForm": "формыОбучения",
            "stepenNauchRuk": "СтепеньНаучРук",
            "ZvanieNauchRuk": "ЗваниеНаучРук",
            "fio": "ФИО",
            "fioDP": "ФИОДП",
            "initialRukVRK": "ИницРукВКР",
            "srokSdachiVRK": "СрокСдачиВКР"
        }
    },

            {
        # Задание на ВКР 2
        "path": "input/Задание на ВКР - 2.docx",
        "out":  "Задание_на_ВКР - 2_{ФИО}_{Группа}.docx",
        "fields": {
            "fioDP": "ФИОДП",
            "group": "Группа",
            "stepenNauchRuk": "СтепеньНаучРук",
            "ZvanieNauchRuk": "ЗваниеНаучРук",
            "initialRukVRK": "ИницРукВКР",
            "initialStudent": "РасшифровкаСтудента",
        }
    },

            {
        # заявление на АП
        "path": "input/Заявление на АП - 2.docx",
        "out":  "Заявление на АП - 2_{ФИО}_{Группа}.docx",
        "fields": {
            "zavKaf": "ЗаведующийКафедрой",
            "fioRP": "ФИОвРП",
            "group": "Группа",
            "naprPodg": "НаправлениеПодготовки",
            "FIONauchRuk": "ФИОНаучРук",
            "kafedraRP": "КафедраРП",
            "doljNauchRuk": "ДолжНаучРук",
            "fio": "ФИО",
        }
    },

            {
        # Заявление на размещение в ЭБС
        "path": "input/Заявление на размещение в ЭБС.docx",
        "out":  "Заявление на размещение в ЭБС_{ФИО}_{Группа}.docx",
        "fields": {
            "fioRP": "ФИОвРП",
            "group": "Группа",
            "naprPodg": "НаправлениеПодготовки",
            "fio": "ФИО",
            "kafedra": "Кафедра",
            "Date": "СегодняшняяДата"
        }
    },

            {
        # Заявление
        "path": "input/заявление.docx",
        "out":  "заявление_{ФИО}_{Группа}.docx",
        "fields": {
            "zavKaf": "ЗаведующийКафедрой",
            "kurs": "Курс",
            "studyForm": "формыОбучения",
            "group": "Группа",
            "naprPodg": "направлениеПодготовки",
            "profile": "ПрофильОбучения",
            "pr_type": "ТипПрактики",
            "vidPractiki": "ВидПрактика",
            "startPracticaDate": "НачалоПрактики",
            "endPracticaDate": "КонецПрактики",
            "org_name": "БазаПрактики",
            "strukPodr": "СтруктурноеПодразделение",
            "RukProfOrg": "РукПрофОрг",
            "dolj": "Должность",
            "tel": "телефон",
            "studEmail": "emailстуд",
            "Date": "СегодняшняяДата",
            "fioRukProfOrg": "ФИОРукПрофОрг",
            "fio": "ФИО",

        }
    },

                {
        # Заявление
        "path": "input/индивидуальное задание.docx",
        "out":  "индивидуальное задание_{ФИО}_{Группа}.docx",
        "fields": {
            "workObject": "ОбъектРаботы",
            "pr_type": "ТипПрактики",
            "kafedra": "Кафедра",
            "fioDP": "ФИОДП",
            "group": "Группа",
            "tel": "телефон",
            "studEmail": "emailстуд",
            "initialRukVRK": "ИницРукВКР",
            "stepenNauchRuk": "СтепеньНаучРук",
            "ZvanieNauchRuk": "ЗваниеНаучРук",
            "org_name": "БазаПрактики",
            "startPracticaDate": "НачалоПрактики",
            "endPracticaDate": "КонецПрактики",
            "RukProfOrg": "РукПрофОрг",
            "RukOrg": "РукВУЗ",
            "initialStudent": "РасшифровкаСтудента",
            "Date": "СегодняшняяДата"

        }
    },
]


# ---------- Хелперы ----------
INVALID_FS = r'[<>:"/\\|?*]'

def safe(v):
    """Строка без NaN/None и лишних пробелов."""
    if pd.isna(v):
        return ""
    return str(v).strip()

def slugify_filename(name: str) -> str:
    """Делаем валидное имя файла под Windows/macOS/Linux."""
    name = re.sub(INVALID_FS, "_", name)
    name = name.rstrip(" .")  # нельзя заканчивать точкой/пробелом в Windows
    return name or "file"

class SafeMap(dict):
    """format_map, который не падает, если ключа нет."""
    def __missing__(self, key):
        return ""

def row_to_ctx(row: pd.Series, mapping: dict) -> dict:
    """Строим контекст для docxtpl по mapping {tpl_key: excel_col}."""
    return {tpl_key: safe(row.get(excel_col)) for tpl_key, excel_col in mapping.items()}

# ---------- Основной код ----------
df = pd.read_excel(EXCEL_PATH, sheet_name=SHEET_NAME)
df = df.fillna("")  # чтобы не ловить NaN

OUT_DIR.mkdir(parents=True, exist_ok=True)

for idx, row in df.iterrows():
    # Для каждого шаблона формируем свой контекст и своё имя файла
    for tpl in TEMPLATES:
        ctx = row_to_ctx(row, tpl["fields"])
        doc = DocxTemplate(tpl["path"])
        doc.render(ctx)

        # Подставляем значения из строки в имя файла
        out_name = tpl["out"].format_map(SafeMap({k: safe(v) for k, v in row.items()}))
        out_path = OUT_DIR / slugify_filename(out_name)

        doc.save(out_path)
        print(f"[OK] {out_path}")
