# filldoc/templates.py
# -*- coding: utf-8 -*-
from dataclasses import dataclass, field
from typing import List, Dict, Any, Tuple

@dataclass
class DetectSpec:
    # базовые сигналы
    required: List[str] = field(default_factory=list)
    optional: List[str] = field(default_factory=list)

    # негативы
    negative_soft: List[str] = field(default_factory=list)   # штраф, но не запрет
    negative_hard: List[str] = field(default_factory=list)   # жёсткий запрет

    # ОБРАТНАЯ СОВМЕСТИМОСТЬ: старый ключ 'negative'
    negative: List[str] = field(default_factory=list)

    # макет и порог
    layout: Dict[str, Any] = field(default_factory=dict)
    threshold: int = 10

    # строгие условия
    must_all: List[str] = field(default_factory=list)              # все должны встретиться
    any_of: List[List[str]] = field(default_factory=list)          # из каждой группы хотя бы одна
    must_exact_lines: List[str] = field(default_factory=list)      # точное совпадение строки
    must_regex: List[str] = field(default_factory=list)            # regex-совпадение
    require_same_line: List[Tuple[str, str]] = field(default_factory=list)  # пара слов в одной строке

    def __post_init__(self):
        # если передали старый 'negative', сливаем его в negative_soft
        if self.negative:
            # сохраняем порядок, убираем дубли
            merged = list(dict.fromkeys((self.negative_soft or []) + self.negative))
            self.negative_soft = merged
            self.negative = []

@dataclass
class TemplateSpec:
    id: str
    human_name: str
    detect: DetectSpec
    settings: Dict[str, Any]
    mapping: List[Dict[str, Any]]  # список правил как в Excel/mapping

# ----------------------------
# Пример: ДНЕВНИК (DN)
# ----------------------------
DN = TemplateSpec(
    id="DN",
    human_name="Дневник практики",
    detect=DetectSpec(
        required=[
            "дневник", "тип:", "наименование базы практики",
            "срок прохождения практики", "студент"
        ],
        optional=[
            "руководитель практики от организации (вуза)",
            "руководитель практики от профильной организации",
            "фамилия, имя, отчество полностью"
        ],
        negative=[
            "график", "совместный график",
            "индивидуальное задание",
            "отчет", "отчёт", "титул отчета",
            "характеристика",
            "договор"
        ],
        layout={"must_have_slash": True, "min_tables": 1, "min_underscores": 3},
        threshold=20,                         # было 10 → стало 20
        must_all=["дневник", "наименование базы практики", "срок прохождения практики"],  # обязаны быть
    ),
    settings={
        "маска_имени_файла": "{{Группа}}_{{ФИО}}_дневник.docx",
        "минимальная_длина_линии": 3,
        "политика_длины_значения": "underline_and_keep_line",
        "filename_globs": ["*дневник*.docx"],
    },
    mapping=[
        # верхняя часть
        {"Field":"ВидПрактики","Method":"between_words","Anchor":"прохождения","Label":"практики","Occur":1},
        {"Field":"ТипПрактики","Method":"between_words","Anchor":"тип:","Label":")","Occur":1},
        {"Field":"Курс","Method":"between_words","Anchor":"студента","Label":"курса","Occur":1},
        {"Field":"Группа","Method":"between_words","Anchor":"группы","Label":"","Occur":1},
        {"Field":"Кафедра","Method":"between_words","Anchor":"кафедры","Label":"","Occur":1},

        # большая линия под примечанием
        {"Field":"ФИО","Method":"anchor_prev","Anchor":"фамилия, имя, отчество полностью","Segment":-1,"Occur":1,"Transform":"Title"},

        # хвост (правее "/")
        {"Field":"РукПрофОрг","Method":"anchor_after_slash","Anchor":"Руководитель практики от профильной организации","Occur":1,"Transform":"FIO_INITIALS_SURNAME"},
        {"Field":"РукВуз","Method":"anchor_after_slash","Anchor":"Руководитель практики от организации (вуза)","Occur":1,"Transform":"FIO_INITIALS_SURNAME"},
        {"Field":"ФИО","Method":"anchor_after_slash","Anchor":"Студент","Occur":1,"Transform":"FIO_INITIALS_SURNAME"},

        # блоки с двоеточием
        {"Field":"БазаПрактики","Method":"anchor_after_colon","Anchor":"Наименование базы практики","Occur":1},
        {"Field":"Срок","Method":"anchor_after_colon","Anchor":"Срок прохождения практики","Occur":1},

        # таблица с датой (в левую ячейку ярлыка)
        {"Field":"ДатаНачала","Method":"table_label","Label":"Дата начала практики","Target":"left(0)","Occur":1,"Transform":"date:%d.%m.%Y"},
    ]
)

# ----------------------------
# Здесь добавляй ещё 29 спецификаций:
# ZAD = TemplateSpec(id="ZAD", human_name="Задание на практику", detect=..., settings=..., mapping=[...])
# TITLE = TemplateSpec(id="TITLE", human_name="Титульный лист", detect=..., settings=..., mapping=[...])
# ...
# ----------------------------

TEMPLATES: Dict[str, TemplateSpec] = {
    DN.id: DN,
    # "ZAD": ZAD,
    # "TITLE": TITLE,
}

# === ТИТУЛ ВКР ===================================================
TITLE_VKR = TemplateSpec(
    id="TITLE_VKR",
    human_name="Титульный лист ВКР",
    detect=DetectSpec(
        required=[
            "«МОСКОВСКИЙ МЕЖДУНАРОДНЫЙ УНИВЕРСИТЕТ»",
            "ВЫПУСКНАЯ КВАЛИФИКАЦИОННАЯ РАБОТА",
            "Направление подготовки:",
            "на тему:",
            "Автор работы:",
            "студент группы",
        ],
        optional=["Руководитель работы:", "Заведующий выпускающей кафедрой:"],
        negative=["график", "совместный", "дневник", "индивидуальное задание", "заявление", "отчет", "отчёт"],
        layout={"min_underscores": 1},
        threshold=25,
        must_all=["ВЫПУСКНАЯ КВАЛИФИКАЦИОННАЯ РАБОТА", "на тему:", "Автор работы:"],
    ),
    settings={
        "маска_имени_файла": "{{Группа}}_{{ФИО}}_титул.docx",
        "минимальная_длина_линии": 3,
        "политика_длины_значения": "underline_and_keep_line",
    },
    mapping=[
        {"Field":"Направление", "Method":"anchor_after_colon", "Anchor":"Направление подготовки", "Occur":1},
        {"Field":"Тема",         "Method":"anchor_after_colon", "Anchor":"на тему",               "Occur":1},
        {"Field":"ФИО",          "Method":"anchor_after_colon", "Anchor":"Автор работы",          "Occur":1, "Transform":"Title"},
        {"Field":"Группа",       "Method":"between_words",      "Anchor":"студент группы",        "Label":"EOL", "Occur":1},
        {"Field":"Форма",        "Method":"between_words",      "Anchor":"форма обучения",        "Label":"EOL", "Occur":1},
        {"Field":"Руководитель", "Method":"anchor_after_colon", "Anchor":"Руководитель работы",   "Occur":1, "Transform":"Title"},
        {"Field":"ЗавКафедрой", "Method":"anchor_after_colon", "Anchor":"Заведующий выпускающей кафедрой", "Occur":1, "Transform":"Title"},
    ],
)

# === ЗАДАНИЕ НА ВКР — вариант 1 ================================
ZAD1 = TemplateSpec(
    id="ZAD1",
    human_name="Задание на ВКР (вариант 1)",
    detect=DetectSpec(
        # Ключевые слова именно варианта 1
        must_all=[
            "ЗАДАНИЕ",
            "на выпускную квалификационную работу",
        ],
        required=[
            "Студенту",  # строка в шапке "Студенту ____ курса группы ________"
            "Тема выпускной квалификационной работы:",
        ],
        optional=[
            "Направление подготовки:",
            "Руководитель выпускной квалификационной работы:",
            "Задание принял к исполнению:",
            "Срок сдачи студентом законченной работы",  # часто присутствует
        ],
        # Отрезаем лексику варианта 2
        negative=[
            "ЗАДАНИЕ по подготовке",
            "Обучающемуся",
            "Тема:",
            "Научный руководитель:",
        ],
        negative_soft=["характеристика", "дневник", "совместный рабочий график"],
        layout={"min_underscores": 2},
        threshold=13,
    ),
    settings={
        "маска_имени_файла": "задание1_{{ФИО}}.docx",
        "минимальная_длина_линии": 3,
    },
    mapping=[
        {"Field":"Курс", "Method":"between_words", "From":"Студенту", "To":"курса", "Transform":"AsIs"},
        {"Field":"Группа", "Method":"between_words", "From":"группы", "To":"EOL", "Transform":"AsIs"},
        {"Field":"ТемаВКР", "Method":"anchor_after_colon", "Anchor":"Тема выпускной квалификационной работы", "Occur":1, "Transform":"AsIs"},
        {"Field":"Направление", "Method":"anchor_after_colon", "Anchor":"Направление подготовки", "Occur":1, "Transform":"AsIs"},
        {"Field":"РукВКР", "Method":"anchor_after_colon", "Anchor":"Руководитель выпускной квалификационной работы", "Occur":1, "Transform":"Title"},
        {"Field":"ФИОКоротко", "Method":"anchor_after_slash", "Anchor":"Задание принял к исполнению", "Occur":1, "Transform":"Initials"},  # студент «И.И. Иванов»
        {"Field":"ДатаСдачи", "Method":"anchor_after_colon", "Anchor":"Срок сдачи студентом законченной работы", "Occur":1, "Transform":"DateDDMMYYYY"},
    ]
)

# === ЗАДАНИЕ НА ВКР — вариант 2 ================================
ZAD2 = TemplateSpec(
    id="ZAD2",
    human_name="Задание на ВКР (вариант 2)",
    detect=DetectSpec(
        required=[
            "ЗАДАНИЕ",
            "по подготовке",
            "Обучающемуся",
            "Тема:",
            "Научный руководитель:",
        ],
        optional=["Задание принял к исполнению:", "Срок сдачи исполнителем законченной работы"],
        negative=["ДНЕВНИК"],
        layout={"min_underscores": 3},
        threshold=9,
    ),
    settings={
        "маска_имени_файла": "{{Группа}}_{{ФИО}}_задание2.docx",
        "минимальная_длина_линии": 3,
        "политика_длины_значения": "underline_and_keep_line",
    },
    mapping=[
        {"Field":"ФИО",          "Method":"between_words",      "Anchor":"Обучающемуся", "Label":"EOL", "Occur":1},
        {"Field":"Тема",         "Method":"anchor_after_colon", "Anchor":"Тема",         "Occur":1},
        {"Field":"Руководитель", "Method":"anchor_after_colon", "Anchor":"Научный руководитель", "Occur":1, "Transform":"Title"},
        # при необходимости позже добавим группу/срок и подписи
    ],
)

# === ДЕТЕКТ для прочих форм (mapping добавим позже) ============
DOPSVE = TemplateSpec(
    id="DOPSVE",
    human_name="Заявление: Дополнительные сведения к диплому",
    detect=DetectSpec(
        required=["заявление", "дополнительные сведения", "ректору", "приложения к диплому"],
        optional=["(ф.и.о. обучающегося)", "курса", "учебная группа", "направление подготовки", "профиль"],
        negative=["каникул", "отказываюсь от предоставления", "последипломного отпуска"],
        layout={"min_tables": 1, "min_underscores": 3},
        threshold=20,
        must_all=["дополнительные сведения"],  # ← добавили
        any_of=[["приложения к диплому", "к диплому"]],
    ),
    settings={"маска_имени_файла": "{{Группа}}_{{ФИО}}_доп_сведения.docx",
              "минимальная_длина_линии": 3, "политика_длины_значения": "underline_and_keep_line",
              "filename_globs": ["*доп*свед*.docx", "*доп_сведения*.docx", "*доп сведения*.docx"],},
    mapping=[
        {"Field": "ФИО", "Method": "replace", "Anchor": "от обучающегося", "Label": "курса", "Occur": 1},
        {"Field": "Курс", "Method": "between_words", "Anchor": "курса", "Label": "формы обучения", "Occur": 1},
        {"Field": "Группа", "Method": "between_words", "Anchor": "группы", "Label": "курса", "Occur": 1},
        {"Field": "Направление", "Method": "between_words", "Anchor": "Направление подготовки", "Label": "профиль", "Occur": 1},
        {"Field": "Профиль", "Method": "between_words", "Anchor": "профиль", "Label": "приложение", "Occur": 1},
        {"Field": "ФИО", "Method": "anchor_prev", "Anchor": "(Ф.И.О. обучающегося)", "Occur": 1, "Segment": -1},
        {"Field": "ФИО", "Method": "anchor_after_slash", "Anchor": "(расшифровка подписи)", "Occur": 1, "Transform": "FIO_INITIALS_SURNAME"},
    ],
)

AP2 = TemplateSpec(
    id="AP2",
    human_name="Заявление на АП (Антиплагиат)",
    detect=DetectSpec(
        required=["Заявление", "С фактом проверки", "Антиплагиат"],
        optional=["От обучающегося", "(ФИО, должность руководителя ВКР)"],
        negative=["ЭБС", "ЗАДАНИЕ"],
        layout={"min_underscores": 2},
        threshold=8,
    ),
    settings={"маска_имени_файла": "{{Группа}}_{{ФИО}}_АП.docx",
              "минимальная_длина_линии": 3, "политика_длины_значения": "underline_and_keep_line"},
    mapping=[

    ],
)

EBS = TemplateSpec(
    id="EBS",
    human_name="Заявление/Согласие на размещение в ЭБС",
    detect=DetectSpec(
        required=["ЗАЯВЛЕНИЕ/СОГЛАСИЕ", "электронно-библиотечной системе", "предоставляю выпускную квалификационную работу на тему"],
        optional=["Ректору", "(ФИО обучающегося)", "(номер группы)"],
        negative=["Антиплагиат"],
        layout={"min_tables": 1},
        threshold=9,
    ),
    settings={"маска_имени_файла": "{{Группа}}_{{ФИО}}_ЭБС.docx",
              "минимальная_длина_линии": 3, "политика_длины_значения": "underline_and_keep_line"},
    mapping=[],
)

KANIK = TemplateSpec(
    id="KANIK",
    human_name="Заявление на каникулы",
    detect=DetectSpec(
        required=[
            "заявление",
            "отказываюсь от предоставления мне каникул",   # ключевая фраза
        ],
        optional=[
            "последипломного отпуска",
            "тел.",
            "e-mail",
        ],
        negative=[
            "дополнительные сведения",
            "приложения к диплому",
            "электронно-библиотечной системе",
            "антиплагиат",
        ],
        layout={"min_underscores": 2},
        threshold=9,
    ),
    settings={"маска_имени_файла": "{{Группа}}_{{ФИО}}_каникулы.docx",
              "минимальная_длина_линии": 3, "политика_длины_значения": "underline_and_keep_line"},
    mapping=[],
)

LSL = TemplateSpec(
    id="LSL",
    human_name="Лист согласования личных сведений",
    detect=DetectSpec(
        required=["Лист согласования личных сведений", "ФИО по паспорту:", "Для иностранных граждан"],
        optional=["«МОСКОВСКИЙ МЕЖДУНАРОДНЫЙ УНИВЕРСИТЕТ»", "MOSCOW INTERNATIONAL UNIVERSITY"],
        negative=["ОЗНАКОМИТЕЛЬНЫЙ ЛИСТ"],
        layout={"min_underscores": 2, "min_tables": 1},
        threshold=9,
    ),
    settings={"маска_имени_файла": "{{Группа}}_{{ФИО}}_лист_согласования.docx",
              "минимальная_длина_линии": 3, "политика_длины_значения": "underline_and_keep_line"},
    mapping=[],
)

OZN = TemplateSpec(
    id="OZN",
    human_name="Ознакомительный лист",
    detect=DetectSpec(
        required=["ОЗНАКОМИТЕЛЬНЫЙ ЛИСТ", "Перечень документов Университета"],
        optional=["Подписать ознакомительный лист необходимо до дня проведения ГИА."],
        negative=["ЗАДАНИЕ", "ДНЕВНИК"],
        layout={"min_tables": 1},
        threshold=9,
    ),
    settings={"маска_имени_файла": "{{Группа}}_{{ФИО}}_ознакомительный_лист.docx",
              "минимальная_длина_линии": 3, "политика_длины_значения": "underline_and_keep_line"},
    mapping=[],
)

INDZAD = TemplateSpec(
    id="INDZAD",
    human_name="Индивидуальное задание",
    detect=DetectSpec(
        must_exact_lines=[
            "ИНДИВИДУАЛЬНОЕ ЗАДАНИЕ",
        ],
        required=[
            "ИНДИВИДУАЛЬНОЕ ЗАДАНИЕ",
            "(тип:",                  # строка вида "(тип: _____________)"
            "Выдано студенту",       # «Выдано студенту …» есть в шапке
            "Сроки прохождения:",    # «Сроки прохождения: с … по …»
        ],
        must_regex=[
            r"Сроки прохождения:\s*с .*202_ г\.\s*по.*202_ г\.",  # терпим произвольные пробелы/подчёркивания
        ],
        negative_hard=["выпускная квалификационная работа"],
        negative_soft=["дневник","задание на вкр","титульный лист","график","договор","отчет","отчёт"],
        layout={"min_underscores": 2},
        threshold=10,
    ),
    settings={
        "маска_имени_файла": "{{Группа}}_{{ФИО}}_инд_задание.docx",
        "минимальная_длина_линии": 3,
        "политика_длины_значения": "underline_and_keep_line",
    },
    mapping=[],
)

# Совместный график
SOVGRAF = TemplateSpec(
    id="SOVGRAF",
    human_name="Совместный рабочий график (план)",
    detect=DetectSpec(
        must_exact_lines=[
            "СОВМЕСТНЫЙ РАБОЧИЙ ГРАФИК (ПЛАН)",
        ],
        required=[
            "СОВМЕСТНЫЙ РАБОЧИЙ ГРАФИК (ПЛАН)",
            "(тип:",
            "Направление подготовки:",
            "________ практики",
        ],
        must_regex=[
            r"Срок прохождения практики:\s*с .*202_ г\.\s*по .*202_ г\.",  # устойчивый шаблон дат
        ],
        negative_hard=["выпускная квалификационная работа"],
        negative_soft=["задание","индивидуальное задание","титульный лист","отчет","отчёт","дневник"],
        layout={"min_tables": 1},
        threshold=10,
    ),
    settings={
        "маска_имени_файла": "график_{{Группа}}.docx",
        "минимальная_длина_линии": 3,
        "политика_длины_значения": "underline_and_keep_line",
    },
    mapping=[],
)

# Титул отчёта по практике
TITLE_OTCH = TemplateSpec(
    id="TITLE_OTCH",
    human_name="Титульный лист отчёта",
    detect=DetectSpec(
        required=["отчет", "направление подготовки", "руководитель"],
        optional=["автор", "студент", "группа", "форма обучения"],
        negative=["график", "совместный", "индивидуальное задание", "дневник", "задание на вкр", "выпускная квалификационная работа"],
        layout={"min_underscores": 1},
        threshold=12,
        must_all=["отчет"],
    ),
    settings={"маска_имени_файла": "{{Группа}}_{{ФИО}}_титул_отчета.docx",
              "минимальная_длина_линии": 3, "политика_длины_значения": "underline_and_keep_line"},
    mapping=[],
)

# Договор (новый)
DOG_NEW = TemplateSpec(
    id="DOG_NEW",
    human_name="Договор (новая форма)",
    detect=DetectSpec(
        # Дадим детектору «жирные» маркеры, которые реально есть в твоём файле договора
        must_exact_lines=["ДОГОВОР № ____________"],   # заголовок
        required=[
            "договор",                                 # на всякий случай, нечувствительно к регистру
            "Полное наименование организации:",        # в блоках «Университет» и «Профильная организация»
            "ИНН",                                     # реквизиты
            "Адреса, реквизиты и подписи Сторон",      # раздел 5
        ],
        optional=[
            "Приложение 1 к договору", "Приложение 2 к договору",
            "Профильная организация", "Университет",
            "Проректор по развитию",
        ],
        negative=["дневник", "задание", "титул", "график", "отчет", "отчёт"],
        layout={"min_underscores": 3},
        threshold=4,        # снизили порог, чтобы уверенно срабатывало
        must_all=["договор"],   # закрепили ключевое слово
    ),
    settings={
        "маска_имени_файла": "договор_новый_{{ФИО}}.docx",
        "минимальная_длина_линии": 3,
        "политика_длины_значения": "underline_and_keep_line",
        "filename_globs": ["*договор*тз*.docx", "*договор новый*.docx"],
    },
    mapping=[
        # ПРЕАМБУЛА — заменить только "1. Название предприятия" (без захода на формат рядом)
        {"Field": "БазаПрактики", "Method": "between_words", "Anchor": "1.", "Label": "именуемая", "Occur": 1},

        {"Field": "РукПрофОрг", "Method": "between_words", "Anchor": "2.", "Label": ", действующего(ей)", "Occur": 1},

        {"Field": "ДатаСозданияОрг", "Method": "between_words", "Anchor": "3.", "Label": "г.", "Occur": 1},

        {"Field": "БазаПрактики", "Method": "between_words",
        "Anchor": "уставом ", "Label": "EOL", "Occur": 1},

        {"Field": "БазаПрактики", "Method": "between_words",
        "Anchor": "уставом\u00A0", "Label": "EOL", "Occur": 1},   # NBSP после "уставом"

        {"Field": "БазаПрактики", "Method": "between_words",
        "Anchor": "уставом\t", "Label": "EOL", "Occur": 1},      # таб после "уставом"

        {"Field": "БазаПрактики", "Method": "between_words", "Anchor": "1.", "Label": "именуемая", "Occur": 1},

        {"Field": "БазаПрактики", "Method": "between_words", "Anchor": "1.", "Label": "EOL", "Occur": 2},

        {"Field":"БазаПрактики","Method":"between_words","Anchor":", уставом 1.","Label":"EOL","Occur":1},
        {"Field":"БазаПрактики","Method":"between_words","Anchor":", уставом\u00A01.","Label":"EOL","Occur":1},  # NBSP
        {"Field":"БазаПрактики","Method":"between_words","Anchor":", уставом\t1.","Label":"EOL","Occur":1},     # таб
    ]
)

# Договор (старый)
DOG_OLD = TemplateSpec(
    id="DOG_OLD",
    human_name="Договор (старая форма)",
    detect=DetectSpec(
        required=["договор"],
        optional=["предмет договора", "стороны", "исполнитель", "заказчик"],
        negative=["дневник", "задание", "титул", "график", "отчет"],
        layout={},
        threshold=7,              # ← было 8
        must_all=["договор"],
    ),
    settings={"маска_имени_файла": "договор_старый_{{ФИО}}.docx",
              "минимальная_длина_линии": 3, "политика_длины_значения": "underline_and_keep_line"},
    mapping=[],
)

# Характеристика
HARAKT = TemplateSpec(
    id="HARAKT",
    human_name="Характеристика",
    detect=DetectSpec(
        # Точный заголовок — сильный якорь
        must_exact_lines=["ХАРАКТЕРИСТИКА"],
        required=[
            "настоящая характеристика дана",                 # вводная строка
            "проходившему (шей) учебную практику",           # без хвоста в скобках, он может меняться
            "за время прохождения практики изучил(а):",
            "в период прохождения практики решались следующие задачи:",
        ],
        optional=[
            "результат работы обучающегося",                 # может быть: "Результат работы обучающегося(щейся):"
            "руководитель практики от профильной организации",
            "(наименование организации)",
            "(фактический адрес)",
        ],
        # Жёстко отбрасываем другие типы:
        negative=[
            "выпускная квалификационная работа",
            "заявление",
            "электронно-библиотечной системе",
            "дополнительные сведения",
            "договор",
            "индивидуальное задание",
            "задание",
            "дневник",
            "титульный лист",
            "совместный рабочий график",
        ],
        layout={"min_underscores": 1},
        threshold=12,
    ),
    settings={
        "маска_имени_файла": "характеристика_{{ФИО}}.docx",
        "минимальная_длина_линии": 3,
        "политика_длины_значения": "underline_and_keep_line",
    },
    mapping=[
        {"Field":"ФИО", "Method":"anchor_prev", "Anchor":"(Ф.И.О. обучающегося)", "Transform":"Title"},
        {"Field":"ТипПрактикиКоротко", "Method":"between_words", "Anchor":"Проходившему (шей) учебную практику (тип:", "Label":")", "Transform":"AsIs"},
        {"Field":"Организация", "Method":"anchor_prev", "Anchor":"(наименование организации)", "Transform":"AsIs"},
        {"Field":"Адрес", "Method":"anchor_prev", "Anchor":"(фактический адрес)", "Transform":"AsIs"},
        {"Field":"РукПрофОргКоротко", "Method":"anchor_after_slash", "Anchor":"Руководитель практики от профильной организации", "Occur":1, "Transform":"Initials"},
    ]
)

# Регистрация (расположи выше ZAD1/ZAD2, чтобы не перехватывали)
TEMPLATES.update({
    "HARAKT": HARAKT,
    # остальные шаблоны...
})

ZAYAV_GEN = TemplateSpec(
    id="ZAYAV_GEN",
    human_name="Заявление (общее)",
    detect=DetectSpec(
        must_exact_lines=["ЗАЯВЛЕНИЕ"],
        required=[
            "ЗАЯВЛЕНИЕ",
            "От обучающегося",
        ],
        optional=[
            "Ректору",
            "группа", "курса",
            "Направление подготовки", "Профиль",
            "(ФИО, должность руководителя ВКР)",
        ],
        # Жёстко отсекаем другие уже существующие типы заявлений
        negative_hard=[
            "Антиплагиат",
            "электронно-библиотечной системе",
            "Дополнительные сведения",
            "ВЫПУСКНАЯ КВАЛИФИКАЦИОННАЯ РАБОТА",
        ],
        negative_soft=[
            "ИНДИВИДУАЛЬНОЕ ЗАДАНИЕ", "ДНЕВНИК", "ГРАФИК", "ТИТУЛЬНЫЙ",
            "ОЗНАКОМИТЕЛЬНЫЙ ЛИСТ",
        ],
        layout={"min_underscores": 2},
        threshold=9,
    ),
    settings={
        "маска_имени_файла": "заявление_{{ФИО}}.docx",
        "минимальная_длина_линии": 3,
        "политика_длины_значения": "underline_and_keep_line",
    },
    mapping=[],
)

# регистрируем
TEMPLATES.update({
    "INDZAD": INDZAD,
    "SOVGRAF": SOVGRAF,
    "TITLE_OTCH": TITLE_OTCH,
    "DOG_NEW": DOG_NEW,
    "DOG_OLD": DOG_OLD,
    "HARAKT": HARAKT,
    "ZAYAV_GEN": ZAYAV_GEN,
})

# регистрируем
TEMPLATES.update({
    "TITLE_VKR": TITLE_VKR,
    "ZAD1": ZAD1,
    "ZAD2": ZAD2,
    "DOPSVE": DOPSVE,
    "AP2": AP2,
    "EBS": EBS,
    "KANIK": KANIK,
    "LSL": LSL,
    "OZN": OZN,
})