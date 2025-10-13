from typing import Optional, Tuple
from docx.text.paragraph import Paragraph
from docx.table import _Cell
from docx.oxml.text.paragraph import CT_P
from .utils import apply_length_policy
import re

def set_paragraph_text_keep_style(p: Paragraph, new_text: str):
    for run in p.runs:
        run.clear()
    if p.runs:
        r = p.runs[0]
    else:
        r = p.add_run()
    r.text = new_text

def replace_underscore_segment_in_paragraph(p: Paragraph, segment_text: str, value: str, policy: str):
    src = p.text
    # добавим пробелы по краям, если их нет
    before_idx = src.find(segment_text)
    left_char = src[before_idx - 1] if before_idx > 0 else " "
    right_char = src[before_idx + len(segment_text)] if before_idx + len(segment_text) < len(src) else " "
    left_sp = "" if left_char.isspace() else " "
    right_sp = "" if right_char.isspace() else " "
    repl = apply_length_policy(segment_text, value, policy)
    replaced = src.replace(segment_text, f"{left_sp}{repl}{right_sp}", 1)
    set_paragraph_text_keep_style(p, replaced)

def replace_after_token_in_paragraph(p: Paragraph, token: str, value: str):
    """
    Заменяем всё ПОСЛЕ первого вхождения token на ' ' + value.
    Левую часть строки не трогаем (подчёркивания и пробелы сохраняются).
    """
    text = p.text
    idx = text.find(token)
    if idx == -1:
        return
    before = text[: idx + len(token)]  # включая сам token
    new_text = before.rstrip() + " " + value  # один пробел после токена
    set_paragraph_text_keep_style(p, new_text)

def write_to_cell(cell: _Cell, value: str):
    for p in list(cell.paragraphs):
        p.clear()
    cell.text = value

def replace_between_in_paragraph(p: Paragraph, left: str, right: str | None, value: str):
    text = p.text
    if right:
        pat = re.compile(rf"({re.escape(left)})\s*.*?\s*({re.escape(right)})",
                         flags=re.IGNORECASE | re.DOTALL)
        def _repl(m):
            L = m.group(1).rstrip()
            R = m.group(2).lstrip()
            return f"{L} {value} {R}"
        new_text = pat.sub(_repl, text, count=1)
    else:
        pat = re.compile(rf"({re.escape(left)})\s*.*$", flags=re.IGNORECASE | re.DOTALL)
        def _repl(m):
            L = m.group(1).rstrip()
            return f"{L} {value}"
        new_text = pat.sub(_repl, text, count=1)
    set_paragraph_text_keep_style(p, new_text)