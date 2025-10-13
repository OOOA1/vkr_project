from typing import Optional, Tuple, List, Dict
import re
from docx.document import Document
from docx.text.paragraph import Paragraph
from docx.table import Table
from .utils import normalize, parse_target_step

UNDERSCORE_SEQ = re.compile(r"_{3,}")  # вписка = 3+ подчеркиваний
PLACEHOLDER_TOKENS = [
    r"И\.?\s*О\.?\s*Фамилия",   # И.О. Фамилия / И. О. Фамилия
]

class Locator:
    def __init__(self, doc: Document, min_line_len: int = 5):
        self.doc = doc
        self.min_line_len = max(1, int(min_line_len or 1))
        self.underline_re = re.compile(r"_{%d,}" % self.min_line_len)

    def anchor_prev_find(self, anchor: str, occur: int = 1) -> Paragraph | None:
    # """Найти абзац с anchor (без регистра) и вернуть ПРЕДЫДУЩИЙ абзац (туда и пишем)."""
        a = normalize(anchor).lower()
        paras = self.iter_paragraphs()
        c = 0
        for i, p in enumerate(paras):
            if a and a in normalize(p.text).lower():
                c += 1
                if c == (occur or 1):
                    return paras[i-1] if i > 0 else None
        return None

    def anchor_after_token_find(self, anchor: str, token: str = "/", occur: int = 1):
        """
        Находит абзац, где есть текст-якорь (без учёта регистра) и символ token (по умолчанию '/').
        Возвращает сам Paragraph. Нумерация occur идёт по абзацам, где встретился anchor.
        """
        a = normalize(anchor).lower()
        c = 0
        for p in self.iter_paragraphs():
            t_norm = normalize(p.text).lower()
            if a and a in t_norm and token in p.text:
                c += 1
                if c == (occur or 1):
                    return p
        return None

    # ---------- helpers ----------
    def iter_paragraphs(self) -> List[Paragraph]:
        return list(self.doc.paragraphs)

    def iter_tables(self) -> List[Table]:
        return list(self.doc.tables)

    def _find_segments_in_text(self, text: str) -> List[str]:
        segs = [m.group(0) for m in self.underline_re.finditer(text or "")]
        for pat in PLACEHOLDER_TOKENS:
            for m in re.finditer(pat, text or "", flags=re.IGNORECASE):
                segs.append(m.group(0))
        return segs

    # ---------- strategies ----------
    def by_number_find(self, number: int) -> Tuple[Paragraph, str] | None:
        count = 0
        for p in self.iter_paragraphs():
            for seg in self._find_segments_in_text(p.text):
                count += 1
                if count == number:
                    return (p, seg)
        return None

    def anchor_after_colon_find(self, anchor: str, occur: int = 1) -> Paragraph | None:
        c = 0
        a = normalize(anchor).lower()
        for p in self.iter_paragraphs():
            text_norm = normalize(p.text)
            if a and a in text_norm.lower() and ":" in p.text:
                c += 1
                if c == (occur or 1):
                    return p
        return None

    def anchor_segment_find(self, anchor: str, seg_index: int = 1, occur: int = 1) -> Tuple[Paragraph, str] | None:
        """
        Находим абзац с anchor (без учета регистра).
        Если seg_index > 0 — берем k-ю «вписку» в ЭТОМ абзаце.
        Если seg_index < 0 — берем k-ю «вписку» в ПРЕДЫДУЩЕМ абзаце (|seg_index|).
        """
        paras = self.iter_paragraphs()
        a = normalize(anchor).lower()
        c = 0
        for i, p in enumerate(paras):
            t_norm = normalize(p.text)
            if a and a in t_norm.lower():
                c += 1
                if c == (occur or 1):
                    target_p = p if seg_index >= 0 else (paras[i-1] if i > 0 else p)
                    segs = self._find_segments_in_text(target_p.text)
                    k = abs(seg_index or 1)
                    if 1 <= k <= len(segs):
                        return (target_p, segs[k - 1])
                    return None
        return None

    def table_label_find(self, label: str, target: str) -> Tuple[Table, int, int] | None:
        direction, step = parse_target_step(target)
        lbl = normalize(label).lower()
        for ti, tbl in enumerate(self.iter_tables()):
            for r, row in enumerate(tbl.rows):
                for c, cell in enumerate(row.cells):
                    if lbl and lbl in normalize(cell.text).lower():
                        rr, cc = r, c
                        if direction == "right":
                            cc += step
                        elif direction == "left":
                            cc -= step
                        elif direction == "below":
                            rr += step
                        elif direction == "above":
                            rr -= step
                        else:
                            cc += step
                        if 0 <= rr < len(tbl.rows) and 0 <= cc < len(tbl.rows[rr].cells):
                            return (tbl, rr, cc)
        return None

    def cell_ref_find(self, table_index: int, row: int, col: int) -> Tuple[Table, int, int] | None:
        tables = self.iter_tables()
        if 1 <= table_index <= len(tables):
            tbl = tables[table_index - 1]
            if 1 <= row <= len(tbl.rows) and 1 <= col <= len(tbl.rows[row - 1].cells):
                return (tbl, row - 1, col - 1)
        return None
    
    def between_words_find(self, left: str, right: str | None, occur: int = 1):
    
    # Находит абзац, где встречаются левый (=Anchor) и правый (=Label) маркеры.
    # Возвращает (Paragraph, left_str, right_str), чтобы потом заменить текст между ними.
    # Если right пустой -> меняем текст от left до конца строки.

        lq = normalize(left).lower()
        rq = normalize(right).lower() if right else None
        c = 0
        for p in self.iter_paragraphs():
            text = p.text
            low = text.lower()
            if lq and lq in normalize(low):
                if rq is None or (rq and rq in normalize(low)):
                    c += 1
                    if c == (occur or 1):
                        return (p, left, right)
        return None