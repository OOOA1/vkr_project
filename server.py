# server.py
# Зависимости:
#   pip install fastapi "uvicorn[standard]" python-multipart pandas openpyxl docxtpl requests

import io
import re
import csv
import zipfile
from pathlib import Path
from typing import Optional, Dict, Tuple, List

import pandas as pd
import requests
from fastapi import FastAPI, File, Form, UploadFile, HTTPException
from fastapi.responses import (
    HTMLResponse,
    JSONResponse,
    StreamingResponse,
    PlainTextResponse,
    FileResponse,
)
from docxtpl import DocxTemplate

from templates_config import TEMPLATES  # твои захардкоженные шаблоны

app = FastAPI(title="DOCX Templater → ZIP (wide or KV)", version="3.4.0")

# === Настройки выдачи шаблона Excel ===
BASE_DIR = Path(__file__).resolve().parent
TEMPLATE_DOWNLOAD_NAME = "main_example.xlsx"
CANDIDATE_TEMPLATE_PATHS: List[Path] = [
    BASE_DIR / "main_example.xlsx",
    BASE_DIR / "main.xlsx",
    BASE_DIR / "main — копия.xlsx",
]

# ============= UI =============
INDEX_HTML = """
<!doctype html><meta charset="utf-8">
<title>DOCX → ZIP</title>
<style>
  :root{color-scheme:light dark}
  body{font-family:system-ui,-apple-system,Segoe UI,Roboto,Arial,sans-serif;margin:40px}
  .card{max-width:980px;margin:0 auto;border:1px solid #e5e7eb;border-radius:16px;padding:24px;box-shadow:0 6px 24px rgba(0,0,0,.06);background:#fff}
  h1{margin:0 0 8px}.muted{color:#6b7280}
  label{font-weight:600;display:block;margin-top:16px}
  input[type=file],input[type=url],input[type=number]{width:100%;padding:12px;border:1px solid #d1d5db;border-radius:10px}
  .row{display:flex;gap:16px;flex-wrap:wrap}.row>div{flex:1 1 300px}
  button,.btn{margin-top:16px;padding:12px 18px;border:0;background:#111827;color:#fff;border-radius:10px;cursor:pointer;text-decoration:none;display:inline-block}
  .ok{color:#059669}.warn{color:#b45309}.err{color:#dc2626}
  table{border-collapse:collapse;font-size:12px}td,th{border:1px solid #e5e7eb;padding:4px 6px}
  code{background:#f3f4f6;padding:2px 6px;border-radius:6px}
</style>
<div class="card">
  <h1>Генерация DOCX → ZIP</h1>
  <p class="muted">Формат по умолчанию — «широкая» таблица: <b>строка 1 — заголовки</b>, <b>строка 2 — значения</b>.
  Если встретится «шпаргалка» вида <code>output_mask</code>, сервис сам переключится в KV-режим (строка 1 — ключи, строка 2 — значения).</p>

  <form id="f" action="/generate" method="post" enctype="multipart/form-data">
    <div class="row">
      <div>
        <label>Excel (.xlsx) или CSV</label>
        <!-- БЕЗ required: можно работать только по ссылке -->
        <input type="file" name="table_file" id="table_file" accept=".xlsx,.csv">
      </div>
      <div>
        <label>Google Sheet (опционально)</label>
        <input type="url" name="gsheet_url" id="gsheet_url" placeholder="https://docs.google.com/spreadsheets/d/.../edit#gid=0">
        <div class="muted" style="font-size:12px">Укажите <b>ссылку</b> или выберите <b>файл</b>.</div>
      </div>
    </div>
    <div class="row">
      <div><label>Строка заголовков (для «широкой» таблицы)</label><input type="number" name="header_row" value="1" min="1"></div>
    </div>

    <div class="row">
      <div>
        <button id="btnInspect" type="button">Проверить</button>
        <button id="btnGen" type="submit">Сгенерировать ZIP</button>
        <a class="btn" id="btnTpl" href="#">Скачать шаблон (main_example.xlsx)</a>
      </div>
      <div><span id="spin" style="display:none;margin-left:12px">Обработка…</span></div>
    </div>
  </form>

  <div id="rep" style="margin-top:16px"></div>
</div>
<script>
  const f=document.getElementById('f'), rep=document.getElementById('rep'),
        spin=document.getElementById('spin'), btnI=document.getElementById('btnInspect'),
        btnG=document.getElementById('btnGen'), file=document.getElementById('table_file'),
        gs=document.getElementById('gsheet_url'), btnTpl=document.getElementById('btnTpl');

  // удобство: если вводят ссылку — очищаем файл; если выбирают файл — чистим ссылку
  file.addEventListener('change',()=>{ if(file.files.length) gs.value=''; });
  gs.addEventListener('input',()=>{ if(gs.value.trim()) file.value=''; });

  function needSource(){
    const hasFile = file.files && file.files.length>0;
    const hasGS = !!gs.value.trim();
    if(!hasFile && !hasGS){ alert('Укажите Google Sheet ИЛИ выберите файл'); return false; }
    return true;
  }

  function kvPairs(pairs){
    if(!pairs||!pairs.length) return '';
    const rows = pairs.map(([k,v])=>`<tr><td>${k}</td><td>${v}</td></tr>`).join('');
    return `<details open><summary>Первые пары</summary><table><thead><tr><th>Ключ</th><th>Значение</th></tr></thead><tbody>${rows}</tbody></table></details>`;
  }
  function previewWide(obj){
    if(!obj) return '';
    const rows = Object.entries(obj).map(([k,v])=>`<tr><th>${k}</th><td>${v}</td></tr>`).join('');
    return `<details open><summary>Предпросмотр записи</summary><table><tbody>${rows}</tbody></table></details>`;
  }

  async function inspect(){
    if(!needSource()) return;
    rep.innerHTML=''; spin.style.display='inline';
    const fd=new FormData(f); const r=await fetch('/inspect',{method:'POST',body:fd}); const d=await r.json();
    spin.style.display='none';
    if(!r.ok){ rep.innerHTML='<p class="err">'+(d.detail||'Ошибка')+'</p>'; return; }
    const meta = d.meta ? `<p class="muted">Источник: <b>${d.meta.source}</b>; режим: <b>${d.meta.mode}</b>; header: <b>${(d.meta.header_row??0)+1}</b></p>` : '';
    const miss = d.missing?.length ? '<p class="warn">Отсутствуют ключевые поля: '+d.missing.join(', ')+'</p>' : '<p class="ok">Ключевые поля найдены</p>';
    const cols = d.columns?.length ? '<p><b>Колонки:</b> '+d.columns.join(', ')+'</p>' : '';
    const body = d.meta?.mode==='wide' ? previewWide(d.preview) : kvPairs(d.preview_pairs);
    rep.innerHTML = meta + cols + miss + body;
  }
  btnI.addEventListener('click', inspect);

  f.addEventListener('submit', async (e)=>{
    e.preventDefault();
    if(!needSource()) return;
    rep.innerHTML=''; btnI.disabled=btnG.disabled=true; spin.style.display='inline';
    try{
      const fd=new FormData(f); const r=await fetch('/generate',{method:'POST',body:fd}); const b=await r.blob();
      if(!r.ok){ rep.innerHTML='<p class="err">'+await b.text()+'</p>'; return; }
      const url=URL.createObjectURL(b); const a=document.createElement('a'); a.href=url; a.download='generated_docs.zip';
      document.body.appendChild(a); a.click(); a.remove(); URL.revokeObjectURL(url);
      rep.innerHTML='<p class="ok">Готово: ZIP скачан</p>';
    } finally { btnI.disabled=btnG.disabled=false; spin.style.display='none'; }
  });

  // Скачивание шаблона через fetch — исключаем влияние формы/кэша браузера
  btnTpl.addEventListener('click', async (e)=>{
    e.preventDefault();
    try{
      const r = await fetch('/template', { method:'GET', cache:'no-store' });
      if(!r.ok){ const t=await r.text(); alert('Не удалось скачать шаблон: '+t); return; }
      const b = await r.blob();
      const url = URL.createObjectURL(b);
      const a = document.createElement('a'); a.href = url; a.download = 'main_example.xlsx';
      document.body.appendChild(a); a.click(); a.remove(); URL.revokeObjectURL(url);
    }catch(err){ alert('Ошибка скачивания: '+err); }
  });
</script>
"""

# ============= Хелперы =============
INVALID_FS = r'[<>:"/\\|?*]'

def safe(v): return "" if (v is None or pd.isna(v)) else str(v).strip()

class SafeMap(dict):
    def __missing__(self, key): return ""

def slugify(name: str) -> str:
    return re.sub(INVALID_FS, "_", name).rstrip(" .") or "file"

def _norm(s: str) -> str:
    return re.sub(r"\s+", "", str(s)).replace("\ufeff","").replace("\xa0","").replace("ё","е").lower()

def expected_headers() -> set:
    exp = {"фио","группа"}
    for tpl in TEMPLATES:
        exp |= {_norm(v) for v in tpl["fields"].values()}
        exp |= {_norm(m) for m in re.findall(r"\{([^}]+)\}", tpl["out"])}
    return exp

def score_columns(cols) -> int:
    exp = expected_headers()
    return sum(1 for c in cols if _norm(c) in exp)

def read_wide_try(file_bytes: bytes, is_xlsx: bool, header_row: int) -> Tuple[pd.DataFrame, Dict]:
    if is_xlsx:
        df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=0, header=max(header_row-1,0))
        return df, {"source":"xlsx", "mode":"wide", "header_row": header_row-1}
    else:
        sample = file_bytes[:2048].decode("utf-8", errors="ignore")
        try: sep = csv.Sniffer().sniff(sample).delimiter
        except Exception: sep = ","
        df = pd.read_csv(io.BytesIO(file_bytes), sep=sep, header=max(header_row-1,0))
        return df, {"source":"csv", "mode":"wide", "header_row": header_row-1}

def read_kv_from_raw(file_bytes: bytes, is_xlsx: bool, key_row: int = 1, val_row: int = 2) -> Tuple[Dict[str,str], Dict]:
    if is_xlsx:
        df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=0, header=None)
    else:
        df = pd.read_csv(io.BytesIO(file_bytes), header=None)
    keys = [safe(x).replace("\ufeff","").replace("\xa0"," ") for x in df.iloc[key_row-1].tolist()]
    vals = [safe(x).replace("\ufeff","").replace("\xa0"," ") for x in df.iloc[val_row-1].tolist()]
    kv = {k: v for k, v in zip(keys, vals) if k}
    return kv, {"source":"xlsx" if is_xlsx else "csv", "mode":"kv", "key_row":key_row-1, "val_row":val_row-1}

def extract_record_from_upload(file: UploadFile, header_row: int) -> Tuple[Dict[str,str], Dict, Optional[list]]:
    data = file.file.read()
    name = (file.filename or "").lower()
    is_xlsx = name.endswith(".xlsx")
    if not (is_xlsx or name.endswith(".csv")):
        raise HTTPException(400, "Поддерживаются только .xlsx или .csv")

    df_wide, meta = read_wide_try(data, is_xlsx, header_row)
    if not df_wide.empty:
        cols = [str(c) for c in df_wide.columns]
        sc = score_columns(cols)
        if sc >= 3:
            row = pick_first_nonempty_row(df_wide)
            row_dict = {str(k): safe(v) for k, v in row.items()}
            meta.update({"mode":"wide", "score": sc})
            return row_dict, meta, cols

    kv, meta_kv = read_kv_from_raw(data, is_xlsx, 1, 2)
    meta_kv.setdefault("score", 0)
    return kv, meta_kv, None

def extract_record_from_gsheet(url: str, header_row: int) -> Tuple[Dict[str,str], Dict, Optional[list]]:
    m = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", url or "")
    if not m: raise HTTPException(400, "Не удалось извлечь spreadsheetId из URL")
    spreadsheet_id = m.group(1)
    gid_match = re.search(r"[#&?]gid=([0-9]+)", url)
    gid = int(gid_match.group(1)) if gid_match else 0
    export = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/export?format=csv&gid={gid}"
    resp = requests.get(export, timeout=30)
    if resp.status_code != 200:
        raise HTTPException(400, f"Google Sheets недоступен (HTTP {resp.status_code})")
    upl = UploadFile(filename="gs.csv", file=io.BytesIO(resp.content))
    rec, meta, cols = extract_record_from_upload(upl, header_row)
    meta.update({"source":"gsheet", "gid": gid})
    return rec, meta, cols

def pick_first_nonempty_row(df: pd.DataFrame) -> pd.Series:
    df = df.fillna("")
    for _, row in df.iterrows():
        if any(safe(v) for v in row.values):
            return row
    raise HTTPException(400, "Не найдена ни одна непустая строка с данными")

# -------- шаблон Excel --------
@app.get("/template")
def download_template():
    """
    Отдаём шаблон main_example.xlsx из корня проекта (BASE_DIR).
    Если не найден — генерируем минимальный шаблон из полей templates_config.py.
    Во всех случаях имя скачивания = main_example.xlsx. Кэш выключен.
    """
    for p in CANDIDATE_TEMPLATE_PATHS:
        if p.exists():
            return FileResponse(
                str(p),
                media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                filename=TEMPLATE_DOWNLOAD_NAME,
                headers={"Cache-Control": "no-store, no-cache, must-revalidate"},
            )

    # Fallback: сформируем шаблон автоматически из полей templates_config.py
    headers_list: List[str] = []
    seen = set()
    for tpl in TEMPLATES:
        for col in tpl["fields"].values():
            if col not in seen:
                seen.add(col)
                headers_list.append(col)

    df = pd.DataFrame([[""] * len(headers_list)], columns=headers_list)  # одна пустая строка
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df.to_excel(w, sheet_name="Sheet1", index=False)
    buf.seek(0)
    return StreamingResponse(
        buf,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={
            "Content-Disposition": f'attachment; filename="{TEMPLATE_DOWNLOAD_NAME}"',
            "Cache-Control": "no-store, no-cache, must-revalidate",
        },
    )

# ============= HTTP API =============
@app.get("/", response_class=HTMLResponse)
def index():
    return HTMLResponse(INDEX_HTML)

@app.post("/inspect")
def inspect(
    table_file: Optional[UploadFile] = File(default=None),
    gsheet_url: Optional[str] = Form(default=None),
    header_row: int = Form(default=1),
):
    # приоритет: если есть ссылка — используем её, иначе файл
    if gsheet_url and gsheet_url.strip():
        record, meta, cols = extract_record_from_gsheet(gsheet_url.strip(), header_row)
    elif table_file and (table_file.filename or "").strip():
        record, meta, cols = extract_record_from_upload(table_file, header_row)
    else:
        raise HTTPException(400, "Укажите Google Sheet ИЛИ выберите файл")

    needed = ["ФИО", "Группа"]
    missing = [k for k in needed if k not in record]

    if meta["mode"] == "wide":
        preview = record
        return JSONResponse({"columns": cols or [], "preview": preview, "missing": missing, "meta": meta})
    else:
        preview_pairs = list(record.items())[:12]
        return JSONResponse({"columns": [], "preview_pairs": preview_pairs, "missing": missing, "meta": meta})

@app.post("/generate")
def generate_zip(
    table_file: Optional[UploadFile] = File(default=None),
    gsheet_url: Optional[str] = Form(default=None),
    header_row: int = Form(default=1),
):
    if gsheet_url and gsheet_url.strip():
        record, meta, _ = extract_record_from_gsheet(gsheet_url.strip(), header_row)
    elif table_file and (table_file.filename or "").strip():
        record, meta, _ = extract_record_from_upload(table_file, header_row)
    else:
        raise HTTPException(400, "Укажите Google Sheet ИЛИ выберите файл")

    fio = safe(record.get("ФИО")) or "record"
    folder = slugify(f"001_{fio}")

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for tpl in TEMPLATES:
            try:
                # контекст: подставляем по маппингу {tpl_key: excel_col}
                ctx = {tpl_key: safe(record.get(excel_col, "")) for tpl_key, excel_col in tpl["fields"].items()}
                doc = DocxTemplate(tpl["path"])
                doc.render(ctx)

                # имя файла из шаблонной маски out
                out_name = tpl["out"].format_map(SafeMap(record))
                out_name = slugify(out_name) or "doc_001.docx"

                out_mem = io.BytesIO(); doc.save(out_mem)
                zf.writestr(f"{folder}/{out_name}", out_mem.getvalue())
            except Exception as e:
                err = slugify(tpl.get("out","file")) + ".ERROR.txt"
                zf.writestr(f"{folder}/{err}", f"Ошибка ({tpl['path']}): {type(e).__name__}: {e}")

    buf.seek(0)
    return StreamingResponse(
        buf,
        media_type="application/zip",
        headers={"Content-Disposition": 'attachment; filename="generated_docs.zip"'}
    )

@app.get("/healthz")
def healthz():
    return PlainTextResponse("ok")
