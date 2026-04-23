
import json
import os
import re
from collections import defaultdict
from datetime import datetime
from pathlib import Path

from flask import Flask, render_template, request, redirect, url_for, flash, abort
from werkzeug.utils import secure_filename
from openpyxl import load_workbook

APP_DIR = Path(__file__).resolve().parent
DATA_DIR = Path(os.environ.get("DATA_DIR", APP_DIR / "instance" / "data"))
DATA_DIR.mkdir(parents=True, exist_ok=True)

BASE_FILE = DATA_DIR / "base_atual.xlsx"
META_FILE = DATA_DIR / "creditos_meta.json"

ALLOWED_EXTENSIONS = {".xlsx", ".xlsm"}

app = Flask(__name__, instance_relative_config=True)
app.secret_key = os.environ.get("SECRET_KEY", "troque-esta-chave-em-producao")
app.config["MAX_CONTENT_LENGTH"] = 80 * 1024 * 1024  # 80MB


def allowed_file(filename: str) -> bool:
    return Path(filename.lower()).suffix in ALLOWED_EXTENSIONS


def clean_cell(value):
    if value is None:
        return ""
    return str(value).strip()


def normalize_rows(rows):
    return [[clean_cell(v) for v in row] for row in rows]


def normalize_header_name(text: str) -> str:
    t = clean_cell(text).upper()
    t = t.replace("Nº", "NO").replace("N°", "NO")
    t = re.sub(r"[^A-Z0-9 ]+", " ", t)
    t = re.sub(r"\s+", " ", t).strip()
    return t


def detect_header_and_rows(rows):
    rows = normalize_rows(rows)
    best_idx = 0
    best_score = -1

    for idx, row in enumerate(rows[:15]):
        normalized = [normalize_header_name(c) for c in row]
        keywords = ["RECIBO", "SETOR", "STATUS", "EXAME", "FUNCIONARIO", "DATA", "TIPO"]
        keyword_hits = sum(1 for c in normalized if any(k in c for k in keywords))
        non_empty = sum(1 for c in row if c)
        score = keyword_hits * 8 + non_empty
        if score > best_score:
            best_score = score
            best_idx = idx

    header = rows[best_idx] if rows else []
    data = rows[best_idx + 1:] if best_idx + 1 < len(rows) else []
    while data and not any(data[-1]):
        data.pop()

    return header, data, best_idx + 1


def find_column_index(header, aliases, fallback_index=None):
    normalized = [normalize_header_name(c) for c in header]

    for alias in aliases:
        alias_norm = normalize_header_name(alias)
        for idx, col in enumerate(normalized):
            if alias_norm and alias_norm in col:
                return idx

    if fallback_index is not None and fallback_index < len(header):
        return fallback_index
    return None


def parse_setor(setor_text):
    raw = clean_cell(setor_text)
    if not raw:
        return "", ""

    # tenta CNPJ/CPF no final ou no meio
    doc_match = re.search(r'(\d{2}\.?\d{3}\.?\d{3}/?\d{4}-?\d{2}|\d{3}\.?\d{3}\.?\d{3}-?\d{2})', raw)
    doc = doc_match.group(1) if doc_match else ""
    doc = format_document(doc)

    empresa = raw
    if doc_match:
        empresa = (raw[:doc_match.start()] + raw[doc_match.end():]).strip(" -–•|")
    empresa = re.sub(r"\s+", " ", empresa).strip(" -–•|")
    return empresa, doc


def format_document(doc):
    digits = re.sub(r"\D", "", doc or "")
    if len(digits) == 14:
        return f"{digits[:2]}.{digits[2:5]}.{digits[5:8]}/{digits[8:12]}-{digits[12:]}"
    if len(digits) == 11:
        return f"{digits[:3]}.{digits[3:6]}.{digits[6:9]}-{digits[9:]}"
    return clean_cell(doc)


def detect_credit_status(row, status_idx_primary, status_idx_secondary):
    values_to_check = []
    if status_idx_primary is not None and status_idx_primary < len(row):
        values_to_check.append(clean_cell(row[status_idx_primary]).upper())
    if status_idx_secondary is not None and status_idx_secondary < len(row):
        values_to_check.append(clean_cell(row[status_idx_secondary]).upper())

    return any(v == "NÃO REALIZADO" or v == "NAO REALIZADO" for v in values_to_check)


def process_workbook(path: Path):
    wb = load_workbook(path, data_only=True)
    months = []
    total_creditos = 0

    for ws in wb.worksheets:
        raw_rows = list(ws.iter_rows(values_only=True))
        header, data, header_row_number = detect_header_and_rows(raw_rows)
        if not header:
            continue

        # tenta achar por nome; senão usa os fallbacks do seu arquivo
        recibo_idx = find_column_index(header, ["RECIBO", "NO RECIBO", "NUMERO RECIBO"], fallback_index=2)  # C
        setor_idx = find_column_index(header, ["SETOR"], fallback_index=11)  # L
        status_h_idx = 7 if len(header) > 7 else None  # H
        status_d_idx = 3 if len(header) > 3 else None  # D

        grouped = defaultdict(lambda: {"credito_count": 0, "recibo": "", "empresa": "", "documento": ""})

        for row in normalize_rows(data):
            if not any(row):
                continue

            if not detect_credit_status(row, status_h_idx, status_d_idx):
                continue

            recibo = clean_cell(row[recibo_idx]) if recibo_idx is not None and recibo_idx < len(row) else ""
            setor = clean_cell(row[setor_idx]) if setor_idx is not None and setor_idx < len(row) else ""
            empresa, documento = parse_setor(setor)

            key = (recibo, empresa, documento)
            grouped[key]["credito_count"] += 1
            grouped[key]["recibo"] = recibo
            grouped[key]["empresa"] = empresa
            grouped[key]["documento"] = documento

        creditos = list(grouped.values())
        creditos.sort(key=lambda x: (x["empresa"], x["recibo"]))

        total_creditos += sum(item["credito_count"] for item in creditos)

        months.append({
            "sheet_name": ws.title,
            "header_row_number": header_row_number,
            "row_count": len([r for r in data if any(clean_cell(c) for c in r)]),
            "credito_receipts_count": len(creditos),
            "credito_total_count": sum(item["credito_count"] for item in creditos),
            "creditos": creditos,
        })

    meta = {
        "filename": path.name,
        "original_filename": path.name,
        "updated_at": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
        "sheet_count": len(months),
        "credito_total_count": total_creditos,
        "months": months,
    }
    return meta


def save_meta(meta: dict):
    META_FILE.write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")


def load_meta():
    if not META_FILE.exists():
        return None
    return json.loads(META_FILE.read_text(encoding="utf-8"))


@app.route("/")
def index():
    meta = load_meta()
    return render_template("index.html", meta=meta)


@app.route("/upload-base", methods=["POST"])
def upload_base():
    file = request.files.get("base_file")
    if not file or not file.filename:
        flash("Selecione uma planilha .xlsx ou .xlsm.")
        return redirect(url_for("index"))

    if not allowed_file(file.filename):
        flash("Formato inválido. Envie apenas .xlsx ou .xlsm.")
        return redirect(url_for("index"))

    temp_name = secure_filename(file.filename)
    temp_path = DATA_DIR / f"_tmp_{temp_name}"
    file.save(temp_path)

    try:
        meta = process_workbook(temp_path)
    except Exception as e:
        temp_path.unlink(missing_ok=True)
        flash(f"Não foi possível ler a planilha: {e}")
        return redirect(url_for("index"))

    BASE_FILE.unlink(missing_ok=True)
    temp_path.replace(BASE_FILE)
    meta["original_filename"] = file.filename
    meta["filename"] = BASE_FILE.name
    save_meta(meta)

    flash("Base atualizada com sucesso.")
    return redirect(url_for("index"))


@app.route("/mes/<sheet_name>")
def view_month(sheet_name):
    meta = load_meta()
    if not meta:
        abort(404)

    month = next((m for m in meta["months"] if m["sheet_name"] == sheet_name), None)
    if not month:
        abort(404)

    return render_template("month.html", meta=meta, month=month)


@app.route("/healthz")
def healthz():
    return {"status": "ok"}
