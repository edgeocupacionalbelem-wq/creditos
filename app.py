
import json
import os
import re
from datetime import datetime
from pathlib import Path

from flask import Flask, render_template, request, redirect, url_for, flash, abort
from werkzeug.utils import secure_filename
from openpyxl import load_workbook

APP_DIR = Path(__file__).resolve().parent
DATA_DIR = Path(os.environ.get("DATA_DIR", APP_DIR / "instance" / "data"))
DATA_DIR.mkdir(parents=True, exist_ok=True)

BASE_FILE = DATA_DIR / "base_atual.xlsx"
META_FILE = DATA_DIR / "base_meta.json"

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
    cleaned = []
    for row in rows:
        cleaned.append([clean_cell(v) for v in row])
    return cleaned


def detect_header_and_rows(rows):
    rows = normalize_rows(rows)
    best_idx = 0
    best_score = -1

    for idx, row in enumerate(rows[:15]):
        non_empty = sum(1 for c in row if c)
        long_cells = sum(1 for c in row if len(c) > 2)
        score = non_empty * 2 + long_cells
        if score > best_score:
            best_score = score
            best_idx = idx

    header = rows[best_idx] if rows else []
    data = rows[best_idx + 1:] if best_idx + 1 < len(rows) else []

    while data and not any(data[-1]):
        data.pop()

    return header, data, best_idx + 1


def workbook_to_metadata(path: Path):
    wb = load_workbook(path, data_only=True)
    sheets = []
    total_rows = 0
    total_cols = 0

    for ws in wb.worksheets:
        raw_rows = list(ws.iter_rows(values_only=True))
        header, data, header_row_number = detect_header_and_rows(raw_rows)

        non_empty_data = [row for row in data if any(cell for cell in row)]
        preview = non_empty_data[:100]

        max_cols = max((len(header), *(len(r) for r in preview)), default=0)
        header = header[:max_cols]
        preview = [r[:max_cols] + [""] * (max_cols - len(r[:max_cols])) for r in preview]

        total_rows += len(non_empty_data)
        total_cols = max(total_cols, max_cols)

        sheets.append({
            "name": ws.title,
            "header_row_number": header_row_number,
            "columns": header,
            "row_count": len(non_empty_data),
            "preview_rows": preview,
        })

    meta = {
        "filename": path.name,
        "original_filename": path.name,
        "updated_at": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
        "sheet_count": len(sheets),
        "total_rows": total_rows,
        "max_columns": total_cols,
        "sheets": sheets,
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
        meta = workbook_to_metadata(temp_path)
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


@app.route("/sheet/<sheet_name>")
def view_sheet(sheet_name):
    meta = load_meta()
    if not meta:
        abort(404)

    sheet = next((s for s in meta["sheets"] if s["name"] == sheet_name), None)
    if not sheet:
        abort(404)

    return render_template("sheet.html", meta=meta, sheet=sheet)


@app.route("/healthz")
def healthz():
    return {"status": "ok"}
