
import os
import re
from collections import defaultdict
from datetime import datetime

from flask import Flask, render_template, request, redirect, url_for, flash
from werkzeug.utils import secure_filename
from openpyxl import load_workbook
from flask_sqlalchemy import SQLAlchemy
from sqlalchemy import func, or_

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "troque-esta-chave")
app.config["MAX_CONTENT_LENGTH"] = 80 * 1024 * 1024

db_url = os.environ.get("DATABASE_URL", "sqlite:///creditos.db")
if db_url.startswith("postgres://"):
    db_url = db_url.replace("postgres://", "postgresql://", 1)

app.config["SQLALCHEMY_DATABASE_URI"] = db_url
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False

db = SQLAlchemy(app)

ALLOWED_EXTENSIONS = {".xlsx", ".xlsm"}

class BaseMetadata(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    original_filename = db.Column(db.String(255), nullable=False)
    updated_at = db.Column(db.String(30), nullable=False)
    months_count = db.Column(db.Integer, nullable=False, default=0)
    total_creditos = db.Column(db.Integer, nullable=False, default=0)

class Credito(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    mes_index = db.Column(db.Integer, nullable=False)
    mes_nome = db.Column(db.String(120), nullable=False, index=True)
    recibo = db.Column(db.String(120), nullable=True, index=True)
    empresa = db.Column(db.String(255), nullable=True, index=True)
    documento = db.Column(db.String(30), nullable=True)
    credito_count = db.Column(db.Integer, nullable=False, default=1)

with app.app_context():
    db.create_all()

def allowed_file(filename: str) -> bool:
    return os.path.splitext(filename.lower())[1] in ALLOWED_EXTENSIONS

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

def format_document(doc):
    digits = re.sub(r"\D", "", doc or "")
    if len(digits) == 14:
        return f"{digits[:2]}.{digits[2:5]}.{digits[5:8]}/{digits[8:12]}-{digits[12:]}"
    if len(digits) == 11:
        return f"{digits[:3]}.{digits[3:6]}.{digits[6:9]}-{digits[9:]}"
    return clean_cell(doc)

def parse_setor(setor_text):
    raw = clean_cell(setor_text)
    if not raw:
        return "", ""
    doc_match = re.search(r'(\d{2}\.?\d{3}\.?\d{3}/?\d{4}-?\d{2}|\d{3}\.?\d{3}\.?\d{3}-?\d{2})', raw)
    doc = doc_match.group(1) if doc_match else ""
    doc = format_document(doc)
    empresa = raw
    if doc_match:
        empresa = (raw[:doc_match.start()] + raw[doc_match.end():]).strip(" -–•|")
    empresa = re.sub(r"\s+", " ", empresa).strip(" -–•|")
    return empresa, doc

def detect_credit_status(row, status_idx_primary, status_idx_secondary):
    values = []
    if status_idx_primary is not None and status_idx_primary < len(row):
        values.append(clean_cell(row[status_idx_primary]).upper())
    if status_idx_secondary is not None and status_idx_secondary < len(row):
        values.append(clean_cell(row[status_idx_secondary]).upper())
    return any(v == "NÃO REALIZADO" or v == "NAO REALIZADO" for v in values)

def process_workbook(filepath):
    wb = load_workbook(filepath, data_only=True)
    months = []
    total_creditos = 0

    for ws_idx, ws in enumerate(wb.worksheets, start=1):
        raw_rows = list(ws.iter_rows(values_only=True))
        header, data, _ = detect_header_and_rows(raw_rows)
        if not header:
            continue

        recibo_idx = find_column_index(header, ["RECIBO", "NO RECIBO", "NUMERO RECIBO"], fallback_index=2)
        setor_idx = find_column_index(header, ["SETOR"], fallback_index=11)
        status_h_idx = 7 if len(header) > 7 else None
        status_d_idx = 3 if len(header) > 3 else None

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
        mes_total = sum(item["credito_count"] for item in creditos)
        total_creditos += mes_total

        months.append({
            "mes_index": ws_idx,
            "mes_nome": ws.title,
            "credito_receipts_count": len(creditos),
            "credito_total_count": mes_total,
            "creditos": creditos,
        })

    return months, total_creditos

def replace_database_with_planilha(filepath, original_filename):
    months, total_creditos = process_workbook(filepath)

    db.session.query(Credito).delete()
    db.session.query(BaseMetadata).delete()

    meta = BaseMetadata(
        original_filename=original_filename,
        updated_at=datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
        months_count=len(months),
        total_creditos=total_creditos,
    )
    db.session.add(meta)

    for month in months:
        for item in month["creditos"]:
            db.session.add(Credito(
                mes_index=month["mes_index"],
                mes_nome=month["mes_nome"],
                recibo=item["recibo"],
                empresa=item["empresa"],
                documento=item["documento"],
                credito_count=item["credito_count"],
            ))

    db.session.commit()

def get_dashboard_data(start_idx=None, end_idx=None, busca=""):
    meta = BaseMetadata.query.order_by(BaseMetadata.id.desc()).first()

    month_rows = (
        db.session.query(
            Credito.mes_index,
            Credito.mes_nome,
            func.count(Credito.id).label("credito_receipts_count"),
            func.coalesce(func.sum(Credito.credito_count), 0).label("credito_total_count"),
        )
        .group_by(Credito.mes_index, Credito.mes_nome)
        .order_by(Credito.mes_index.asc())
        .all()
    )

    months_summary = []
    for row in month_rows:
        months_summary.append({
            "mes_index": row.mes_index,
            "mes_nome": row.mes_nome,
            "credito_receipts_count": int(row.credito_receipts_count or 0),
            "credito_total_count": int(row.credito_total_count or 0),
        })

    if months_summary:
        min_idx = months_summary[0]["mes_index"]
        max_idx = months_summary[-1]["mes_index"]
    else:
        min_idx = 1
        max_idx = 1

    if start_idx is None:
        start_idx = min_idx
    if end_idx is None:
        end_idx = max_idx
    if start_idx > end_idx:
        start_idx, end_idx = end_idx, start_idx

    q = Credito.query.filter(Credito.mes_index >= start_idx, Credito.mes_index <= end_idx)
    busca = (busca or "").strip()
    if busca:
        term = f"%{busca}%"
        q = q.filter(
            or_(
                Credito.mes_nome.ilike(term),
                Credito.recibo.ilike(term),
                Credito.empresa.ilike(term),
                Credito.documento.ilike(term),
            )
        )

    creditos = q.order_by(Credito.mes_index.asc(), Credito.empresa.asc(), Credito.recibo.asc()).all()

    months_map = {m["mes_index"]: {**m, "creditos": []} for m in months_summary}
    for c in creditos:
        if c.mes_index in months_map:
            months_map[c.mes_index]["creditos"].append(c)

    months_filtered = [months_map[idx] for idx in sorted(months_map) if start_idx <= idx <= end_idx]
    resumo_total = sum(m["credito_total_count"] for m in months_filtered)
    resumo_recibos = sum(m["credito_receipts_count"] for m in months_filtered)

    return meta, months_summary, months_filtered, resumo_total, resumo_recibos, min_idx, max_idx, start_idx, end_idx, busca

@app.route("/")
def index():
    try:
        inicio = int(request.args.get("inicio", "0")) or None
    except ValueError:
        inicio = None
    try:
        fim = int(request.args.get("fim", "0")) or None
    except ValueError:
        fim = None
    busca = request.args.get("busca", "")

    meta, months_summary, months_filtered, resumo_total, resumo_recibos, min_idx, max_idx, start_idx, end_idx, busca = get_dashboard_data(inicio, fim, busca)

    return render_template(
        "index.html",
        meta=meta,
        months_summary=months_summary,
        months_filtered=months_filtered,
        resumo_total=resumo_total,
        resumo_recibos=resumo_recibos,
        min_idx=min_idx,
        max_idx=max_idx,
        inicio=start_idx,
        fim=end_idx,
        busca=busca,
    )

@app.route("/upload-base", methods=["POST"])
def upload_base():
    file = request.files.get("base_file")
    if not file or not file.filename:
        flash("Selecione uma planilha .xlsx ou .xlsm.")
        return redirect(url_for("index"))
    if not allowed_file(file.filename):
        flash("Formato inválido. Envie apenas .xlsx ou .xlsm.")
        return redirect(url_for("index"))

    filename = secure_filename(file.filename)
    temp_dir = os.environ.get("TMPDIR", "/tmp")
    os.makedirs(temp_dir, exist_ok=True)
    temp_path = os.path.join(temp_dir, filename)
    file.save(temp_path)

    try:
        replace_database_with_planilha(temp_path, file.filename)
        flash("Base atualizada com sucesso.")
    except Exception as e:
        flash(f"Não foi possível ler a planilha: {e}")
    finally:
        try:
            os.remove(temp_path)
        except OSError:
            pass

    return redirect(url_for("index"))

@app.route("/healthz")
def healthz():
    return {"status": "ok"}
