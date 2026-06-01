"""
Sweety → GTG  |  Web App mensile
"""
import os
import json
import uuid
import shutil
from pathlib import Path
from datetime import datetime, timedelta
from flask import (Flask, render_template, request, redirect,
                   url_for, send_file, abort, session)

import pipeline as pl
import docgen
import emailer

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "sweety-secret-2026")

ACCESS_CODE = os.environ.get("ACCESS_CODE", "sweety2026")
SESSIONS_DIR = Path("sessions")
SESSIONS_DIR.mkdir(exist_ok=True)

# Dati aziendali (configurabili via env)
SWEETY = {
    "name": os.environ.get("SWEETY_NAME", "Sweety Pact S.r.l."),
    "address_line1": os.environ.get("SWEETY_ADDR1", "MD-2025, str. Acad. Natalia Gheorghiu 30, of. 262, mun."),
    "address_line2": os.environ.get("SWEETY_ADDR2", "Chisinau, RM"),
    "cod_fiscal": os.environ.get("SWEETY_CF", "1025600056490"),
    "email": os.environ.get("SWEETY_EMAIL", "info@sweetypact.com"),
    "phone": os.environ.get("SWEETY_PHONE", "+37360045404"),
    "bank_name": os.environ.get("SWEETY_BANK", "BC MOLDINDCONBANK SA"),
    "bank_bic": os.environ.get("SWEETY_BIC", "MOLDMD2X(XXX)"),
    "bank_iban": os.environ.get("SWEETY_IBAN", "MD54ML022510000000006428"),
}
GTG = {
    "name": os.environ.get("GTG_NAME", "Good to Great S.r.l."),
    "cod_fiscal": os.environ.get("GTG_CF", "48911391"),
    "address": os.environ.get("GTG_ADDRESS", "România, comuna Arad, str. Poetului, nr.1/c, hala 21"),
}
ADMIN_CONTACT = {
    "name": os.environ.get("ADMIN_NAME", "Adrian"),
    "email": os.environ.get("ADMIN_EMAIL", ""),
}


_MONTHS_IT = ["","Gennaio","Febbraio","Marzo","Aprile","Maggio","Giugno",
              "Luglio","Agosto","Settembre","Ottobre","Novembre","Dicembre"]


def _billing_months(n: int = 6):
    """Restituisce gli ultimi n mesi come lista di (valore, etichetta)."""
    result = []
    today = datetime.today()
    year, month = today.year, today.month
    for _ in range(n):
        result.append((f"{year}-{month:02d}", f"{_MONTHS_IT[month]} {year}"))
        month -= 1
        if month == 0:
            month = 12
            year -= 1
    return result


# ── Helpers ────────────────────────────────────────────────────────────────────

def _session_dir(sid: str) -> Path:
    return SESSIONS_DIR / sid


def _save_session(sid: str, data: dict) -> None:
    p = _session_dir(sid)
    p.mkdir(exist_ok=True)
    with open(p / "data.json", "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, default=str)


def _load_session(sid: str):
    try:
        with open(_session_dir(sid) / "data.json", encoding="utf-8") as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return None


def _save_file(sid: str, name: str, data: bytes) -> None:
    (_session_dir(sid) / name).write_bytes(data)


def _load_file(sid: str, name: str):
    try:
        return (_session_dir(sid) / name).read_bytes()
    except FileNotFoundError:
        return None


def _cleanup_old_sessions() -> None:
    """Elimina sessioni più vecchie di 48 ore."""
    cutoff = datetime.now() - timedelta(hours=48)
    for p in SESSIONS_DIR.iterdir():
        if p.is_dir():
            try:
                mtime = datetime.fromtimestamp(p.stat().st_mtime)
                if mtime < cutoff:
                    shutil.rmtree(p)
            except Exception:
                pass


def _check_access() -> bool:
    code = request.args.get("code") or request.form.get("code") or session.get("code")
    if code == ACCESS_CODE:
        session["code"] = code
        return True
    return False


def _url(endpoint: str, **kwargs):
    kwargs["code"] = ACCESS_CODE
    return url_for(endpoint, **kwargs)


# ── Routes ────────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    if not _check_access():
        return render_template("login.html")
    _cleanup_old_sessions()
    return render_template("upload.html", code=ACCESS_CODE,
                           billing_months=_billing_months())


@app.route("/upload", methods=["POST"])
def upload():
    if not _check_access():
        return redirect(url_for("index"))

    gtg_file = request.files.get("gtg_csv")
    storico_file = request.files.get("storico_csv")

    if not gtg_file or not gtg_file.filename:
        return render_template("upload.html", code=ACCESS_CODE,
                               billing_months=_billing_months(),
                               error="Carica il report GTG mensile (obbligatorio).")

    billing_period = request.form.get("billing_period", "").strip()
    gtg_bytes = gtg_file.read()
    storico_bytes = storico_file.read() if (storico_file and storico_file.filename) else None

    try:
        result = pl.process(gtg_bytes, gtg_file.filename, storico_bytes, billing_period)
    except Exception as e:
        return render_template("upload.html", code=ACCESS_CODE,
                               billing_months=_billing_months(),
                               error=f"Errore nel file: {e}")

    sid = str(uuid.uuid4())
    _save_session(sid, result)
    if storico_bytes:
        _save_file(sid, "storico.csv", storico_bytes)

    return redirect(_url("review", sid=sid))


@app.route("/review/<sid>")
def review(sid: str):
    if not _check_access():
        return redirect(url_for("index"))
    data = _load_session(sid)
    if not data:
        return redirect(_url("index"))

    # Suggerisci numero fattura (incrementa automaticamente)
    today = datetime.today()
    suggested_sales = f"{today.month + 1}/{today.year}"
    suggested_admin = f"ADM-{today.month + 1:02d}/{today.year}"

    return render_template("review.html",
                           sid=sid,
                           code=ACCESS_CODE,
                           data=data,
                           invoice_sales=suggested_sales,
                           invoice_admin=suggested_admin)


@app.route("/confirm/<sid>", methods=["POST"])
def confirm(sid: str):
    if not _check_access():
        return redirect(url_for("index"))
    data = _load_session(sid)
    if not data:
        return redirect(_url("index"))

    # Leggi selezioni dal form
    selected_names = set(request.form.getlist("selected"))
    invoice_sales = request.form.get("invoice_sales", "").strip()
    invoice_admin = request.form.get("invoice_admin", "").strip()

    # Clienti confermati
    confirmed = [c for c in data["candidates"] if c["nome"] in selected_names]
    totals = pl.compute_totals(data["candidates"], selected_names)

    # Genera documenti
    docs = docgen.generate_all(
        confirmed, totals, data["period"],
        invoice_sales, invoice_admin,
        SWEETY, GTG, ADMIN_CONTACT
    )

    # Genera master aggiornato
    storico_bytes = _load_file(sid, "storico.csv")
    master_bytes = pl.generate_updated_master(storico_bytes, confirmed, data["period"])

    # Salva tutto in sessione
    for fname, fbytes in docs.items():
        _save_file(sid, fname, fbytes)
    _save_file(sid, "master_aggiornato.csv", master_bytes)

    confirm_data = {
        "confirmed": confirmed,
        "totals": totals,
        "invoice_sales": invoice_sales,
        "invoice_admin": invoice_admin,
        "period": data["period"],
        "docs": list(docs.keys()) + ["master_aggiornato.csv"],
    }
    _save_session(sid, {**data, **confirm_data, "_confirmed": True})

    # Invia email
    email_error = None
    try:
        period_label = docgen._period_label_en(data["period"])
        emailer.send_confirmation(
            period_label, totals, confirmed, docs, master_bytes,
            invoice_sales, invoice_admin)
    except Exception as e:
        email_error = str(e)

    return redirect(_url("done", sid=sid, email_error=email_error or ""))


@app.route("/done/<sid>")
def done(sid: str):
    if not _check_access():
        return redirect(url_for("index"))
    data = _load_session(sid)
    if not data or not data.get("_confirmed"):
        return redirect(_url("index"))

    email_error = request.args.get("email_error", "")
    return render_template("done.html",
                           sid=sid,
                           code=ACCESS_CODE,
                           data=data,
                           email_error=email_error)


@app.route("/download/<sid>/<filename>")
def download(sid: str, filename: str):
    if not _check_access():
        abort(403)
    # Sicurezza: solo nomi file semplici
    if "/" in filename or ".." in filename:
        abort(400)
    path = _session_dir(sid) / filename
    if not path.exists():
        abort(404)
    return send_file(path, as_attachment=True, download_name=filename)


if __name__ == "__main__":
    app.run(debug=True)
