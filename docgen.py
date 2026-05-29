"""
Generazione documenti in memoria:
- Lead Register (Excel)
- Activity Summary (Word)
- Bozza email Sales (txt)
- Bozza email Admin (txt)
"""
import io
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH


def _thin_border():
    s = Side(style="thin", color="CCCCCC")
    return Border(left=s, right=s, top=s, bottom=s)


def _fmt_euro(n: float) -> str:
    return f"€{n:,.0f}".replace(",", ".")


def _period_label(period: str) -> str:
    try:
        y, m = period.split("-")
        months = ["","Gennaio","Febbraio","Marzo","Aprile","Maggio","Giugno",
                  "Luglio","Agosto","Settembre","Ottobre","Novembre","Dicembre"]
        return f"{months[int(m)]} {y}"
    except Exception:
        return period


# ── Lead Register (Excel) ──────────────────────────────────────────────────────

def generate_lead_register(confirmed: list, totals: dict,
                           period: str, invoice_sales: str) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Lead Register"

    hdr_fill = PatternFill("solid", fgColor="1F4E79")
    alt_fill = PatternFill("solid", fgColor="F2F2F2")
    tot_fill = PatternFill("solid", fgColor="FFF2CC")
    hdr_font = Font(bold=True, color="FFFFFF", size=10)
    bold = Font(bold=True, size=10)
    normal = Font(size=10)

    headers = ["#", "Nome Cliente", "Data Contratto", "N° Macchine", "Importo (€)", "Periodo", "Fattura"]
    ws.append(headers)
    for col in range(1, len(headers)+1):
        c = ws.cell(1, col)
        c.fill = hdr_fill; c.font = hdr_font
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = _thin_border()

    for i, cl in enumerate(confirmed, 1):
        ws.append([i, cl["nome"], cl["data"], cl["n_macchine"],
                   cl["importo"], period, invoice_sales])
        r = ws.max_row
        for col in range(1, 8):
            c = ws.cell(r, col)
            c.fill = alt_fill if i % 2 == 0 else PatternFill()
            c.font = normal
            c.border = _thin_border()
            c.alignment = Alignment(
                horizontal="left" if col == 2 else "center", vertical="center")

    # Totali
    ws.append([])
    for label, value in [
        ("Fee fissa", totals["fixed_fee"]),
        ("Totale variabile", totals["variabile"]),
        ("TOTALE FATTURA", totals["totale_sales"]),
    ]:
        ws.append(["", label, "", "", value, "", ""])
        r = ws.max_row
        for col in range(1, 8):
            c = ws.cell(r, col)
            c.fill = tot_fill
            c.border = _thin_border()
            c.font = bold
        ws.cell(r, 5).alignment = Alignment(horizontal="center")

    ws.column_dimensions["A"].width = 5
    ws.column_dimensions["B"].width = 45
    ws.column_dimensions["C"].width = 16
    ws.column_dimensions["D"].width = 14
    ws.column_dimensions["E"].width = 14
    ws.column_dimensions["F"].width = 12
    ws.column_dimensions["G"].width = 14

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ── Activity Summary (Word) ────────────────────────────────────────────────────

def generate_activity_summary(totals: dict, period: str,
                               sweety: dict, gtg: dict) -> bytes:
    doc = Document()

    # Intestazione
    h = doc.add_heading(f"Activity Summary — {_period_label(period)}", 0)
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph()
    doc.add_paragraph(f"Periodo di riferimento: {_period_label(period)}")
    doc.add_paragraph(f"Data emissione: {datetime.today().strftime('%d/%m/%Y')}")
    doc.add_paragraph()

    doc.add_heading("Sales Support", level=1)
    table = doc.add_table(rows=1, cols=2)
    table.style = "Table Grid"
    hdr = table.rows[0].cells
    hdr[0].text = "Descrizione"
    hdr[1].text = "Importo"
    for c in hdr:
        for p in c.paragraphs:
            for r in p.runs:
                r.bold = True

    rows_data = [
        (f"Fee fissa mensile Sales Support", _fmt_euro(totals["fixed_fee"])),
        (f"Nuove attivazioni ({totals['macchine']} × {_fmt_euro(500)})", _fmt_euro(totals["variabile"])),
        ("TOTALE FATTURA SALES SUPPORT", _fmt_euro(totals["totale_sales"])),
    ]
    for desc, imp in rows_data:
        row = table.add_row().cells
        row[0].text = desc
        row[1].text = imp

    doc.add_paragraph()
    doc.add_heading("Administrative Services", level=1)
    table2 = doc.add_table(rows=1, cols=2)
    table2.style = "Table Grid"
    hdr2 = table2.rows[0].cells
    hdr2[0].text = "Descrizione"
    hdr2[1].text = "Importo"
    for c in hdr2:
        for p in c.paragraphs:
            for r in p.runs:
                r.bold = True
    row2 = table2.add_row().cells
    row2[0].text = "Servizi amministrativi mensili"
    row2[1].text = _fmt_euro(totals["totale_admin"])

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


# ── Bozza email Sales (txt) ────────────────────────────────────────────────────

def generate_email_draft_sales(totals: dict, period: str,
                                invoice_sales: str,
                                sweety: dict, gtg: dict) -> bytes:
    period_label = _period_label(period)
    body = f"""Oggetto: Fattura Sales Support {period_label} — {invoice_sales}

Gentili {gtg.get('name', 'GTG')},

in allegato trovate la fattura n. {invoice_sales} relativa ai servizi di Sales Support per il mese di {period_label}.

Riepilogo:
  • Fee fissa mensile:          {_fmt_euro(totals['fixed_fee'])}
  • Nuove attivazioni ({totals['macchine']} macchine): {_fmt_euro(totals['variabile'])}
  • TOTALE:                     {_fmt_euro(totals['totale_sales'])}

Dati bancari per il pagamento:
  Beneficiario: {sweety.get('name', 'Sweety Pact S.r.l.')}
  IBAN: {sweety.get('bank_iban', '')}
  BIC: {sweety.get('bank_bic', '')}
  Banca: {sweety.get('bank_name', '')}

Cordiali saluti,
{sweety.get('name', 'Sweety Pact S.r.l.')}
{sweety.get('email', '')}
"""
    return body.encode("utf-8")


# ── Bozza email Admin (txt) ────────────────────────────────────────────────────

def generate_email_draft_admin(totals: dict, period: str,
                                invoice_admin: str,
                                sweety: dict, admin_contact: dict) -> bytes:
    period_label = _period_label(period)
    body = f"""Oggetto: Fattura Administrative Services {period_label} — {invoice_admin}

Gentile {admin_contact.get('name', '')},

in allegato trovate la fattura n. {invoice_admin} per i Servizi Amministrativi del mese di {period_label}.

Importo: {_fmt_euro(totals['totale_admin'])}

Cordiali saluti,
{sweety.get('name', 'Sweety Pact S.r.l.')}
"""
    return body.encode("utf-8")


# ── Genera tutti i documenti ───────────────────────────────────────────────────

def generate_all(confirmed: list, totals: dict, period: str,
                 invoice_sales: str, invoice_admin: str,
                 sweety: dict, gtg: dict, admin_contact: dict) -> dict:
    """
    Genera tutti i documenti e li restituisce come dict {filename: bytes}.
    """
    return {
        f"lead_register_{period}.xlsx": generate_lead_register(
            confirmed, totals, period, invoice_sales),
        f"activity_summary_{period}.docx": generate_activity_summary(
            totals, period, sweety, gtg),
        f"email_sales_{period}.txt": generate_email_draft_sales(
            totals, period, invoice_sales, sweety, gtg),
        f"email_admin_{period}.txt": generate_email_draft_admin(
            totals, period, invoice_admin, sweety, admin_contact),
    }
