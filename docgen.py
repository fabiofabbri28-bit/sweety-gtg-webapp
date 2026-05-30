"""
Generazione documenti in memoria:
- Lead Register (Excel)
- Activity Summary (Word)
- Bozza email Sales (txt)
- Bozza email Admin (txt)
"""
import io
import re
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


MONTHS_EN = ["", "January", "February", "March", "April", "May", "June",
             "July", "August", "September", "October", "November", "December"]
MONTHS_ABBR = ["", "JAN", "FEB", "MAR", "APR", "MAY", "JUN",
               "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"]
MONTHS_IT = ["", "Gennaio", "Febbraio", "Marzo", "Aprile", "Maggio", "Giugno",
             "Luglio", "Agosto", "Settembre", "Ottobre", "Novembre", "Dicembre"]

BLUE_DARK  = "1F4E79"
BLUE_MID   = "2E75B6"
BLUE_LIGHT = "DAE8F5"

_BLUE  = RGBColor(0x1F, 0x4E, 0x79)
_GREY  = RGBColor(0x66, 0x66, 0x66)
_WHITE = RGBColor(0xFF, 0xFF, 0xFF)


# ── Helpers generici ───────────────────────────────────────────────────────────

def _thin_border():
    s = Side(style="thin", color="CCCCCC")
    return Border(left=s, right=s, top=s, bottom=s)


def _fmt_euro(n: float) -> str:
    return f"€{n:,.0f}".replace(",", ".")


def _period_label(period: str) -> str:
    """Etichetta italiana: 'Maggio 2026'"""
    try:
        y, m = period.split("-")
        return f"{MONTHS_IT[int(m)]} {y}"
    except Exception:
        return period


def _period_label_en(period: str) -> str:
    """Etichetta inglese: 'May 2026'"""
    try:
        y, m = period.split("-")
        return f"{MONTHS_EN[int(m)]} {y}"
    except Exception:
        return period


def _period_abbr(period: str) -> str:
    """'2026-05' → 'MAY-2026'"""
    try:
        y, m = period.split("-")
        return f"{MONTHS_ABBR[int(m)]}-{y}"
    except Exception:
        return period


def _data_ym_to_en(data_ym: str) -> str:
    """'2026-04' → 'April 2026'"""
    try:
        y, m = data_ym.split("-")
        return f"{MONTHS_EN[int(m)]} {y}"
    except Exception:
        return data_ym


def _city_from_address(address: str) -> str:
    """Estrae comune dall'indirizzo italiano."""
    if not address:
        return ""
    m = re.search(r'\b\d{5}\s+([A-Za-zÀ-ÿ\s\-\']+)', address)
    if m:
        city = m.group(1).strip()
        return re.sub(r'\s*\([A-Z]{2}\)\s*$', '', city).strip()
    parts = [p.strip() for p in address.split(",") if p.strip()]
    return parts[-1] if parts else address


# ── Helpers Word ────────────────────────────────────────────────────────────────

def _set_cell_bg(cell, hex_color: str):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    for existing in tcPr.findall(qn("w:shd")):
        tcPr.remove(existing)
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tcPr.append(shd)


def _word_cell_text(cell, text: str, bold=False, size=10,
                    color: RGBColor = None,
                    align=WD_ALIGN_PARAGRAPH.LEFT):
    p = cell.paragraphs[0]
    p.clear()
    p.alignment = align
    run = p.add_run(text)
    run.bold = bold
    run.font.size = Pt(size)
    if color:
        run.font.color.rgb = color


def _fill_info_cell(cell, lines: list):
    """Riempie una cella Word con righe di testo multiple."""
    for i, line in enumerate(lines):
        p = cell.paragraphs[0] if i == 0 else cell.add_paragraph()
        if i == 0:
            p.clear()
        run = p.add_run(line)
        run.font.size = Pt(9)


def _add_section_heading(doc, text: str):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(4)
    run = p.add_run(text)
    run.bold = True
    run.font.color.rgb = _BLUE
    run.font.size = Pt(11)
    return p


# ── Lead Register (Excel) ──────────────────────────────────────────────────────

def generate_lead_register(confirmed: list, totals: dict,
                           period: str, invoice_sales: str) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Lead Register"

    hdr_fill  = PatternFill("solid", fgColor=BLUE_DARK)
    alt_fill  = PatternFill("solid", fgColor=BLUE_LIGHT)
    hdr_font  = Font(bold=True, color="FFFFFF", size=10)
    body_font = Font(size=10)

    headers = ["Lead ID", "Origin period", "Customer", "City", "N. machines", "Billable"]
    ws.append(headers)
    for col in range(1, len(headers) + 1):
        c = ws.cell(1, col)
        c.fill = hdr_fill
        c.font = hdr_font
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = _thin_border()
    ws.row_dimensions[1].height = 22

    prefix = _period_abbr(period)
    seq = 0
    for cl in confirmed:
        for row_data in cl.get("rows", []):
            seq += 1
            lead_id = f"{prefix}-{seq:03d}"
            origin  = _data_ym_to_en(row_data.get("data_ym", period))
            city    = _city_from_address(row_data.get("indirizzo", ""))
            ws.append([lead_id, origin, cl["nome"], city, 1, "Yes"])
            r = ws.max_row
            fill = alt_fill if seq % 2 == 0 else PatternFill()
            for col in range(1, 7):
                c = ws.cell(r, col)
                c.fill = fill
                c.font = body_font
                c.border = _thin_border()
                c.alignment = Alignment(
                    horizontal="left" if col == 3 else "center",
                    vertical="center",
                )
            ws.row_dimensions[r].height = 18

    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 42
    ws.column_dimensions["D"].width = 22
    ws.column_dimensions["E"].width = 14
    ws.column_dimensions["F"].width = 12

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ── Activity Summary (Word) ────────────────────────────────────────────────────

def generate_activity_summary(totals: dict, period: str,
                               sweety: dict, gtg: dict) -> bytes:
    doc = Document()

    for sec in doc.sections:
        sec.top_margin    = Cm(2.0)
        sec.bottom_margin = Cm(2.0)
        sec.left_margin   = Cm(2.5)
        sec.right_margin  = Cm(2.5)

    # ── Titolo ──────────────────────────────────────────────────────────────────
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("ACTIVITY SUMMARY")
    r.bold = True
    r.font.size = Pt(18)
    r.font.color.rgb = _BLUE

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(_period_label_en(period))
    r.font.size = Pt(12)
    r.font.color.rgb = _BLUE

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(f"Issued: {datetime.today().strftime('%d %B %Y')}")
    r.font.size = Pt(9)
    r.font.color.rgb = _GREY

    doc.add_paragraph()

    # ── Provider / Client ────────────────────────────────────────────────────────
    pc = doc.add_table(rows=2, cols=2)
    pc.style = "Table Grid"

    _set_cell_bg(pc.cell(0, 0), BLUE_DARK)
    _set_cell_bg(pc.cell(0, 1), BLUE_DARK)
    _word_cell_text(pc.cell(0, 0), "PROVIDER", bold=True, color=_WHITE)
    _word_cell_text(pc.cell(0, 1), "CLIENT",   bold=True, color=_WHITE)

    _fill_info_cell(pc.cell(1, 0), [
        sweety.get("name", ""),
        sweety.get("address_line1", ""),
        sweety.get("address_line2", ""),
        f"Fiscal Code: {sweety.get('cod_fiscal', '')}",
        sweety.get("email", "info@sweetypact.com"),
        sweety.get("phone", "+37360045404"),
    ])
    _fill_info_cell(pc.cell(1, 1), [
        gtg.get("name", ""),
        gtg.get("address", "România, comuna Arad, str. Poetului, nr.1/c, hala 21"),
        f"Fiscal Code: {gtg.get('cod_fiscal', '')}",
    ])

    for row in pc.rows:
        for cell in row.cells:
            cell.width = Cm(8.25)

    doc.add_paragraph()

    # ── Sezione 1: Sales Support ─────────────────────────────────────────────────
    _add_section_heading(doc, "1.  SALES SUPPORT")

    ss = doc.add_table(rows=1, cols=2)
    ss.style = "Table Grid"
    _set_cell_bg(ss.cell(0, 0), BLUE_MID)
    _set_cell_bg(ss.cell(0, 1), BLUE_MID)
    _word_cell_text(ss.cell(0, 0), "Description", bold=True, color=_WHITE)
    _word_cell_text(ss.cell(0, 1), "Amount", bold=True, color=_WHITE,
                    align=WD_ALIGN_PARAGRAPH.RIGHT)

    ss_rows = [
        ("Fixed fee — Sales Support",
         _fmt_euro(totals["fixed_fee"])),
        (f"Operational activations — {totals['macchine']} machines × {_fmt_euro(500)}",
         _fmt_euro(totals["variabile"])),
        ("TOTAL SALES SUPPORT",
         _fmt_euro(totals["totale_sales"])),
    ]
    for i, (desc, amt) in enumerate(ss_rows):
        row = ss.add_row()
        is_tot = i == len(ss_rows) - 1
        if is_tot:
            _set_cell_bg(row.cells[0], BLUE_LIGHT)
            _set_cell_bg(row.cells[1], BLUE_LIGHT)
        _word_cell_text(row.cells[0], desc, bold=is_tot)
        _word_cell_text(row.cells[1], amt,  bold=is_tot,
                        align=WD_ALIGN_PARAGRAPH.RIGHT)

    for row in ss.rows:
        row.cells[0].width = Cm(12.5)
        row.cells[1].width = Cm(4.0)

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


# ── Bozza email Sales (txt) ────────────────────────────────────────────────────

def generate_email_draft_sales(totals: dict, period: str,
                                invoice_sales: str,
                                sweety: dict, gtg: dict) -> bytes:
    period_label = _period_label(period)
    body = (
        f"Oggetto: Sweety Pact – Sales Support {period_label} | Fattura n. {invoice_sales}\n\n"
        f"Gentili,\n\n"
        f"Vi trasmettiamo in allegato la documentazione relativa alla fatturazione mensile "
        f"per il mese di {period_label}.\n\n"
        f"RIEPILOGO FATTURA SALES SUPPORT n. {invoice_sales}:\n"
        f"  • Fee fissa mensile:                         {_fmt_euro(totals['fixed_fee'])}\n"
        f"  • Attivazioni operative ({totals['macchine']} macchine × {_fmt_euro(500)}): "
        f"{_fmt_euro(totals['variabile'])}\n"
        f"  • TOTALE:                                    {_fmt_euro(totals['totale_sales'])}\n\n"
        f"ALLEGATI:\n"
        f"  - Fattura Sales Support n. {invoice_sales}\n"
        f"  - Activity Summary {period_label}\n"
        f"  - Lead Register {period_label}\n"
        f"  - Master clienti aggiornato\n\n"
        f"Per qualsiasi chiarimento o informazione aggiuntiva siamo a vostra completa "
        f"disposizione.\n\n"
        f"Cordiali saluti,\n"
        f"{sweety.get('name', 'Sweety Pact S.r.l.')}\n"
        f"{sweety.get('email', 'info@sweetypact.com')}\n"
        f"{sweety.get('phone', '+37360045404')}\n"
    )
    return body.encode("utf-8")


# ── Bozza email Admin (txt) ────────────────────────────────────────────────────

def generate_email_draft_admin(totals: dict, period: str,
                                invoice_admin: str,
                                sweety: dict, admin_contact: dict) -> bytes:
    period_label = _period_label(period)
    name = admin_contact.get("name", "")
    body = (
        f"Oggetto: Sweety Pact – Administrative Services {period_label} | "
        f"Fattura n. {invoice_admin}\n\n"
        f"Gentile{' ' + name if name else ''},\n\n"
        f"Vi trasmettiamo in allegato la fattura n. {invoice_admin} per i Servizi "
        f"Amministrativi & Pre-contabilità relativi al mese di {period_label}.\n\n"
        f"  • Importo: {_fmt_euro(totals['totale_admin'])}\n\n"
        f"Per qualsiasi chiarimento siamo a vostra disposizione.\n\n"
        f"Cordiali saluti,\n"
        f"{sweety.get('name', 'Sweety Pact S.r.l.')}\n"
        f"{sweety.get('email', 'info@sweetypact.com')}\n"
        f"{sweety.get('phone', '+37360045404')}\n"
    )
    return body.encode("utf-8")


# ── Genera tutti i documenti ───────────────────────────────────────────────────

def generate_all(confirmed: list, totals: dict, period: str,
                 invoice_sales: str, invoice_admin: str,
                 sweety: dict, gtg: dict, admin_contact: dict) -> dict:
    return {
        f"lead_register_{period}.xlsx": generate_lead_register(
            confirmed, totals, period, invoice_sales),
        f"activity_summary_{period}.docx": generate_activity_summary(
            totals, period, sweety, gtg),
        f"email_sales_{period}.txt": generate_email_draft_sales(
            totals, period, invoice_sales, sweety, gtg),
    }
