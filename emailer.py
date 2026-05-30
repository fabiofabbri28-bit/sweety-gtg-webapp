"""
Invio email con allegati via Resend API (HTTPS) — SMTP bloccato su Render free tier.
"""
import os
import base64
import requests

RESEND_API_KEY = os.environ.get("RESEND_API_KEY", "")
SMTP_EMAIL     = os.environ.get("SMTP_EMAIL", "amministrazione.swt@gmail.com")
CONFIRM_EMAIL  = os.environ.get("CONFIRM_EMAIL", "directionsweetypactsrl@gmail.com")

RESEND_API_URL = "https://api.resend.com/emails"
FROM_ADDRESS   = os.environ.get("FROM_ADDRESS", "Sweety Pact <onboarding@resend.dev>")


def send_confirmation(period_label: str, totals: dict,
                      confirmed: list, docs: dict,
                      master_bytes: bytes,
                      invoice_sales: str = "",
                      invoice_admin: str = "") -> None:
    """Invia email di fatturazione mensile via Resend con tutti i documenti allegati."""
    if not RESEND_API_KEY:
        raise ValueError("RESEND_API_KEY non configurata — vai su Render > Environment")

    si_list = "\n".join(
        f"  - {c['nome']} ({c['n_macchine']} macch. → {c['importo']} EUR)"
        for c in confirmed
    )

    sales_line = f" – Fattura n. {invoice_sales}" if invoice_sales else ""
    admin_line = f" – Fattura n. {invoice_admin}" if invoice_admin else ""

    body = (
        f"Gentili,\n\n"
        f"Vi trasmettiamo in allegato la documentazione relativa alla fatturazione mensile "
        f"per il mese di {period_label}.\n\n"
        f"RIEPILOGO SALES SUPPORT{sales_line}:\n"
        f"  • Fee fissa mensile:                    {totals['fixed_fee']} EUR\n"
        f"  • Attivazioni operative "
        f"({totals['macchine']} macchine × 500 EUR):  {totals['variabile']} EUR\n"
        f"  • TOTALE SALES SUPPORT:                 {totals['totale_sales']} EUR\n\n"
        f"RIEPILOGO ADMIN SERVICES{admin_line}:\n"
        f"  • Servizi Amministrativi & Pre-contabilità:  {totals['totale_admin']} EUR\n\n"
        f"CLIENTI NUOVI CONFERMATI ({len(confirmed)}):\n{si_list}\n\n"
        f"In allegato trovate:\n"
        f"  - Lead Register {period_label}\n"
        f"  - Activity Summary {period_label}\n"
        f"  - Bozza email Sales Support\n"
        f"  - Bozza email Admin Services\n"
        f"  - Master clienti aggiornato "
        f"(da ricaricare come storico il mese prossimo)\n\n"
        f"Per qualsiasi chiarimento siamo a vostra completa disposizione.\n\n"
        f"Cordiali saluti,\n"
        f"Sweety Pact S.r.l.\n"
        f"info@sweetypact.com\n"
        f"+37360045404\n"
    )

    subject = f"Sweety Pact – Sales Support {period_label}{sales_line}"

    attachments = []
    for filename, data in docs.items():
        attachments.append({
            "filename": filename,
            "content": base64.b64encode(data).decode("ascii"),
        })
    attachments.append({
        "filename": "master_aggiornato.csv",
        "content": base64.b64encode(master_bytes).decode("ascii"),
    })

    payload = {
        "from": FROM_ADDRESS,
        "to": [CONFIRM_EMAIL],
        "reply_to": SMTP_EMAIL,
        "subject": subject,
        "text": body,
        "attachments": attachments,
    }

    resp = requests.post(
        RESEND_API_URL,
        headers={"Authorization": f"Bearer {RESEND_API_KEY}"},
        json=payload,
        timeout=15,
    )

    if resp.status_code not in (200, 201):
        raise RuntimeError(f"Resend error {resp.status_code}: {resp.text[:200]}")
