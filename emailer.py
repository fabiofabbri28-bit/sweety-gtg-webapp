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

    sales_line = f" | Invoice n. {invoice_sales}" if invoice_sales else ""

    body = (
        f"Good morning,\n\n"
        f"as a follow-up to the sales support services provided, please find attached "
        f"the Activity Summary Report and the related invoice for the month of {period_label}, "
        f"prepared in line with the contractual terms in force.\n\n"
        f"The report is provided as supporting documentation for invoicing and reconciliation "
        f"purposes.\n\n"
        f"The reporting period includes activities accounted for in accordance with the "
        f"contractual criteria and the operational reconciliation procedures adopted between "
        f"the parties.\n\n"
        f"Should you need any further clarification, we remain at your disposal.\n\n"
        f"Kind regards,\n\n"
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
