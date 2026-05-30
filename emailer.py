"""
Invio email con allegati via Resend API (HTTPS) — SMTP bloccato su Render free tier.
"""
import os
import base64
import requests

RESEND_API_KEY = os.environ.get("RESEND_API_KEY", "")
SMTP_EMAIL = os.environ.get("SMTP_EMAIL", "amministrazione.swt@gmail.com")
CONFIRM_EMAIL = os.environ.get("CONFIRM_EMAIL", "directionsweetypactsrl@gmail.com")

RESEND_API_URL = "https://api.resend.com/emails"
FROM_ADDRESS = "Sweety GTG <onboarding@resend.dev>"


def send_confirmation(period_label: str, totals: dict,
                      confirmed: list, docs: dict,
                      master_bytes: bytes) -> None:
    """
    Invia email di conferma via Resend con tutti i documenti allegati.
    """
    if not RESEND_API_KEY:
        raise ValueError("RESEND_API_KEY non configurata — vai su Render > Environment")

    si_list = "\n".join(
        f"  - {c['nome']} ({c['n_macchine']} macch. → {c['importo']} EUR)"
        for c in confirmed
    )
    body = (
        f"Revisione fatturazione {period_label} confermata.\n\n"
        f"CLIENTI CONFERMATI ({len(confirmed)}):\n{si_list}\n\n"
        f"RIEPILOGO FATTURA SALES SUPPORT:\n"
        f"  Fee fissa mensile:          {totals['fixed_fee']} EUR\n"
        f"  Attivazioni operative:      {totals['variabile']} EUR  ({totals['macchine']} macchine)\n"
        f"  TOTALE SALES SUPPORT:       {totals['totale_sales']} EUR\n\n"
        f"RIEPILOGO FATTURA ADMIN SERVICES:\n"
        f"  TOTALE ADMIN SERVICES:      {totals['totale_admin']} EUR\n\n"
        f"In allegato trovate:\n"
        f"  - Lead Register {period_label}\n"
        f"  - Activity Summary {period_label}\n"
        f"  - Bozze email (Sales + Admin)\n"
        f"  - Master clienti aggiornato\n\n"
        f"IMPORTANTE: il file 'master_aggiornato.csv' va ricaricato come storico il mese prossimo.\n"
    )

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
        "subject": f"Conferma fatturazione {period_label}",
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
