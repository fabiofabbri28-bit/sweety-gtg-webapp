"""
Invio email con allegati tramite SMTP Gmail.
"""
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders

SMTP_HOST = "smtp.gmail.com"
SMTP_PORT = 587
SMTP_EMAIL = os.environ.get("SMTP_EMAIL", "")
SMTP_PASSWORD = os.environ.get("SMTP_PASSWORD", "")
CONFIRM_EMAIL = os.environ.get("CONFIRM_EMAIL", "directionsweetypactsrl@gmail.com")


def send_confirmation(period_label: str, totals: dict,
                      confirmed: list, docs: dict,
                      master_bytes: bytes) -> None:
    """
    Invia email di conferma con:
    - corpo riepilogativo
    - tutti i documenti allegati
    - master aggiornato allegato (per il mese prossimo)
    """
    msg = MIMEMultipart()
    msg["From"] = SMTP_EMAIL
    msg["To"] = CONFIRM_EMAIL
    msg["Subject"] = f"✅ Conferma fatturazione {period_label}"

    # Corpo email
    si_list = "\n".join(
        f"  • {c['nome']} ({c['n_macchine']} macch. → €{c['importo']:,})".replace(",", ".")
        for c in confirmed
    )
    body = f"""Revisione fatturazione confermata per {period_label}.

CLIENTI CONFERMATI ({len(confirmed)}):
{si_list}

RIEPILOGO:
  Fee fissa:          €{totals['fixed_fee']:,}
  Totale variabile:   €{totals['variabile']:,}  ({totals['macchine']} macchine)
  TOTALE SALES:       €{totals['totale_sales']:,}
  TOTALE ADMIN:       €{totals['totale_admin']:,}

I documenti sono allegati a questa email.
Il file "master_aggiornato.csv" va ricaricato il mese prossimo come storico.
""".replace(",", ".")

    msg.attach(MIMEText(body, "plain", "utf-8"))

    # Allega documenti generati
    for filename, data in docs.items():
        part = MIMEBase("application", "octet-stream")
        part.set_payload(data)
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f'attachment; filename="{filename}"')
        msg.attach(part)

    # Allega master aggiornato
    part_master = MIMEBase("application", "octet-stream")
    part_master.set_payload(master_bytes)
    encoders.encode_base64(part_master)
    part_master.add_header("Content-Disposition", 'attachment; filename="master_aggiornato.csv"')
    msg.attach(part_master)

    with smtplib.SMTP_SSL(SMTP_HOST, 465, timeout=10) as server:
        server.ehlo()
        server.login(SMTP_EMAIL, SMTP_PASSWORD)
        server.send_message(msg)
