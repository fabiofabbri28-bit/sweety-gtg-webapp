# Sweety GTG — Web App

## Obiettivo
App web per il workflow mensile di fatturazione Sweety → GTG.
Irene carica il report CSV di GTG + lo storico, revisiona i clienti, clicca Conferma.
Email inviata automaticamente con i documenti allegati.

## Stack
- Python 3.x / Flask
- pandas, openpyxl, python-docx, rapidfuzz
- Bootstrap 5 (frontend)
- Deploy: Render (free tier)

## Flusso
1. Irene apre l'app (link Render)
2. Carica report GTG mensile (CSV) + storico clienti già fatturati (CSV, opzionale)
3. L'app mostra tabella con spunte — Irene de-flagga i clienti da escludere
4. Clicca "Conferma ✅" → email a directionsweetypactsrl@gmail.com con allegati
5. Documenti scaricabili anche dall'app

## Variabili d'ambiente (Render)
- `ACCESS_CODE` — codice accesso app (es. sweety2026)
- `SMTP_EMAIL` — email mittente (es. fabiofabbri28@gmail.com)
- `SMTP_PASSWORD` — app password Gmail
- `CONFIRM_EMAIL` — destinatario conferma (directionsweetypactsrl@gmail.com)
- `RATE_PER_CONTRACT` — 500
- `FIXED_FEE` — 1000
- `ADMIN_FEE` — 300

## Storico clienti
Il file storico (CSV) contiene i clienti già fatturati nei mesi precedenti.
Dopo ogni conferma, l'app allega all'email il master aggiornato da usare il mese successivo.
Formato accettato: qualsiasi CSV con colonna "Legal Name" o "Nome_Cliente".
