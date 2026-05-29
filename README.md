# Sweety GTG Web App

## Descrizione
Web app mensile per la gestione del flusso di fatturazione tra Sweety Pact S.r.l. e Good to Great S.r.l. Permette a Irene di caricare il report GTG mensile, revisionare i clienti da fatturare e confermare con un click — l'app genera automaticamente tutti i documenti e invia l'email di conferma.

## Obiettivo
Eliminare il passaggio manuale di file e sostituirlo con un flusso autonomo: Irene carica i file, rivede i dati, conferma. Zero intervento tecnico necessario.

## Flusso di utilizzo mensile

1. **Upload** — Irene carica il report GTG mensile (CSV) e lo storico clienti già fatturati (CSV o Excel)
2. **Revisione** — L'app mostra la tabella dei nuovi clienti rilevati, pre-selezionati, con totali aggiornati in tempo reale
3. **Conferma** — Irene de-seleziona chi non fatturare e clicca "Conferma ✅"
4. **Output automatico** — L'app genera tutti i documenti e invia l'email a `directionsweetypactsrl@gmail.com`
5. **Download** — Irene scarica i documenti (incluso `master_aggiornato.csv` da usare il mese successivo come storico)

## Stack tecnologico
- Python 3.x + Flask (web server)
- Gunicorn (server produzione)
- pandas, openpyxl, python-docx, rapidfuzz (elaborazione dati)
- Bootstrap 5 (frontend)
- Render.com (hosting)
- Gmail SMTP (invio email)

## Architettura

```
Browser (Irene)
    │
    ▼
Flask App (app.py)
    ├── pipeline.py    ← carica CSV, fuzzy match, calcola candidati
    ├── docgen.py      ← genera Excel, Word, .txt in memoria
    └── emailer.py     ← invia email con allegati via Gmail SMTP
```

Ogni sessione è identificata da un UUID. I file caricati e generati vengono salvati in `sessions/<uuid>/` e cancellati automaticamente dopo 48 ore.

## Struttura del progetto

```
sweety-gtg-webapp/
├── app.py                  ← Flask routes e logica sessioni
├── pipeline.py             ← Parsing CSV GTG + fuzzy matching + calcoli
├── docgen.py               ← Generazione documenti (Excel, Word, .txt)
├── emailer.py              ← Invio email SMTP
├── requirements.txt        ← Dipendenze Python
├── render.yaml             ← Configurazione deploy Render
├── templates/
│   ├── base.html           ← Layout comune Bootstrap 5
│   ├── login.html          ← Pagina accesso con codice
│   ├── upload.html         ← Upload file GTG + storico
│   ├── review.html         ← Tabella revisione con checkbox e totali live
│   └── done.html           ← Pagina finale con link download
└── sessions/               ← Dati sessioni temporanei (non in git)
```

## Setup e configurazione

### Variabili d'ambiente (obbligatorie in produzione)

| Variabile | Descrizione | Default |
|-----------|-------------|---------|
| `ACCESS_CODE` | Codice accesso per Irene | `sweety2026` |
| `SMTP_EMAIL` | Email Gmail mittente | — |
| `SMTP_PASSWORD` | App password Gmail (non la password normale) | — |
| `CONFIRM_EMAIL` | Email destinatario conferma | `directionsweetypactsrl@gmail.com` |
| `SECRET_KEY` | Chiave sessioni Flask | generata da Render |
| `RATE_PER_CONTRACT` | Importo per contratto (€) | `500` |
| `FIXED_FEE` | Fee fissa sales support (€) | `1000` |
| `ADMIN_FEE` | Fee servizi amministrativi (€) | `300` |

### App password Gmail
Per abilitare l'invio email:
1. Vai su [myaccount.google.com/apppasswords](https://myaccount.google.com/apppasswords)
2. Crea una nuova app password per "Mail"
3. Inserisci la password generata nella variabile `SMTP_PASSWORD` su Render

### Accesso
L'app è accessibile tramite `?code=<ACCESS_CODE>` nell'URL (es. `https://sweety-gtg.onrender.com/?code=sweety2026`).

## Logica di rilevamento nuovi clienti

- Il report GTG viene analizzato per stato: `Firmato` e `In prova` = nuovi da fatturare; `Attivo` = già esistente, ignorato
- I nuovi clienti vengono confrontati con lo storico tramite **fuzzy matching** (soglia 90%) per evitare doppia fatturazione in caso di variazioni ortografiche nel nome
- Più righe con lo stesso cliente = più macchine → l'importo si moltiplica per il numero di macchine

## Note importanti
- Lo storico da caricare ogni mese è il file `master_aggiornato.csv` allegato all'email di conferma del mese precedente
- I documenti generati non contengono nomi di clienti (solo conteggi aggregati) — per conformità privacy
- Le sessioni scadono dopo 48 ore — scaricare i file prima

## Stato del progetto
Attivo — Maggio 2026
