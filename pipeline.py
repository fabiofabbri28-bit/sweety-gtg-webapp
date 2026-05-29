"""
Core processing logic — adattato dagli script della pipeline locale.
Accetta GTG CSV + storico opzionale, restituisce i dati per la revisione.
"""
import io
import os
import pandas as pd
from datetime import datetime
from rapidfuzz import fuzz

RATE = int(os.environ.get("RATE_PER_CONTRACT", 500))
FIXED_FEE = int(os.environ.get("FIXED_FEE", 1000))
ADMIN_FEE = int(os.environ.get("ADMIN_FEE", 300))

GTG_COMMERCIAL_COLS = ["Cliente", "Data Contratto", "Stato"]


# ── Caricamento GTG CSV ────────────────────────────────────────────────────────

def load_gtg_csv(file_bytes: bytes, filename: str) -> pd.DataFrame:
    """Carica il report GTG (formato commerciale con ';')."""
    text = file_bytes.decode("utf-8-sig")
    df = pd.read_csv(io.StringIO(text), sep=";")
    df.columns = df.columns.str.strip()

    if not all(c in df.columns for c in GTG_COMMERCIAL_COLS):
        # Prova formato interno
        df2 = pd.read_csv(io.StringIO(text))
        df2.columns = df2.columns.str.strip()
        return _normalize_internal(df2)

    return _normalize_commercial(df)


def _normalize_commercial(df: pd.DataFrame) -> pd.DataFrame:
    """Converte formato commerciale GTG in formato interno."""
    df = df.copy()
    df["Stato"] = df["Stato"].str.strip()
    out = pd.DataFrame()
    out["nome"] = df["Cliente"].str.strip()
    out["data"] = pd.to_datetime(df["Data Contratto"], dayfirst=True, errors="coerce")
    out["stato_orig"] = df["Stato"]
    # Firmato / In prova → nuovi; Attivo → già esistente
    out["is_existing"] = df["Stato"].apply(lambda s: s.strip() == "Attivo")
    return out


def _normalize_internal(df: pd.DataFrame) -> pd.DataFrame:
    out = pd.DataFrame()
    out["nome"] = df.get("Nome_Cliente", df.get("Cliente", pd.Series(dtype=str))).str.strip()
    out["data"] = pd.to_datetime(df.get("Data_Attivazione", df.get("Data Contratto", pd.Series())), errors="coerce")
    out["stato_orig"] = df.get("Stato", "")
    out["is_existing"] = df.get("Stato", "").apply(
        lambda s: str(s).strip() == "Attivo" and "già" in str(s).lower()
    ) if "Note" not in df.columns else df.get("Note", "").str.contains("già", na=False)
    return out


# ── Caricamento storico ────────────────────────────────────────────────────────

def load_storico(file_bytes: bytes) -> set:
    """
    Carica lo storico clienti già fatturati.
    Accetta CSV (separatore , o ;) oppure Excel (.xlsx).
    Restituisce un set di nomi (lowercase stripped).
    """
    # Prova Excel
    try:
        df = pd.read_excel(io.BytesIO(file_bytes))
        df.columns = df.columns.str.strip()
        for col in ["Legal Name", "Nome_Cliente", "nome", "Cliente", "Customer"]:
            if col in df.columns:
                return set(df[col].dropna().str.strip().str.lower())
    except Exception:
        pass

    # Prova CSV
    for enc in ["utf-8-sig", "utf-8", "latin-1"]:
        try:
            text = file_bytes.decode(enc)
            for sep in [";", ","]:
                try:
                    df = pd.read_csv(io.StringIO(text), sep=sep)
                    df.columns = df.columns.str.strip()
                    for col in ["Legal Name", "Nome_Cliente", "nome", "Cliente", "Customer"]:
                        if col in df.columns:
                            return set(df[col].dropna().str.strip().str.lower())
                except Exception:
                    continue
        except Exception:
            continue
    return set()


# ── Fuzzy match ────────────────────────────────────────────────────────────────

def _is_already_billed(nome: str, billed_set: set, threshold: int = 90) -> bool:
    nome_low = nome.lower()
    if nome_low in billed_set:
        return True
    for existing in billed_set:
        if fuzz.token_sort_ratio(nome_low, existing) >= threshold:
            return True
    return False


# ── Elaborazione principale ────────────────────────────────────────────────────

def process(gtg_bytes: bytes, gtg_filename: str,
            storico_bytes=None) -> dict:
    """
    Carica e processa i file.
    Restituisce dizionario con tutto il necessario per la pagina di revisione.
    """
    df = load_gtg_csv(gtg_bytes, gtg_filename)

    # Storico clienti già fatturati
    billed_set = set()
    if storico_bytes:
        billed_set = load_storico(storico_bytes)

    # Determina periodo dal CSV
    dates = df["data"].dropna()
    if not dates.empty:
        period = dates.max().strftime("%Y-%m")
    else:
        period = datetime.today().strftime("%Y-%m")

    # Separa nuovi da già esistenti
    candidates = []
    skipped = []

    # Raggruppa per nome (più righe = più macchine)
    groups = df[~df["is_existing"]].groupby("nome")
    for nome, grp in groups:
        if _is_already_billed(nome, billed_set):
            skipped.append(nome)
            continue
        data_min = grp["data"].min()
        n_macchine = len(grp)
        candidates.append({
            "nome": nome,
            "data": data_min.strftime("%d/%m/%Y") if pd.notna(data_min) else "",
            "n_macchine": n_macchine,
            "importo": n_macchine * RATE,
            "selected": True,
        })

    # Ordina per data
    candidates.sort(key=lambda x: x["data"])

    return {
        "period": period,
        "candidates": candidates,
        "skipped": skipped,
        "rate": RATE,
        "fixed_fee": FIXED_FEE,
        "admin_fee": ADMIN_FEE,
        "n_total_gtg": len(df),
        "n_skipped": len(skipped),
    }


# ── Calcolo totali ─────────────────────────────────────────────────────────────

def compute_totals(candidates: list, selected_names: set) -> dict:
    """Calcola i totali in base alle selezioni."""
    macchine = 0
    variabile = 0
    for c in candidates:
        if c["nome"] in selected_names:
            macchine += c["n_macchine"]
            variabile += c["importo"]
    return {
        "macchine": macchine,
        "variabile": variabile,
        "fixed_fee": FIXED_FEE,
        "totale_sales": variabile + FIXED_FEE,
        "totale_admin": ADMIN_FEE,
    }


# ── Genera master aggiornato (CSV da allegare) ─────────────────────────────────

def generate_updated_master(old_storico_bytes,
                            confirmed_candidates: list,
                            period: str) -> bytes:
    """
    Genera il master aggiornato (CSV) da allegare all'email dopo la conferma.
    Include lo storico precedente + i nuovi clienti confermati.
    """
    rows = []

    # Carica vecchio storico se presente
    if old_storico_bytes:
        try:
            text = old_storico_bytes.decode("utf-8-sig")
            for sep in [";", ","]:
                try:
                    old_df = pd.read_csv(io.StringIO(text), sep=sep)
                    old_df.columns = old_df.columns.str.strip()
                    for col in ["Legal Name", "Nome_Cliente", "nome"]:
                        if col in old_df.columns:
                            for _, row in old_df.iterrows():
                                rows.append({
                                    "Nome_Cliente": str(row[col]).strip(),
                                    "Periodo_Fatturato": row.get("Periodo_Fatturato", ""),
                                    "Data_Inserimento": row.get("Data_Inserimento", ""),
                                })
                            break
                    break
                except Exception:
                    continue
        except Exception:
            pass

    # Aggiungi i nuovi confermati
    today = datetime.today().strftime("%Y-%m-%d")
    for c in confirmed_candidates:
        rows.append({
            "Nome_Cliente": c["nome"],
            "Periodo_Fatturato": period,
            "Data_Inserimento": today,
        })

    df = pd.DataFrame(rows)
    buf = io.BytesIO()
    df.to_csv(buf, index=False, encoding="utf-8-sig")
    return buf.getvalue()
