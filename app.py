import streamlit as st
import pandas as pd
import sqlite3
import os
import re
import tempfile
import logging
from contextlib import contextmanager
from datetime import datetime
from io import BytesIO

# ---------------------------------------------------------
# CONFIGURAZIONE PAGINA E SICUREZZA
# ---------------------------------------------------------
st.set_page_config(page_title="Riconciliazione Verifiche Impianti", page_icon="📊", layout="wide")

# Password letta prioritariamente da variabile d'ambiente, altrimenti da st.secrets.
PASSWORD_ACCESSO = os.environ.get("VERIFICHE_PASSWORD")
if not PASSWORD_ACCESSO:
    try:
        PASSWORD_ACCESSO = st.secrets["password"]
    except Exception:
        PASSWORD_ACCESSO = None

if not PASSWORD_ACCESSO:
    st.error("Configurazione di sicurezza incompleta: variabile d'ambiente VERIFICHE_PASSWORD o st.secrets['password'] assente.")
    st.stop()

# Logging strutturato di base
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S"
)
logger = logging.getLogger("verifiche")

# ---------------------------------------------------------
# COSTANTI E MAPPATURE
# ---------------------------------------------------------
DB_NAME = "verifiche.db"
BACKUP_DIR = "backups"
MAX_BACKUP = 15
SOGLIA_NPP = 0.20
TIPO_A = 'Ascensori (DPR 162/99)'
TIPO_E = 'Messa a Terra (DPR 462/01)'

RINOMINA = {
    'numero_verifica': 'Numero verifica', 'protocollo': 'Protocollo', 'tipo_impianto': 'Reparto',
    'codice_impianto': 'Codice impianto', 'data_pianificata': 'Data pianificata', 'stato': 'Stato',
    'fattura': 'Fattura', 'data_t0': 'Data T0', 'data_t1': 'Data T1',
    'verbale': 'Verbale', 'esito_verifica': 'Esito verifica',
    'verbale_da_inviare': 'Verbale da inviare / Approvato'
}

# =========================================================
# ==== LOGICA E FUNZIONALITÀ DI ELABORAZIONE DATI ====
# =========================================================

def pulisci_testo(v):
    """Normalizza un valore di cella: None se vuoto/NaN, testo pulito altrimenti."""
    if v is None:
        return None
    try:
        if pd.isna(v):
            return None
    except (TypeError, ValueError):
        pass
    s = str(v).strip()
    if s == '' or s.lower() in ('nan', 'none', 'nat'):
        return None
    if re.fullmatch(r'-?\d+\.0+', s):
        s = s.split('.')[0]
    return s

def _parse_una_data(v):
    """Restituisce 'YYYY-MM-DD' oppure None."""
    t = pulisci_testo(v)
    if t is None:
        return None
    if re.match(r'^\d{4}-\d{2}-\d{2}', t):
        dt = pd.to_datetime(t[:10], errors='coerce', format='%Y-%m-%d')
    else:
        dt = pd.to_datetime(t, errors='coerce', dayfirst=True)
    return None if pd.isna(dt) else dt.strftime('%Y-%m-%d')

def tipo_da_protocollo(protocollo):
    """Identifica il reparto dal sezionale (/A per Ascensori, /E per Messa a Terra)."""
    p = (pulisci_testo(protocollo) or '').upper()
    if p.endswith('/A'):
        return TIPO_A
    if p.endswith('/E'):
        return TIPO_E
    return None

def leggi_excel(file):
    """Legge l'Excel come testo e pulisce i nomi delle colonne."""
    df = pd.read_excel(file, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]
    return df

@contextmanager
def db_connection(commit=False):
    """Gestisce l'apertura e la chiusura della connessione al database SQLite."""
    conn = sqlite3.connect(DB_NAME)
    try:
        yield conn
        if commit:
            conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()

def init_db():
    """Inizializza la struttura del database SQLite."""
    with db_connection(commit=True) as conn:
        cursor = conn.cursor()
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS verifiche (
                numero_verifica INTEGER PRIMARY KEY,
                protocollo TEXT,
                codice_impianto TEXT,
                codice_fiscale TEXT,
                data_pianificata TEXT,
                importo REAL,
                tipo_impianto TEXT,
                fattura TEXT,
                stato TEXT,
                data_t0 TEXT,
                data_t1 TEXT
            )
        ''')
        cursor.execute("PRAGMA table_info(verifiche)")
        colonne = [c[1] for c in cursor.fetchall()]
        
        nuove_colonne = {
            'giorni_trascorsi': 'INTEGER',
            'fascia_tempo': 'TEXT',
            'giorni_da_pianificata': 'INTEGER',
            'verbale': 'TEXT',
            'esito_verifica': 'TEXT',
            'verbale_da_inviare': 'TEXT'
        }
        for col_name, col_type in nuove_colonne.items():
            if col_name not in colonne:
                cursor.execute(f"ALTER TABLE verifiche ADD COLUMN {col_name} {col_type}")

        cursor.execute('''
            CREATE TABLE IF NOT EXISTS verifiche_extra (
                numero_verifica INTEGER PRIMARY KEY,
                protocollo TEXT,
                tipo_impianto TEXT,
                data_pianificata TEXT,
                fattura TEXT,
                data_t1 TEXT
            )
        ''')

def crea_backup(motivo="auto"):
    """Crea una copia di sicurezza (.db) prima di apportare modifiche."""
    if not os.path.exists(DB_NAME):
        return None
    os.makedirs(BACKUP_DIR, exist_ok=True)
    motivo = re.sub(r'[^A-Za-z0-9_-]', '', motivo) or "auto"
    percorso = os.path.join(BACKUP_DIR, f"verifiche_{datetime.now().strftime('%Y%m%d_%H%M%S_%f')}_{motivo}.db")
    src = sqlite3.connect(DB_NAME)
    dst = sqlite3.connect(percorso)
    try:
        src.backup(dst)
    finally:
        dst.close()
        src.close()
    file_backup = sorted(f for f in os.listdir(BACKUP_DIR) if f.startswith("verifiche_") and f.endswith(".db"))
    for vecchio in file_backup[:-MAX_BACKUP]:
        try:
            os.remove(os.path.join(BACKUP_DIR, vecchio))
        except OSError:
            pass
    logger.info("Backup creato: %s", percorso)
    return percorso

def ultimo_backup():
    if not os.path.isdir(BACKUP_DIR):
        return None
    file_backup = sorted(f for f in os.listdir(BACKUP_DIR) if f.startswith("verifiche_") and f.endswith(".db"))
    if not file_backup:
        return None
    ts = os.path.getmtime(os.path.join(BACKUP_DIR, file_backup[-1]))
    return datetime.fromtimestamp(ts)

def valida_db_caricato(raw):
    """Verifica l'integrità di un file database caricato."""
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".db")
    try:
        tmp.write(raw)
        tmp.close()
        try:
            c = sqlite3.connect(tmp.name)
            try:
                esito = c.execute("PRAGMA integrity_check").fetchone()[0]
                if esito != 'ok':
                    return False, "Il file è danneggiato (controllo di integrità fallito)."
                tabelle = {r[0] for r in c.execute("SELECT name FROM sqlite_master WHERE type='table'")}
                if 'verifiche' not in tabelle:
                    return False, "Il file non contiene la tabella 'verifiche'."
                n = c.execute("SELECT COUNT(*) FROM verifiche").fetchone()[0]
                return True, f"File valido: contiene {n} verifiche."
            finally:
                c.close()
        except sqlite3.DatabaseError:
            return False, "Il file non è un database SQLite valido."
    finally:
        try:
            os.unlink(tmp.name)
        except OSError:
            pass

def carica_stato_db(conn):
    df = pd.read_sql_query(
        "SELECT numero_verifica, protocollo, tipo_impianto, data_pianificata, data_t0, stato, verbale, esito_verifica, verbale_da_inviare FROM verifiche", conn
    )
    return df.set_index('numero_verifica')

def carica_numeri_extra(conn):
    return {r[0] for r in conn.execute("SELECT numero_verifica FROM verifiche_extra")}

# ---------------- PREPARAZIONE E SALVATAGGIO T0 ----------------
def prepara_t0(df_raw):
    d = pd.DataFrame({
        'num': pd.to_numeric(df_raw['Numero verifica'].apply(pulisci_testo), errors='coerce'),
        'protocollo': df_raw['Protocollo'].apply(lambda x: (pulisci_testo(x) or '').upper()),
        'codice_impianto': df_raw['Codice impianto'].apply(pulisci_testo),
        'codice_fiscale': df_raw['Codice fiscale'].apply(pulisci_testo),
        'data_pian': df_raw['Data pianificata'].apply(_parse_una_data),
        'verbale': df_raw['Verbale'].apply(pulisci_testo) if 'Verbale' in df_raw.columns else None,
        'esito_verifica': df_raw['Esito verifica'].apply(pulisci_testo) if 'Esito verifica' in df_raw.columns else None,
        'verbale_da_inviare': df_raw['Verbale da inviare'].apply(pulisci_testo) if 'Verbale da inviare' in df_raw.columns else None,
    })
    if 'Importo' in df_raw.columns:
        d['importo'] = pd.to_numeric(df_raw['Importo'].astype(str).str.replace(',', '.', regex=False), errors='coerce')
    else:
        d['importo'] = float('nan')

    d['tipo'] = d['protocollo'].apply(tipo_da_protocollo)
    m_num = d['num'].isna()
    m_data = ~m_num & d['data_pian'].isna()
    m_tipo = ~m_num & ~m_data & d['tipo'].isna()
    valido = ~(m_num | m_data | m_tipo)
    out = d[valido].copy()
    out['num'] = out['num'].astype(int)
    n_dup = int(out['num'].duplicated(keep='last').sum())
    out = out.drop_duplicates('num', keep='last')
    info = {
        'righe_file': len(df_raw),
        'senza_numero': int(m_num.sum()),
        'senza_data': int(m_data.sum()),
        'altro_protocollo': int(m_tipo.sum()),
        'duplicati': n_dup,
    }
    return out, info

def pianifica_t0(df_t0, db, extra_nums):
    db_d = db.to_dict('index')
    nuove = presenti = date_cambiate = 0
    for r in df_t0.itertuples():
        rec = db_d.get(r.num)
        if rec is None:
            nuove += 1
        else:
            presenti += 1
            if rec['data_pianificata'] != r.data_pian:
                date_cambiate += 1
    da_extra = len(set(df_t0['num']) & set(extra_nums))
    return {'nuove': nuove, 'presenti': presenti, 'date_cambiate': date_cambiate, 'da_extra': da_extra}

SQL_UPSERT_T0 = '''
    INSERT INTO verifiche (
        numero_verifica, protocollo, codice_impianto, codice_fiscale, data_pianificata,
        importo, tipo_impianto, fattura, stato, data_t0, verbale, esito_verifica, verbale_da_inviare
    )
    VALUES (?, ?, ?, ?, ?, ?, ?, NULL, 'In Attesa', ?, ?, ?, ?)
    ON CONFLICT(numero_verifica) DO UPDATE SET
        protocollo=excluded.protocollo,
        codice_impianto=excluded.codice_impianto,
        codice_fiscale=excluded.codice_fiscale,
        data_pianificata=excluded.data_pianificata,
        importo=COALESCE(excluded.importo, verifiche.importo),
        tipo_impianto=excluded.tipo_impianto,
        stato = CASE WHEN verifiche.fattura IS NULL THEN 'In Attesa' ELSE verifiche.stato END,
        verbale = COALESCE(excluded.verbale, verifiche.verbale),
        esito_verifica = COALESCE(excluded.esito_verifica, verifiche.esito_verifica),
        verbale_da_inviare = COALESCE(excluded.verbale_da_inviare, verifiche.verbale_da_inviare)
'''

def applica_t0(conn, df_t0, dt0_str):
    righe = [
        (int(r.num), r.protocollo, r.codice_impianto, r.codice_fiscale, r.data_pian,
         None if pd.isna(r.importo) else float(r.importo), r.tipo, dt0_str,
         r.verbale, r.esito_verifica, r.verbale_da_inviare)
        for r in df_t0.itertuples()
    ]
    cur = conn.cursor()
    cur.executemany(SQL_UPSERT_T0, righe)
    cur.executemany("DELETE FROM verifiche_extra WHERE numero_verifica = ?", [(int(n),) for n in df_t0['num']])
    logger.info("T0 applicato: %d verifiche", len(righe))

# ---------------- PREPARAZIONE E SALVATAGGIO T1 ----------------
def prepara_t1(df_raw):
    d = pd.DataFrame({
        'num': pd.to_numeric(df_raw['Numero verifica'].apply(pulisci_testo), errors='coerce'),
        'fattura': df_raw['Fattura'].apply(pulisci_testo) if 'Fattura' in df_raw.columns else None,
        'data_pian': df_raw['Data pianificata'].apply(_parse_una_data),
        'protocollo': (df_raw['Protocollo'].apply(lambda x: (pulisci_testo(x) or '').upper())
                       if 'Protocollo' in df_raw.columns else None),
        'verbale': df_raw['Verbale'].apply(pulisci_testo) if 'Verbale' in df_raw.columns else None,
        'esito_verifica': df_raw['Esito verifica'].apply(pulisci_testo) if 'Esito verifica' in df_raw.columns else None,
        'verbale_da_inviare': df_raw['Verbale da inviare'].apply(pulisci_testo) if 'Verbale da inviare' in df_raw.columns else None,
    })
    m_num = d['num'].isna()
    d = d[~m_num].copy()
    d['num'] = d['num'].astype(int)
    d['_ha_fattura'] = d['fattura'].notna()
    d = d.sort_values('_ha_fattura', kind='stable')
    n_dup = int(d['num'].duplicated(keep='last').sum())
    d = d.drop_duplicates('num', keep='last').drop(columns='_ha_fattura')
    info = {'righe_file': len(df_raw), 'senza_numero': int(m_num.sum()), 'duplicati': n_dup}
    return d, info

def _txt(v):
    return v if isinstance(v, str) and v != '' else None

def rileva_reparti_t1(d, db):
    conteggio = {}
    db_tipi = db['tipo_impianto'].to_dict()
    for r in d.itertuples():
        tipo = None
        prot = _txt(r.protocollo)
        if prot:
            tipo = tipo_da_protocollo(prot)
        if tipo is None:
            tipo = db_tipi.get(int(r.num))
        if tipo:
            conteggio[tipo] = conteggio.get(tipo, 0) + 1
    return conteggio

def _processa_riga_t1(r, db_d, dt1, dt1_str):
    num = int(r.num)
    fattura = _txt(r.fattura)
    verbale = _txt(r.verbale)
    esito = _txt(r.esito_verifica)
    v_inviare = _txt(r.verbale_da_inviare)
    rec = db_d.get(num)

    if rec is None:
        protocollo = _txt(r.protocollo) or ''
        tipo = tipo_da_protocollo(protocollo) if protocollo else None
        if protocollo and tipo is None:
            return 'scartata_altro', None
        return 'extra', (num, protocollo, tipo, _txt(r.data_pian), fattura, dt1_str)

    if rec['stato'] == 'Fatturata':
        return 'gia_fatturata', None

    if fattura:
        return 'fatturata', (fattura, dt1_str, 0, 'N/D', 0, verbale, esito, v_inviare, num)

    if rec['stato'] == 'Non Più Presente':
        return 'riattivata', (verbale, esito, v_inviare, num)
        
    return 'senza_fattura', (verbale, esito, v_inviare, num)

def _calcola_npp(db_d, presenti, date_t1, reparti_coperti):
    base_attesa = 0
    npp = []
    for num, rec in db_d.items():
        if (rec['stato'] == 'In Attesa'
                and rec['data_pianificata'] in date_t1
                and rec['tipo_impianto'] in reparti_coperti):
            base_attesa += 1
            if num not in presenti:
                npp.append(int(num))
    return base_attesa, npp

def pianifica_t1(d, dt1, db, reparti_coperti):
    dt1_str = dt1.strftime('%Y-%m-%d')
    db_d = db.to_dict('index')
    presenti = set(d['num'].tolist())
    date_t1 = {x for x in (_txt(v) for v in d['data_pian']) if x}

    fatturate, riattivate, extra, date_incoerenti, aggiornamenti_verbali = [], [], [], [], []
    gia_fatturate = senza_fattura = scartate_altro = 0

    for r in d.itertuples():
        esito, payload = _processa_riga_t1(r, db_d, dt1, dt1_str)
        if esito == 'fatturata':
            fatturate.append(payload)
        elif esito == 'riattivata':
            riattivate.append(payload[3])
            aggiornamenti_verbali.append(payload)
        elif esito == 'extra':
            extra.append(payload)
        elif esito == 'gia_fatturata':
            gia_fatturate += 1
        elif esito == 'senza_fattura':
            senza_fattura += 1
            if payload and any(payload[:3]):
                aggiornamenti_verbali.append(payload)
        elif esito == 'scartata_altro':
            scartate_altro += 1

    base_attesa, npp = _calcola_npp(db_d, presenti, date_t1, reparti_coperti)
    pct_npp = (len(npp) / base_attesa) if base_attesa else 0.0

    return {
        'dt1_str': dt1_str,
        'tot_righe': len(d),
        'date_t1': date_t1,
        'reparti_coperti': list(reparti_coperti),
        'fatturate': fatturate,
        'riattivate': riattivate,
        'aggiornamenti_verbali': aggiornamenti_verbali,
        'extra': extra,
        'npp': npp,
        'gia_fatturate': gia_fatturate,
        'senza_fattura': senza_fattura,
        'scartate_altro': scartate_altro,
        'date_incoerenti': date_incoerenti,
        'base_attesa': base_attesa,
        'pct_npp': pct_npp,
        'oltre_soglia': pct_npp > SOGLIA_NPP and len(npp) > 0,
    }

def riepilogo_piano(piano, info1):
    righe = [
        ("Nuove fatture rilevate (verranno registrate)", len(piano['fatturate'])),
        ("Già fatturate in precedenza (ignorate per evitare raddoppi)", piano['gia_fatturate']),
        ("Presenti nel T1 ma ancora senza fattura", piano['senza_fattura']),
        ("Non programmate nel T0 (report a parte)", len(piano['extra'])),
        ("Escluse (protocollo diverso da /A e /E)", piano['scartate_altro']),
        ("Righe senza numero verifica (scartate)", info1['senza_numero']),
        ("Righe duplicate nel file Excel (scartate)", info1['duplicati']),
    ]
    df = pd.DataFrame(righe, columns=["Esito Riconciliazione", "Conteggio Righe"])
    quadra = int(df["Conteggio Righe"].sum()) == info1['righe_file']
    return df, quadra

SQL_UPSERT_EXTRA = '''
    INSERT INTO verifiche_extra (numero_verifica, protocollo, tipo_impianto, data_pianificata, fattura, data_t1)
    VALUES (?, ?, ?, ?, ?, ?)
    ON CONFLICT(numero_verifica) DO UPDATE SET
        fattura = COALESCE(excluded.fattura, verifiche_extra.fattura),
        data_pianificata = COALESCE(excluded.data_pianificata, verifiche_extra.data_pianificata)
'''

def applica_t1(conn, piano):
    cur = conn.cursor()
    cur.executemany('''
        UPDATE verifiche
        SET fattura = ?, stato = 'Fatturata', data_t1 = ?, giorni_trascorsi = ?, fascia_tempo = ?, giorni_da_pianificata = ?,
            verbale = COALESCE(?, verbale), esito_verifica = COALESCE(?, esito_verifica), verbale_da_inviare = COALESCE(?, verbale_da_inviare)
        WHERE numero_verifica = ?
    ''', piano['fatturate'])
    
    cur.executemany('''
        UPDATE verifiche
        SET verbale = COALESCE(?, verbale), esito_verifica = COALESCE(?, esito_verifica), verbale_da_inviare = COALESCE(?, verbale_da_inviare)
        WHERE numero_verifica = ?
    ''', piano['aggiornamenti_verbali'])

    cur.executemany("UPDATE verifiche SET stato = 'In Attesa' WHERE numero_verifica = ?",
                    [(n,) for n in piano['riattivate']])
    cur.executemany("UPDATE verifiche SET stato = 'Non Più Presente' WHERE numero_verifica = ?",
                    [(n,) for n in piano['npp']])
    cur.executemany(SQL_UPSERT_EXTRA, piano['extra'])
    logger.info("T1 applicato con successo.")

# =========================================================
# ==== REPORTISTICA RIELABORATA E GIORNALIERA ====
# =========================================================

def genera_report_giornaliero(df_sub):
    """
    Calcola l'aggregazione giornaliera focalizzata su:
    Pianificate (T0), Fatturate, In Attesa (con sotto-fasi verbali) e Non Più Presenti.
    """
    df_a = df_sub.copy()
    
    # Classificazione degli stati principali
    df_a['Is_Fatturata'] = df_a['stato'] == 'Fatturata'
    df_a['Is_InAttesa'] = df_a['stato'] == 'In Attesa'
    df_a['Is_NPP'] = df_a['stato'] == 'Non Più Presente'
    
    # Classificazione delle sotto-fasi dei verbali (solo per le verifiche 'In Attesa')
    df_a['Ha_Verbale'] = df_a['verbale'].notna()
    df_a['Approvato'] = df_a['verbale_da_inviare'].astype(str).str.upper() == 'SI'
    
    df_a['Senza_Verbale'] = df_a['Is_InAttesa'] & (~df_a['Ha_Verbale'])
    df_a['In_Attesa_Ingegnere'] = df_a['Is_InAttesa'] & df_a['Ha_Verbale'] & (~df_a['Approvato'])
    df_a['Pronta_Fattura'] = df_a['Is_InAttesa'] & df_a['Ha_Verbale'] & df_a['Approvato']

    # Raggruppamento per data pianificata
    giornaliero = df_a.groupby('data_pianificata').agg(
        Pianificate=('numero_verifica', 'count'),
        Fatturate=('Is_Fatturata', 'sum'),
        In_Attesa=('Is_InAttesa', 'sum'),
        Non_Piu_Presenti=('Is_NPP', 'sum'),
        Senza_Verbale=('Senza_Verbale', 'sum'),
        In_Attesa_Ingegnere=('In_Attesa_Ingegnere', 'sum'),
        Pronta_Fattura=('Pronta_Fattura', 'sum')
    ).reset_index()

    giornaliero['% Fatturato'] = (giornaliero['Fatturate'] / giornaliero['Pianificate'] * 100).round(1)

    giornaliero.columns = [
        'Data Pianificata', 'Pianificate (T0)', 'Fatturate', 'In Attesa', 'Non Più Presenti',
        'In Attesa: Senza Verbale', 'In Attesa: Firma Ing.', 'In Attesa: Pronta Fattura', '% Fatturato'
    ]

    # Riordinamento colonne
    return giornaliero[[
        'Data Pianificata', 'Pianificate (T0)', 'Fatturate', '% Fatturato', 'In Attesa',
        'In Attesa: Senza Verbale', 'In Attesa: Firma Ing.', 'In Attesa: Pronta Fattura', 'Non Più Presenti'
    ]]

def costruisci_excel(df_filtrato, df_extra, descrizione_periodo):
    """Genera il file Excel con la nuova struttura focalizzata sui dati giornalieri divisi per reparto."""
    buf = BytesIO()
    
    df_asc = df_filtrato[df_filtrato['tipo_impianto'] == TIPO_A]
    df_mat = df_filtrato[df_filtrato['tipo_impianto'] == TIPO_E]
    
    g_tot = genera_report_giornaliero(df_filtrato)
    g_asc = genera_report_giornaliero(df_asc) if not df_asc.empty else pd.DataFrame()
    g_mat = genera_report_giornaliero(df_mat) if not df_mat.empty else pd.DataFrame()

    col_det = ['numero_verifica', 'protocollo', 'tipo_impianto', 'codice_impianto', 'data_pianificata', 'stato',
               'verbale', 'esito_verifica', 'verbale_da_inviare', 'fattura', 'data_t0', 'data_t1']

    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        pd.DataFrame({
            'Voce': ['Periodo di pianificazione', 'Generato il'],
            'Valore': [descrizione_periodo, datetime.now().strftime('%d/%m/%Y %H:%M')]
        }).to_excel(writer, sheet_name='Info', index=False)
        
        g_tot.to_excel(writer, sheet_name='Riepilogo Totale', index=False)
        if not g_asc.empty:
            g_asc.to_excel(writer, sheet_name='Ascensori', index=False)
        if not g_mat.empty:
            g_mat.to_excel(writer, sheet_name='Messa a Terra', index=False)
            
        df_filtrato[df_filtrato['stato'] == 'In Attesa'][col_det].rename(columns=RINOMINA).to_excel(
            writer, sheet_name='In Attesa', index=False)
        df_filtrato[df_filtrato['stato'] == 'Non Più Presente'][col_det].rename(columns=RINOMINA).to_excel(
            writer, sheet_name='Non Più Presenti', index=False)
        df_extra.to_excel(writer, sheet_name='Fuori Programmazione', index=False)
        df_filtrato[col_det].rename(columns=RINOMINA).to_excel(writer, sheet_name='Dettaglio Completo', index=False)
        
        for ws in writer.book.worksheets:
            for col in ws.columns:
                larghezza = max((len(str(c.value)) for c in col if c.value is not None), default=8) + 2
                ws.column_dimensions[col[0].column_letter].width = min(larghezza, 45)
                
    return buf.getvalue()

def mostra_sezione_report_giornaliero(df_sub, nome_titolo):
    """Visualizza KPI, tabella di dettaglio e grafico a barre per un reparto o per il totale."""
    if df_sub.empty:
        st.info(f"Nessuna verifica trovata per {nome_titolo} nel periodo selezionato.")
        return

    report_df = genera_report_giornaliero(df_sub)

    tot_pian = report_df['Pianificate (T0)'].sum()
    tot_fat = report_df['Fatturate'].sum()
    tot_att = report_df['In Attesa'].sum()
    tot_npp = report_df['Non Più Presenti'].sum()
    pct_fat = (tot_fat / tot_pian * 100) if tot_pian else 0.0

    # 1. Metriche sintetiche (KPI)
    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric("Pianificate (T0)", tot_pian)
    k2.metric("Fatturate", tot_fat, f"{pct_fat:.1f}% conv.")
    k3.metric("In Attesa di Fattura", tot_att)
    k4.metric("Non Più Presenti", tot_npp)
    k5.metric("% Conversione", f"{pct_fat:.1f}%")

    st.markdown(f"##### 📊 Dettaglio Giornaliero — {nome_titolo}")
    st.dataframe(report_df, use_container_width=True, hide_index=True)

    st.markdown(f"##### 📈 Grafico di Conversione Giornaliera — {nome_titolo}")
    st.bar_chart(
        report_df.set_index('Data Pianificata')[
            ['Fatturate', 'In Attesa', 'Non Più Presenti']
        ]
    )

# =========================================================
# ==== INTERFACCIA UTENTE E AUTENTICAZIONE ====
# =========================================================
def check_password():
    if "password_correct" not in st.session_state:
        st.session_state["password_correct"] = False
    if not st.session_state["password_correct"]:
        st.title("🔒 Accesso Riservato")
        st.markdown("Inserisci la password per accedere al gestionale delle verifiche.")
        pwd = st.text_input("Password", type="password")
        if st.button("Accedi"):
            if pwd == PASSWORD_ACCESSO:
                st.session_state["password_correct"] = True
                st.rerun()
            else:
                st.error("Password errata. Riprova.")
        return False
    return True

if not check_password():
    st.stop()

init_db()

# ---------------- SIDEBAR ----------------
with st.sidebar:
    st.header("💾 Gestione Sicurezza Dati")
    st.info("Prima di ogni salvataggio l'app crea un backup automatico nella cartella 'backups'.")
    ub = ultimo_backup()
    st.caption(f"Ultimo backup automatico: {ub.strftime('%d/%m/%Y %H:%M')}" if ub else "Nessun backup automatico ancora creato.")
    if os.path.exists(DB_NAME):
        with open(DB_NAME, "rb") as f:
            contenuto_db = f.read()
        st.download_button(
            label="📥 Scarica Backup Database (.db)",
            data=contenuto_db,
            file_name=f"backup_verifiche_{datetime.now().strftime('%Y%m%d')}.db",
            mime="application/octet-stream"
        )
    st.divider()
    st.subheader("Ripristina Backup")
    uploaded_db = st.file_uploader("Carica un file .db salvato in precedenza", type=["db"])
    if uploaded_db is not None:
        raw_db = uploaded_db.getvalue()
        db_valido, msg_db = valida_db_caricato(raw_db)
        if not db_valido:
            st.error(msg_db)
        else:
            st.success(msg_db)
            if st.button("⚠️ Sovrascrivi e Ripristina Database"):
                crea_backup("pre-ripristino")
                with open(DB_NAME, "wb") as f:
                    f.write(raw_db)
                init_db()
                logger.info("Database ripristinato da file caricato")
                st.success("Database ripristinato con successo. Ricarica la pagina.")
    st.divider()
    if st.button("🚪 Esci (Logout)"):
        st.session_state["password_correct"] = False
        st.rerun()

# ---------------- INTERFACCIA PRINCIPALE ----------------
st.title("📊 Gestione & Riconciliazione Verifiche")
st.markdown("Monitoraggio quotidiano programmate, fatturate, in attesa e assenti per reparto ($T_0 \\to T_1$)")

tabs = st.tabs(["📥 1. Import Programmate (T0)", "🔄 2. Import Consuntivo (T1)", "📊 3. Report Giornaliero per Reparto"])

# ---------------- TAB 1: IMPORT T0 ----------------
with tabs[0]:
    st.header("1. Caricamento Estrapolazione Programmate (T0)")
    c1, c2 = st.columns([1, 2])
    data_caricamento_t0 = c1.date_input("Data di acquisizione (T0):", datetime.today(), key="d_t0")
    file_t0 = st.file_uploader("Trascina il file Excel delle Programmate", type=["xlsx", "xls"], key="file_t0")
    if file_t0:
        df_raw_t0 = leggi_excel(file_t0)
        req_cols = ['Numero verifica', 'Protocollo', 'Codice impianto', 'Codice fiscale', 'Data pianificata']
        mancanti = [c for c in req_cols if c not in df_raw_t0.columns]
        if mancanti:
            st.error(f"Errore: nel file mancano le colonne: {', '.join(mancanti)}.")
        else:
            df_t0, info0 = prepara_t0(df_raw_t0)
            with db_connection() as conn:
                db_stato = carica_stato_db(conn)
                extra_nums = carica_numeri_extra(conn)
            piano0 = pianifica_t0(df_t0, db_stato, extra_nums)
            n_a = int((df_t0['tipo'] == TIPO_A).sum())
            n_e = int((df_t0['tipo'] == TIPO_E).sum())
            st.success(f"File letto. Verifiche valide: **{len(df_t0)}** (/A: {n_a} | /E: {n_e}) su {info0['righe_file']} righe.")
            
            p1, p2, p3, p4 = st.columns(4)
            p1.metric("Nuove verifiche", piano0['nuove'])
            p2.metric("Già presenti nel database", piano0['presenti'])
            p3.metric("Con data pianificata modificata", piano0['date_cambiate'])
            p4.metric("Prima 'fuori programmazione'", piano0['da_extra'])
            
            st.dataframe(df_t0.rename(columns={'num': 'Numero verifica', 'protocollo': 'Protocollo',
                                               'tipo': 'Reparto', 'data_pian': 'Data pianificata'})[
                ['Numero verifica', 'Protocollo', 'Reparto', 'Data pianificata']].head())
            
            if st.button("💾 Salva Programmazione (T0) nel Database", type="primary", disabled=df_t0.empty):
                crea_backup("pre-T0")
                dt0_str = data_caricamento_t0.strftime('%Y-%m-%d')
                with db_connection(commit=True) as conn:
                    applica_t0(conn, df_t0, dt0_str)
                st.success(f"Dati salvati con successo. Data di caricamento originaria ($T_0$): {dt0_str}")

# ---------------- TAB 2: IMPORT T1 ----------------
with tabs[1]:
    st.header("2. Riconciliazione Consuntivo / Emissione Fatture (T1)")
    c1, c2 = st.columns([1, 2])
    data_caricamento_t1 = c1.date_input("Data del controllo consuntivo (T1):", datetime.today(), key="d_t1")
    file_t1 = st.file_uploader("Trascina il file Excel aggiornato con Fatture", type=["xlsx", "xls"], key="file_t1")
    if file_t1:
        df_raw_t1 = leggi_excel(file_t1)
        mancanti = [c for c in ['Numero verifica', 'Data pianificata'] if c not in df_raw_t1.columns]
        if mancanti:
            st.error(f"Il file deve contenere 'Numero verifica' e 'Data pianificata'. Mancano: {', '.join(mancanti)}.")
        else:
            d1, info1 = prepara_t1(df_raw_t1)
            with db_connection() as conn:
                db_stato = carica_stato_db(conn)
            rilevati = rileva_reparti_t1(d1, db_stato)
            tot_rilevati = sum(rilevati.values())
            default_reparti = [t for t in (TIPO_A, TIPO_E) if tot_rilevati and rilevati.get(t, 0) / tot_rilevati >= 0.10]
            reparti_coperti = st.multiselect(
                "Reparti coperti da questo file T1:",
                [TIPO_A, TIPO_E], default=default_reparti,
                key=f"reparti_t1_{file_t1.name}_{file_t1.size}"
            )
            piano = pianifica_t1(d1, data_caricamento_t1.date() if isinstance(data_caricamento_t1, datetime) else data_caricamento_t1,
                                 db_stato, reparti_coperti)
            
            st.markdown("### 🔍 Anteprima riconciliazione")
            riepilogo, quadra = riepilogo_piano(piano, info1)
            st.dataframe(riepilogo, hide_index=True)
            
            k1, k2, k3 = st.columns(3)
            k1.metric("Verifiche T0 che diventerebbero 'Non Più Presenti'", len(piano['npp']))
            k2.metric("Su verifiche 'In Attesa' (stesse date e reparti)", piano['base_attesa'])
            k3.metric("Tornerebbero 'In Attesa'", len(piano['riattivate']))
            
            bloccato = False
            conferma = True
            if piano['oltre_soglia']:
                conferma = st.checkbox(
                    f"⚠ Il {piano['pct_npp']:.0%} delle verifiche 'In Attesa' con queste date risulta assente dal T1. Confermo che il file è completo.",
                    key="conferma_npp"
                )
            
            if st.button("⚡ Esegui Riconciliazione Automatica", type="primary", disabled=(bloccato or not conferma)):
                crea_backup("pre-T1")
                with db_connection(commit=True) as conn:
                    applica_t1(conn, piano)
                st.success(f"Riconciliazione completata rispetto alla data $T_1 = {piano['dt1_str']}$!")

# ---------------- TAB 3: REPORT GIORNALIERO E RIELABORATO ----------------
with tabs[2]:
    st.header("3. Rendicontazione Giornaliera e Stato Avanzamento per Reparto")
    with db_connection() as conn:
        df_db = pd.read_sql_query("SELECT * FROM verifiche", conn)
        df_extra = pd.read_sql_query("SELECT * FROM verifiche_extra", conn)
    
    if not df_db.empty:
        df_db['data_pianificata_dt'] = pd.to_datetime(df_db['data_pianificata'], errors='coerce')
        valid_dates = df_db['data_pianificata_dt'].dropna()
        if not valid_dates.empty:
            min_date = valid_dates.min().date()
            max_date = valid_dates.max().date()
        else:
            min_date = datetime.today().date()
            max_date = datetime.today().date()
            
        st.markdown("### 📅 Filtro Periodo Pianificazione")
        c1, c2 = st.columns([1, 2])
        filtro_tipo = c1.radio("Scegli l'ampiezza dell'analisi:", ["Tutto il database", "Seleziona Range Personalizzato"])
        df_extra_filtrato = df_extra.copy()
        
        if filtro_tipo == "Tutto il database":
            df_filtrato = df_db.copy()
            descr_periodo = f"{min_date.strftime('%d/%m/%Y')} - {max_date.strftime('%d/%m/%Y')} (intero storico)"
        else:
            date_range = c2.date_input(
                "Seleziona la data di Inizio e Fine:",
                value=(min_date, max_date),
                min_value=min_date,
                max_value=max_date
            )
            if isinstance(date_range, tuple) and len(date_range) == 2:
                start_date, end_date = date_range
                mask = (df_db['data_pianificata_dt'].dt.date >= start_date) & (df_db['data_pianificata_dt'].dt.date <= end_date)
                df_filtrato = df_db.loc[mask]
                if not df_extra.empty:
                    extra_dt = pd.to_datetime(df_extra['data_pianificata'], errors='coerce').dt.date
                    df_extra_filtrato = df_extra.loc[(extra_dt >= start_date) & (extra_dt <= end_date)]
                descr_periodo = f"{start_date.strftime('%d/%m/%Y')} - {end_date.strftime('%d/%m/%Y')}"
            else:
                st.warning("Seleziona anche la data di fine dal calendario.")
                st.stop()
                
        st.download_button(
            "📤 Esporta report in Excel",
            data=costruisci_excel(df_filtrato, df_extra_filtrato, descr_periodo),
            file_name=f"report_verifiche_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.markdown("---")
        
        # Schede per Reparto
        tab_asc, tab_mat, tab_tot = st.tabs([
            "🛗 Ascensori (Sez. /A)", 
            "⚡ Messa a Terra (Sez. /E)", 
            "🌐 Totale Complessivo"
        ])

        with tab_asc:
            df_asc = df_filtrato[df_filtrato['tipo_impianto'] == TIPO_A]
            mostra_sezione_report_giornaliero(df_asc, "Reparto Ascensori (Sezionale /A)")

        with tab_mat:
            df_mat = df_filtrato[df_filtrato['tipo_impianto'] == TIPO_E]
            mostra_sezione_report_giornaliero(df_mat, "Reparto Messa a Terra (Sezionale /E)")

        with tab_tot:
            mostra_sezione_report_giornaliero(df_filtrato, "Totale Aziendale Complessivo")

        st.markdown("---")
        st.subheader("📋 Dettaglio Completo Verifiche Filtrate")
        df_display = df_filtrato[['numero_verifica', 'data_pianificata', 'tipo_impianto', 'codice_impianto', 
                                  'verbale', 'esito_verifica', 'verbale_da_inviare', 'fattura', 'stato',
                                  'data_t0', 'data_t1']]
        st.dataframe(df_display.rename(columns=RINOMINA), hide_index=True)
    else:
        st.info("Nessun dato salvato nel sistema. Effettua un primo caricamento nel Tab 1.")
