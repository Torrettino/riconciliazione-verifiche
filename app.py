import streamlit as st
import pandas as pd
import sqlite3
import os
from datetime import datetime
from io import BytesIO

# ---------------------------------------------------------
# CONFIGURAZIONE PAGINA
# ---------------------------------------------------------
st.set_page_config(page_title="Riconciliazione Verifiche Impianti", page_icon="📊", layout="wide")

DB_NAME = "verifiche.db"

def get_db_connection():
    return sqlite3.connect(DB_NAME)

def init_db():
    conn = get_db_connection()
    cursor = conn.cursor()
    
    # 1. Crea la tabella base se è la primissima volta
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
    
    # 2. Controlla le colonne esistenti (Migrazione automatica senza perdere dati)
    cursor.execute("PRAGMA table_info(verifiche)")
    colonne_esistenti = [colonna[1] for colonna in cursor.fetchall()]
    
    # Se mancano le nuove colonne temporali, le aggiunge
    if 'giorni_trascorsi' not in colonne_esistenti:
        cursor.execute("ALTER TABLE verifiche ADD COLUMN giorni_trascorsi INTEGER")
    if 'fascia_tempo' not in colonne_esistenti:
        cursor.execute("ALTER TABLE verifiche ADD COLUMN fascia_tempo TEXT")
        
    conn.commit()
    conn.close()

init_db()

# ---------------------------------------------------------
# SIDEBAR: BACKUP E SALVATAGGIO DATI (Protezione Cloud)
# ---------------------------------------------------------
with st.sidebar:
    st.header("💾 Gestione Sicurezza Dati")
    st.info("Su piattaforme Cloud, scarica periodicamente il database per avere un backup sicuro sul tuo PC.")
    
    # 1. Download Database
    if os.path.exists(DB_NAME):
        with open(DB_NAME, "rb") as f:
            st.download_button(
                label="📥 Scarica Backup Database (.db)",
                data=f,
                file_name=f"backup_verifiche_{datetime.now().strftime('%Y%m%d')}.db",
                mime="application/octet-stream"
            )
            
    st.divider()
    
    # 2. Upload Database (Ripristino)
    st.subheader("Ripristina Backup")
    uploaded_db = st.file_uploader("Carica un file .db salvato in precedenza", type=["db"])
    if uploaded_db is not None:
        if st.button("⚠️ Sovrascrivi e Ripristina Database"):
            with open(DB_NAME, "wb") as f:
                f.write(uploaded_db.getbuffer())
            st.success("Database ripristinato con successo! Ricarica la pagina.")

# ---------------------------------------------------------
# INTERFACCIA PRINCIPALE
# ---------------------------------------------------------
st.title("📊 Gestione & Riconciliazione Verifiche")
st.markdown("Monitoraggio conversioni in fattura e reportistica temporale ($T_0 \\to T_1$)")

tabs = st.tabs(["📥 1. Import Programmate (T0)", "🔄 2. Import Consuntivo (T1)", "📊 3. Report & Tempi"])

# ---------------------------------------------------------
# TAB 1: IMPORT PROGRAMMATE (T0)
# ---------------------------------------------------------
with tabs[0]:
    st.header("1. Caricamento Estrapolazione Programmate (T0)")
    
    c1, c2 = st.columns([1, 2])
    data_caricamento_t0 = c1.date_input("Data di acquisizione (T0):", datetime.today(), key="d_t0")
    file_t0 = st.file_uploader("Trascina il file Excel delle Programmate (Es. 08/09/2026)", type=["xlsx", "xls"], key="file_t0")
    
    if file_t0:
        df_t0 = pd.read_excel(file_t0)
        req_cols = ['Numero verifica', 'Protocollo', 'Codice impianto', 'Codice fiscale', 'Data pianificata', 'Importo']
        
        if all(col in df_t0.columns for col in req_cols):
            # Normalizzazione stringhe e date
            df_t0['Protocollo'] = df_t0['Protocollo'].astype(str).str.strip().str.upper()
            df_t0['tipo_impianto'] = df_t0['Protocollo'].apply(
                lambda x: 'Ascensori (DPR 162/99)' if x.endswith('/A') else ('Messa a Terra (DPR 462/01)' if x.endswith('/E') else 'Altro')
            )
            df_t0['Data pianificata'] = pd.to_datetime(df_t0['Data pianificata'], errors='coerce').dt.strftime('%Y-%m-%d')
            
            # Filtro righe valide
            df_t0 = df_t0.dropna(subset=['Numero verifica', 'Data pianificata'])
            
            st.success(f"File letto. Trovate **{len(df_t0)}** verifiche programmate.")
            st.dataframe(df_t0[['Numero verifica', 'Protocollo', 'tipo_impianto', 'Data pianificata']].head())
            
            if st.button("💾 Salva Programmazione (T0) nel Database", type="primary"):
                conn = get_db_connection()
                cursor = conn.cursor()
                dt0_str = data_caricamento_t0.strftime('%Y-%m-%d')
                
                for _, row in df_t0.iterrows():
                    # Inserimento sicuro. Se esiste già, aggiorna i dati ma MANTIENE la data T0 originaria.
                    cursor.execute('''
                        INSERT INTO verifiche (numero_verifica, protocollo, codice_impianto, codice_fiscale, data_pianificata, importo, tipo_impianto, fattura, stato, data_t0)
                        VALUES (?, ?, ?, ?, ?, ?, ?, NULL, 'In Attesa', ?)
                        ON CONFLICT(numero_verifica) DO UPDATE SET
                            protocollo=excluded.protocollo,
                            codice_impianto=excluded.codice_impianto,
                            codice_fiscale=excluded.codice_fiscale,
                            data_pianificata=excluded.data_pianificata,
                            importo=excluded.importo,
                            tipo_impianto=excluded.tipo_impianto,
                            stato = CASE WHEN fattura IS NULL THEN 'In Attesa' ELSE stato END
                    ''', (
                        int(row['Numero verifica']),
                        str(row['Protocollo']),
                        str(row['Codice impianto']),
                        str(row['Codice fiscale']),
                        row['Data pianificata'],
                        float(row['Importo']) if pd.notna(row['Importo']) else 0.0,
                        row['tipo_impianto'],
                        dt0_str
                    ))
                conn.commit()
                conn.close()
                st.success(f"Dati salvati con successo. Data di caricamento originaria ($T_0$): {dt0_str}")
        else:
            st.error("Errore: Il file caricato non contiene tutte le colonne necessarie.")

# ---------------------------------------------------------
# TAB 2: IMPORT CONSUNTIVO (T1) E RICONCILIAZIONE
# ---------------------------------------------------------
with tabs[1]:
    st.header("2. Riconciliazione Consuntivo / Emissione Fatture (T1)")
    
    c1, c2 = st.columns([1, 2])
    data_caricamento_t1 = c1.date_input("Data del controllo consuntivo (T1):", datetime.today(), key="d_t1")
    file_t1 = st.file_uploader("Trascina il file Excel aggiornato (Es. 08/10/2026)", type=["xlsx", "xls"], key="file_t1")
    
    if file_t1:
        df_t1 = pd.read_excel(file_t1)
        if 'Numero verifica' in df_t1.columns and 'Fattura' in df_t1.columns and 'Data pianificata' in df_t1.columns:
            df_t1['Data pianificata'] = pd.to_datetime(df_t1['Data pianificata'], errors='coerce').dt.strftime('%Y-%m-%d')
            date_presenti_t1 = df_t1['Data pianificata'].dropna().unique().tolist()
            
            st.info(f"File consuntivo analizzato. Verifiche: **{len(df_t1)}** | Date coinvolte: **{len(date_presenti_t1)}**")
            
            if st.button("⚡ Esegui Riconciliazione Automatica", type="primary"):
                conn = get_db_connection()
                cursor = conn.cursor()
                dt1_str = data_caricamento_t1.strftime('%Y-%m-%d')
                dt1_obj = data_caricamento_t1
                
                verifiche_presenti_t1 = set()
                
                # 1. Aggiornamento Fatture ed Età Temporale
                for _, row in df_t1.iterrows():
                    num_ver = int(row['Numero verifica'])
                    verifiche_presenti_t1.add(num_ver)
                    
                    fatt_raw = str(row['Fattura']).strip()
                    num_fat = fatt_raw if pd.notna(row['Fattura']) and fatt_raw.lower() != 'nan' and fatt_raw != '' else None
                    
                    if num_fat:
                        cursor.execute("SELECT data_t0 FROM verifiche WHERE numero_verifica = ?", (num_ver,))
                        res = cursor.fetchone()
                        
                        giorni = None
                        fascia = "N/D"
                        if res and res[0]:
                            dt0_obj = datetime.strptime(res[0], '%Y-%m-%d').date()
                            giorni = (dt1_obj - dt0_obj).days
                            
                            if giorni <= 7: fascia = "1. Entro 1 Settimana"
                            elif giorni <= 14: fascia = "2. Entro 2 Settimane"
                            elif giorni <= 30: fascia = "3. Entro 1 Mese"
                            else: fascia = "4. Oltre 1 Mese"
                        
                        cursor.execute('''
                            UPDATE verifiche 
                            SET fattura = ?, stato = 'Fatturata', data_t1 = ?, giorni_trascorsi = ?, fascia_tempo = ?
                            WHERE numero_verifica = ?
                        ''', (num_fat, dt1_str, giorni, fascia, num_ver))
                
                # 2. Identificazione Chirurgica "Non Più Presenti"
                if date_presenti_t1:
                    placeholders = ','.join('?' for _ in date_presenti_t1)
                    query = f"SELECT numero_verifica FROM verifiche WHERE stato != 'Fatturata' AND data_pianificata IN ({placeholders})"
                    cursor.execute(query, date_presenti_t1)
                    in_attesa_db = cursor.fetchall()
                    
                    for (nv,) in in_attesa_db:
                        if nv not in verifiche_presenti_t1:
                            cursor.execute("UPDATE verifiche SET stato = 'Non Più Presente' WHERE numero_verifica = ?", (nv,))
                
                conn.commit()
                conn.close()
                st.success(f"Riconciliazione completata con successo rispetto alla data $T_1 = {dt1_str}$!")
        else:
            st.error("Il file deve contenere 'Numero verifica', 'Fattura' e 'Data pianificata'.")

# ---------------------------------------------------------
# TAB 3: REPORT SINTETICO E TEMPORALE
# ---------------------------------------------------------
with tabs[2]:
    st.header("3. Report Sintetico & Tempi di Conversione")
    conn = get_db_connection()
    df_db = pd.read_sql_query("SELECT * FROM verifiche", conn)
    conn.close()
    
    if not df_db.empty:
        date_list = sorted(df_db['data_pianificata'].unique(), reverse=True)
        data_sel = st.selectbox("📅 Seleziona la Data Pianificata da analizzare:", ["Tutte le date"] + date_list)
        
        # Filtraggio
        if data_sel == "Tutte le date":
            df_giorno = df_db.copy()
        else:
            df_giorno = df_db[df_db['data_pianificata'] == data_sel]
            
        tot_prog = len(df_giorno)
        tot_fat = len(df_giorno[df_giorno['stato'] == 'Fatturata'])
        tot_att = len(df_giorno[df_giorno['stato'] == 'In Attesa'])
        tot_non_pres = len(df_giorno[df_giorno['stato'] == 'Non Più Presente'])
        perc_conv = (tot_fat / tot_prog * 100) if tot_prog > 0 else 0.0
        
        # -- 1. RIEPILOGO GENERALE --
        st.subheader("1. Riepilogo Volumi")
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Totale Programmate", tot_prog)
        m2.metric("Diventate Fattura", tot_fat, f"{perc_conv:.1f}% di conversione")
        m3.metric("In Attesa di Fattura", tot_att)
        m4.metric("Non Più Presenti", tot_non_pres)
        
        st.divider()
        
        # -- 2. CONFRONTO TEMPORALE --
        st.subheader(f"2. Età di Fatturazione (Dal caricamento $T_0$ al consuntivo $T_1$)")
        df_fatturate = df_giorno[df_giorno['stato'] == 'Fatturata']
        
        t1, t2, t3, t4 = st.columns(4)
        e_1_sett = len(df_fatturate[df_fatturate['fascia_tempo'] == '1. Entro 1 Settimana'])
        e_2_sett = len(df_fatturate[df_fatturate['fascia_tempo'] == '2. Entro 2 Settimane'])
        e_1_mese = len(df_fatturate[df_fatturate['fascia_tempo'] == '3. Entro 1 Mese'])
        oltre_mese = len(df_fatturate[df_fatturate['fascia_tempo'] == '4. Oltre 1 Mese'])
        
        t1.metric("≤ 7 Giorni", e_1_sett, f"{(e_1_sett/tot_fat*100):.1f}% del fatturato" if tot_fat>0 else "0%")
        t2.metric("8 - 14 Giorni", e_2_sett, f"{(e_2_sett/tot_fat*100):.1f}% del fatturato" if tot_fat>0 else "0%")
        t3.metric("15 - 30 Giorni", e_1_mese, f"{(e_1_mese/tot_fat*100):.1f}% del fatturato" if tot_fat>0 else "0%")
        t4.metric("> 30 Giorni", oltre_mese, f"{(oltre_mese/tot_fat*100):.1f}% del fatturato" if tot_fat>0 else "0%")
        
        st.divider()
        
        # -- 3. DATI GREZZI --
        st.subheader("Dettaglio Verifiche")
        df_display = df_giorno[['numero_verifica', 'data_pianificata', 'tipo_impianto', 'codice_impianto', 'fattura', 'stato', 'data_t0', 'data_t1', 'giorni_trascorsi', 'fascia_tempo', 'importo']]
        st.dataframe(df_display, use_container_width=True)
        
    else:
        st.info("Nessun dato salvato nel sistema. Effettua un primo caricamento nel Tab 1.")
