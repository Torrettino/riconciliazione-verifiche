import streamlit as st
import pandas as pd
import sqlite3
from datetime import datetime

# ---------------------------------------------------------
# CONFIGURAZIONE PAGINA E DATABASE
# ---------------------------------------------------------
st.set_page_config(page_title="Riconciliazione Verifiche Impianti", layout="wide")

def get_db_connection():
    conn = sqlite3.connect("verifiche.db")
    return conn

def init_db():
    conn = get_db_connection()
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
    conn.commit()
    conn.close()

init_db()

# ---------------------------------------------------------
# INTERFACCIA PRINCIPALE
# ---------------------------------------------------------
st.title("📊 Gestione & Riconciliazione Verifiche Impianti")
st.markdown("Monitoraggio del rateo di conversione tra **Verifiche Programmate** e **Fatture Emesse**")

tabs = st.tabs(["📥 1. Import Programmate (T0)", "🔄 2. Import Consuntivo (T1)", "📊 3. Report Sintetico"])

# ---------------------------------------------------------
# TAB 1: IMPORT PROGRAMMATE
# ---------------------------------------------------------
with tabs[0]:
    st.header("1. Caricamento Estrapolazione Programmate (T0)")
    file_t0 = st.file_uploader("Trascina o seleziona il file Excel delle Programmate", type=["xlsx", "xls"], key="t0")
    
    if file_t0 is not None:
        df_t0 = pd.read_excel(file_t0)
        req_cols = ['Numero verifica', 'Protocollo', 'Codice impianto', 'Codice fiscale', 'Data pianificata', 'Importo']
        
        if all(col in df_t0.columns for col in req_cols):
            df_t0['tipo_impianto'] = df_t0['Protocollo'].astype(str).apply(
                lambda x: 'Ascensori (DPR 162/99)' if x.strip().endswith('/A') else ('Messa a Terra (DPR 462/01)' if x.strip().endswith('/E') else 'Altro')
            )
            df_t0['Data pianificata'] = pd.to_datetime(df_t0['Data pianificata']).dt.strftime('%Y-%m-%d')
            data_pianificata = df_t0['Data pianificata'].iloc[0]
            
            st.success(f"File riconosciuto! Data Pianificata: **{data_pianificata}** | Verifiche totali: **{len(df_t0)}**")
            st.dataframe(df_t0[['Numero verifica', 'Protocollo', 'tipo_impianto', 'Codice impianto', 'Codice fiscale', 'Data pianificata', 'Importo']].head())
            
            if st.button("💾 Salva Programmazione nel Database"):
                conn = get_db_connection()
                cursor = conn.cursor()
                for _, row in df_t0.iterrows():
                    cursor.execute('''
                        INSERT INTO verifiche (numero_verifica, protocollo, codice_impianto, codice_fiscale, data_pianificata, importo, tipo_impianto, fattura, stato, data_t0)
                        VALUES (?, ?, ?, ?, ?, ?, ?, NULL, 'In Attesa', ?)
                        ON CONFLICT(numero_verifica) DO UPDATE SET
                            protocollo=excluded.protocollo,
                            codice_impianto=excluded.codice_impianto,
                            codice_fiscale=excluded.codice_fiscale,
                            data_pianificata=excluded.data_pianificata,
                            importo=excluded.importo,
                            tipo_impianto=excluded.tipo_impianto
                    ''', (
                        int(row['Numero verifica']),
                        str(row['Protocollo']),
                        str(row['Codice impianto']),
                        str(row['Codice fiscale']),
                        row['Data pianificata'],
                        float(row['Importo']) if pd.notna(row['Importo']) else 0.0,
                        row['tipo_impianto'],
                        datetime.now().strftime('%Y-%m-%d')
                    ))
                conn.commit()
                conn.close()
                st.balloons()
                st.success(f"Programmazione del {data_pianificata} salvata correttamente!")
        else:
            st.error("Il file Excel non contiene tutte le colonne richieste.")

# ---------------------------------------------------------
# TAB 2: IMPORT CONSUNTIVO
# ---------------------------------------------------------
with tabs[1]:
    st.header("2. Riconciliazione Consuntivo a 1 Mese (T1)")
    file_t1 = st.file_uploader("Trascina o seleziona il file Excel aggiornato dopo un mese", type=["xlsx", "xls"], key="t1")
    
    if file_t1 is not None:
        df_t1 = pd.read_excel(file_t1)
        if 'Numero verifica' in df_t1.columns and 'Fattura' in df_t1.columns:
            st.info(f"File consuntivo letto. Righe trovate: **{len(df_t1)}**")
            
            if st.button("⚡ Esegui Riconciliazione Automatica"):
                conn = get_db_connection()
                cursor = conn.cursor()
                
                # Registra le fatture emesse
                verifiche_presenti_t1 = set()
                for _, row in df_t1.iterrows():
                    num_ver = int(row['Numero verifica'])
                    verifiche_presenti_t1.add(num_ver)
                    num_fat = str(row['Fattura']).strip() if pd.notna(row['Fattura']) and str(row['Fattura']).lower() != 'nan' else None
                    
                    if num_fat:
                        cursor.execute('''
                            UPDATE verifiche 
                            SET fattura = ?, stato = 'Fatturata', data_t1 = ?
                            WHERE numero_verifica = ?
                        ''', (num_fat, datetime.now().strftime('%Y-%m-%d'), num_ver))
                
                # Identifica le verifiche non più presenti nel file T1 rispetto alla data pianificata
                cursor.execute("SELECT numero_verifica, data_pianificata FROM verifiche WHERE stato != 'Fatturata'")
                in_attesa = cursor.fetchall()
                for num_ver, d_pian in in_attesa:
                    if num_ver not in verifiche_presenti_t1:
                        cursor.execute("UPDATE verifiche SET stato = 'Non Più Presente' WHERE numero_verifica = ?", (num_ver,))
                
                conn.commit()
                conn.close()
                st.success("Riconciliazione completata con successo!")
        else:
            st.error("Il file deve contenere le colonne 'Numero verifica' e 'Fattura'.")

# ---------------------------------------------------------
# TAB 3: REPORT SINTETICO
# ---------------------------------------------------------
with tabs[2]:
    st.header("3. Report Sintetico e Rateo di Conversione")
    conn = get_db_connection()
    df_db = pd.read_sql_query("SELECT * FROM verifiche", conn)
    conn.close()
    
    if not df_db.empty:
        date_list = sorted(df_db['data_pianificata'].unique(), reverse=True)
        data_sel = st.selectbox("Seleziona la Data Pianificata da analizzare:", date_list)
        
        if data_sel:
            df_giorno = df_db[df_db['data_pianificata'] == data_sel]
            
            tot_prog = len(df_giorno)
            tot_fat = len(df_giorno[df_giorno['stato'] == 'Fatturata'])
            tot_att = len(df_giorno[df_giorno['stato'] == 'In Attesa'])
            tot_non_pres = len(df_giorno[df_giorno['stato'] == 'Non Più Presente'])
            
            perc_conv = (tot_fat / tot_prog * 100) if tot_prog > 0 else 0.0
            
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("1. Totale Programmate", tot_prog)
            m2.metric("2. Diventate Fattura", tot_fat, f"{perc_conv:.1f}%")
            m3.metric("3. In Attesa di Fattura", tot_att)
            m4.metric("4. Non Più Presenti", tot_non_pres)
            
            st.divider()
            st.subheader("Dettaglio Verifiche del Giorno")
            st.dataframe(df_giorno[['numero_verifica', 'protocollo', 'tipo_impianto', 'codice_impianto', 'codice_fiscale', 'fattura', 'stato', 'importo']])
    else:
        st.info("Nessun dato salvato nel sistema. Effettua un primo caricamento nel Tab 1.")
