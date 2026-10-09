import streamlit as st
import pandas as pd
import sqlite3
import os
from datetime import datetime

# ---------------------------------------------------------
# CONFIGURAZIONE PAGINA E SICUREZZA
# ---------------------------------------------------------
st.set_page_config(page_title="Riconciliazione Verifiche Impianti", page_icon="📊", layout="wide")

PASSWORD_ACCESSO = "Elti2026!"  # <-- Modifica qui la tua password aziendale

TIPO_A = 'Ascensori (DPR 162/99)'
TIPO_E = 'Messa a Terra (DPR 462/01)'


def tipo_da_protocollo(protocollo):
    """Restituisce il reparto in base al sezionale; None se non è /A o /E."""
    p = str(protocollo).strip().upper()
    if p.endswith('/A'):
        return TIPO_A
    if p.endswith('/E'):
        return TIPO_E
    return None


def check_password():
    """Restituisce True se l'utente ha inserito la password corretta."""
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


# Blocca l'esecuzione se la password non è inserita
if not check_password():
    st.stop()

# ---------------------------------------------------------
# DATABASE SETUP
# ---------------------------------------------------------
DB_NAME = "verifiche.db"


def get_db_connection():
    return sqlite3.connect(DB_NAME)


def init_db():
    conn = get_db_connection()
    cursor = conn.cursor()

    # 1. Tabella base (verifiche programmate T0)
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

    # 2. Migrazione automatica colonne
    cursor.execute("PRAGMA table_info(verifiche)")
    colonne_esistenti = [colonna[1] for colonna in cursor.fetchall()]

    if 'giorni_trascorsi' not in colonne_esistenti:
        cursor.execute("ALTER TABLE verifiche ADD COLUMN giorni_trascorsi INTEGER")
    if 'fascia_tempo' not in colonne_esistenti:
        cursor.execute("ALTER TABLE verifiche ADD COLUMN fascia_tempo TEXT")

    # 3. Verifiche comparse nel T1 ma mai programmate nel T0 (report a parte)
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

    conn.commit()
    conn.close()


init_db()

# ---------------------------------------------------------
# SIDEBAR: BACKUP E SALVATAGGIO DATI
# ---------------------------------------------------------
with st.sidebar:
    st.header("💾 Gestione Sicurezza Dati")
    st.info("Scarica periodicamente il database per avere un backup sicuro sul tuo PC.")

    if os.path.exists(DB_NAME):
        with open(DB_NAME, "rb") as f:
            st.download_button(
                label="📥 Scarica Backup Database (.db)",
                data=f,
                file_name=f"backup_verifiche_{datetime.now().strftime('%Y%m%d')}.db",
                mime="application/octet-stream"
            )

    st.divider()

    st.subheader("Ripristina Backup")
    uploaded_db = st.file_uploader("Carica un file .db salvato in precedenza", type=["db"])
    if uploaded_db is not None:
        if st.button("⚠️ Sovrascrivi e Ripristina Database"):
            with open(DB_NAME, "wb") as f:
                f.write(uploaded_db.getbuffer())
            st.success("Database ripristinato con successo! Ricarica la pagina.")

    st.divider()
    if st.button("🚪 Esci (Logout)"):
        st.session_state["password_correct"] = False
        st.rerun()

# ---------------------------------------------------------
# INTERFACCIA PRINCIPALE
# ---------------------------------------------------------
st.title("📊 Gestione & Riconciliazione Verifiche")
st.markdown("Monitoraggio conversioni in fattura e reportistica temporale per reparto ($T_0 \\to T_1$)")

tabs = st.tabs(["📥 1. Import Programmate (T0)", "🔄 2. Import Consuntivo (T1)", "📊 3. Report & Tempi per Reparto"])

# ---------------------------------------------------------
# TAB 1: IMPORT PROGRAMMATE (T0)
# ---------------------------------------------------------
with tabs[0]:
    st.header("1. Caricamento Estrapolazione Programmate (T0)")

    c1, c2 = st.columns([1, 2])
    data_caricamento_t0 = c1.date_input("Data di acquisizione (T0):", datetime.today(), key="d_t0")
    file_t0 = st.file_uploader("Trascina il file Excel delle Programmate", type=["xlsx", "xls"], key="file_t0")

    if file_t0:
        df_t0 = pd.read_excel(file_t0)
        req_cols = ['Numero verifica', 'Protocollo', 'Codice impianto', 'Codice fiscale', 'Data pianificata', 'Importo']

        if all(col in df_t0.columns for col in req_cols):
            df_t0['Protocollo'] = df_t0['Protocollo'].astype(str).str.strip().str.upper()
            df_t0['tipo_impianto'] = df_t0['Protocollo'].apply(tipo_da_protocollo)

            # Si lavora solo con i sezionali /A e /E: gli altri protocolli vengono scartati
            n_scartate = int(df_t0['tipo_impianto'].isna().sum())
            df_t0 = df_t0[df_t0['tipo_impianto'].notna()].copy()

            df_t0['Data pianificata'] = pd.to_datetime(df_t0['Data pianificata'], errors='coerce').dt.strftime('%Y-%m-%d')
            df_t0 = df_t0.dropna(subset=['Numero verifica', 'Data pianificata'])

            st.success(f"File letto. Trovate **{len(df_t0)}** verifiche programmate (/A: {int((df_t0['tipo_impianto'] == TIPO_A).sum())} | /E: {int((df_t0['tipo_impianto'] == TIPO_E).sum())}).")
            if n_scartate > 0:
                st.warning(f"{n_scartate} righe con protocollo diverso da /A e /E sono state escluse.")
            st.dataframe(df_t0[['Numero verifica', 'Protocollo', 'tipo_impianto', 'Data pianificata']].head())

            if st.button("💾 Salva Programmazione (T0) nel Database", type="primary"):
                conn = get_db_connection()
                cursor = conn.cursor()
                dt0_str = data_caricamento_t0.strftime('%Y-%m-%d')

                for _, row in df_t0.iterrows():
                    # data_t0 NON viene aggiornata in caso di conflitto: resta la data originaria di prima acquisizione
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
    file_t1 = st.file_uploader("Trascina il file Excel aggiornato con Fatture", type=["xlsx", "xls"], key="file_t1")

    if file_t1:
        df_t1 = pd.read_excel(file_t1)
        if 'Numero verifica' in df_t1.columns and 'Fattura' in df_t1.columns and 'Data pianificata' in df_t1.columns:
            df_t1['Data pianificata'] = pd.to_datetime(df_t1['Data pianificata'], errors='coerce').dt.strftime('%Y-%m-%d')
            date_presenti_t1 = df_t1['Data pianificata'].dropna().unique().tolist()
            ha_protocollo = 'Protocollo' in df_t1.columns

            st.info(f"File consuntivo analizzato. Verifiche: **{len(df_t1)}** | Date coinvolte: **{len(date_presenti_t1)}**")
            if not ha_protocollo:
                st.caption("Il file T1 non ha la colonna 'Protocollo': le eventuali verifiche non programmate non potranno essere assegnate a un reparto.")

            if st.button("⚡ Esegui Riconciliazione Automatica", type="primary"):
                conn = get_db_connection()
                cursor = conn.cursor()
                dt1_str = data_caricamento_t1.strftime('%Y-%m-%d')
                dt1_obj = data_caricamento_t1

                verifiche_presenti_t1 = set()

                # Contatori per il controllo di quadratura
                nuove_fatturate = 0
                gia_fatturate = 0
                presenti_senza_fattura = 0
                extra_t1 = 0
                scartate_altro = 0

                for _, row in df_t1.iterrows():
                    num_ver = int(row['Numero verifica'])
                    verifiche_presenti_t1.add(num_ver)

                    fatt_raw = str(row['Fattura']).strip()
                    num_fat = fatt_raw if pd.notna(row['Fattura']) and fatt_raw.lower() != 'nan' and fatt_raw != '' else None

                    cursor.execute("SELECT data_t0, stato FROM verifiche WHERE numero_verifica = ?", (num_ver,))
                    res = cursor.fetchone()

                    # --- CASO A: verifica NON presente nel T0 -> report a parte ---
                    if res is None:
                        protocollo = ''
                        tipo = None
                        if ha_protocollo and pd.notna(row['Protocollo']):
                            protocollo = str(row['Protocollo']).strip().upper()
                            tipo = tipo_da_protocollo(protocollo)
                            if tipo is None:
                                scartate_altro += 1
                                continue

                        data_pian = row['Data pianificata'] if pd.notna(row['Data pianificata']) else None
                        cursor.execute('''
                            INSERT INTO verifiche_extra (numero_verifica, protocollo, tipo_impianto, data_pianificata, fattura, data_t1)
                            VALUES (?, ?, ?, ?, ?, ?)
                            ON CONFLICT(numero_verifica) DO UPDATE SET
                                fattura = COALESCE(excluded.fattura, verifiche_extra.fattura),
                                data_pianificata = excluded.data_pianificata
                        ''', (num_ver, protocollo, tipo, data_pian, num_fat, dt1_str))
                        extra_t1 += 1
                        continue

                    data_t0_db, stato_db = res

                    # --- CASO B: già fatturata in un T1 precedente -> non si tocca nulla ---
                    # (data_t1, giorni e fascia restano quelli della PRIMA fatturazione rilevata)
                    if stato_db == 'Fatturata':
                        gia_fatturate += 1
                        continue

                    # --- CASO C: nuova fattura rilevata ---
                    if num_fat:
                        giorni = None
                        fascia = "N/D"
                        if data_t0_db:
                            dt0_obj = datetime.strptime(data_t0_db, '%Y-%m-%d').date()
                            giorni = (dt1_obj - dt0_obj).days

                            if giorni <= 7:
                                fascia = "1. Entro 1 Settimana"
                            elif giorni <= 14:
                                fascia = "2. Entro 2 Settimane"
                            elif giorni <= 30:
                                fascia = "3. Entro 1 Mese"
                            else:
                                fascia = "4. Oltre 1 Mese"

                        cursor.execute('''
                            UPDATE verifiche
                            SET fattura = ?, stato = 'Fatturata', data_t1 = ?, giorni_trascorsi = ?, fascia_tempo = ?
                            WHERE numero_verifica = ?
                        ''', (num_fat, dt1_str, giorni, fascia, num_ver))
                        nuove_fatturate += 1

                    # --- CASO D: presente nel T1 ma ancora senza fattura ---
                    else:
                        if stato_db == 'Non Più Presente':
                            cursor.execute("UPDATE verifiche SET stato = 'In Attesa' WHERE numero_verifica = ?", (num_ver,))
                        presenti_senza_fattura += 1

                # --- Verifiche del T0 (stesse date pianificate del T1) sparite dal T1 ---
                non_piu_presenti = 0
                if date_presenti_t1:
                    placeholders = ','.join('?' for _ in date_presenti_t1)
                    query = f"SELECT numero_verifica FROM verifiche WHERE stato = 'In Attesa' AND data_pianificata IN ({placeholders})"
                    cursor.execute(query, date_presenti_t1)
                    in_attesa_db = cursor.fetchall()

                    for (nv,) in in_attesa_db:
                        if nv not in verifiche_presenti_t1:
                            cursor.execute("UPDATE verifiche SET stato = 'Non Più Presente' WHERE numero_verifica = ?", (nv,))
                            non_piu_presenti += 1

                conn.commit()
                conn.close()

                st.success(f"Riconciliazione completata rispetto alla data $T_1 = {dt1_str}$!")

                riepilogo = pd.DataFrame({
                    "Esito": [
                        "Nuove fatture rilevate",
                        "Già fatturate in precedenza (invariate)",
                        "Presenti nel T1 ma ancora senza fattura",
                        "Non programmate nel T0 (report a parte)",
                        "Escluse (protocollo diverso da /A e /E)",
                    ],
                    "Verifiche": [nuove_fatturate, gia_fatturate, presenti_senza_fattura, extra_t1, scartate_altro]
                })
                st.dataframe(riepilogo, hide_index=True, use_container_width=False)
                st.metric("Verifiche del T0 non più presenti nel T1", non_piu_presenti)

                # Controllo di quadratura: ogni riga del T1 deve avere un esito
                somma_esiti = nuove_fatturate + gia_fatturate + presenti_senza_fattura + extra_t1 + scartate_altro
                if somma_esiti == len(df_t1):
                    st.success(f"✔ Quadratura OK: {somma_esiti} esiti su {len(df_t1)} righe del file T1.")
                else:
                    st.warning(f"⚠ Quadratura NON OK: {somma_esiti} esiti su {len(df_t1)} righe del file T1 (possibili numeri verifica duplicati nel file).")
        else:
            st.error("Il file deve contenere 'Numero verifica', 'Fattura' e 'Data pianificata'.")

# ---------------------------------------------------------
# TAB 3: REPORT SINTETICO E TEMPORALE PER SEZIONALE
# ---------------------------------------------------------
with tabs[2]:
    st.header("3. Report Sintetico & Tempi di Conversione per Reparto")
    conn = get_db_connection()
    df_db = pd.read_sql_query("SELECT * FROM verifiche", conn)
    df_extra = pd.read_sql_query("SELECT * FROM verifiche_extra", conn)
    conn.close()

    if not df_db.empty:
        df_db['data_pianificata_dt'] = pd.to_datetime(df_db['data_pianificata'], errors='coerce')
        valid_dates = df_db['data_pianificata_dt'].dropna()

        if not valid_dates.empty:
            min_date = valid_dates.min().date()
            max_date = valid_dates.max().date()
        else:
            min_date = datetime.today().date()
            max_date = datetime.today().date()

        st.markdown("### 📅 Filtro Periodo Pianificazione (Globale)")
        c1, c2 = st.columns([1, 2])

        filtro_tipo = c1.radio("Scegli l'ampiezza dell'analisi:", ["Tutto il database", "Seleziona Range Personalizzato"])

        df_extra_filtrato = df_extra.copy()

        if filtro_tipo == "Tutto il database":
            df_filtrato = df_db.copy()
            st.info(f"Stai analizzando l'intero storico: dal **{min_date.strftime('%d/%m/%Y')}** al **{max_date.strftime('%d/%m/%Y')}**")
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
                    mask_extra = (extra_dt >= start_date) & (extra_dt <= end_date)
                    df_extra_filtrato = df_extra.loc[mask_extra]

                st.info(f"Verifiche programmate dal **{start_date.strftime('%d/%m/%Y')}** al **{end_date.strftime('%d/%m/%Y')}**")
            else:
                st.warning("Seleziona anche la data di fine dal calendario per visualizzare il report.")
                st.stop()

        # Funzione di supporto per mostrare le metriche di un singolo reparto
        def mostra_report_reparto(df_reparto, titolo_reparto):
            st.markdown("---")
            st.subheader(titolo_reparto)

            # Base di partenza: tutto ciò che era in T0 e cosa ne è stato
            tot_prog = len(df_reparto)
            tot_fat = len(df_reparto[df_reparto['stato'] == 'Fatturata'])
            tot_att = len(df_reparto[df_reparto['stato'] == 'In Attesa'])
            tot_non_pres = len(df_reparto[df_reparto['stato'] == 'Non Più Presente'])

            def pct(n):
                return f"{(n / tot_prog * 100):.1f}% del T0" if tot_prog > 0 else "0%"

            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Totale Programmate (T0)", tot_prog)
            m2.metric("Diventate Fattura", tot_fat, pct(tot_fat), delta_color="off")
            m3.metric("In Attesa di Fattura", tot_att, pct(tot_att), delta_color="off")
            m4.metric("Non Più Presenti", tot_non_pres, pct(tot_non_pres), delta_color="off")

            if tot_fat + tot_att + tot_non_pres == tot_prog:
                st.caption(f"✔ Quadratura OK: {tot_fat} + {tot_att} + {tot_non_pres} = {tot_prog}")
            else:
                st.warning(f"⚠ Quadratura NON OK: {tot_fat} + {tot_att} + {tot_non_pres} = {tot_fat + tot_att + tot_non_pres} invece di {tot_prog}")

            df_fatturate = df_reparto[df_reparto['stato'] == 'Fatturata']
            t1, t2, t3, t4 = st.columns(4)
            e_1_sett = len(df_fatturate[df_fatturate['fascia_tempo'] == '1. Entro 1 Settimana'])
            e_2_sett = len(df_fatturate[df_fatturate['fascia_tempo'] == '2. Entro 2 Settimane'])
            e_1_mese = len(df_fatturate[df_fatturate['fascia_tempo'] == '3. Entro 1 Mese'])
            oltre_mese = len(df_fatturate[df_fatturate['fascia_tempo'] == '4. Oltre 1 Mese'])

            t1.metric("≤ 7 Giorni", e_1_sett, f"{(e_1_sett / tot_fat * 100):.1f}% del fatturato" if tot_fat > 0 else "0%", delta_color="off")
            t2.metric("8 - 14 Giorni", e_2_sett, f"{(e_2_sett / tot_fat * 100):.1f}% del fatturato" if tot_fat > 0 else "0%", delta_color="off")
            t3.metric("15 - 30 Giorni", e_1_mese, f"{(e_1_mese / tot_fat * 100):.1f}% del fatturato" if tot_fat > 0 else "0%", delta_color="off")
            t4.metric("> 30 Giorni", oltre_mese, f"{(oltre_mese / tot_fat * 100):.1f}% del fatturato" if tot_fat > 0 else "0%", delta_color="off")

            somma_fasce = e_1_sett + e_2_sett + e_1_mese + oltre_mese
            if somma_fasce == tot_fat:
                st.caption(f"✔ Quadratura fasce OK: {somma_fasce} = {tot_fat} fatturate")
            else:
                st.warning(f"⚠ Quadratura fasce NON OK: somma fasce {somma_fasce} invece di {tot_fat} fatturate")

        # 1. Reparto Ascensori (/A)
        df_ascensori = df_filtrato[df_filtrato['tipo_impianto'] == TIPO_A]
        mostra_report_reparto(df_ascensori, "🛗 Reparto Ascensori (Sezionale /A)")

        # 2. Reparto Messa a Terra (/E)
        df_messaaterra = df_filtrato[df_filtrato['tipo_impianto'] == TIPO_E]
        mostra_report_reparto(df_messaaterra, "⚡ Reparto Messa a Terra (Sezionale /E)")

        # Controllo globale: i due reparti devono coprire tutto il filtrato
        if len(df_ascensori) + len(df_messaaterra) == len(df_filtrato):
            st.caption(f"✔ Controllo globale OK: {len(df_ascensori)} + {len(df_messaaterra)} = {len(df_filtrato)} verifiche")
        else:
            st.warning(f"⚠ Controllo globale NON OK: {len(df_ascensori)} + {len(df_messaaterra)} ≠ {len(df_filtrato)}. Nel database ci sono verifiche con protocollo diverso da /A e /E.")

        # 3. Verifiche comparse nel T1 ma non programmate nel T0
        st.markdown("---")
        st.subheader("🆕 Verifiche presenti nel T1 ma non programmate nel T0")
        if df_extra_filtrato.empty:
            st.success("Nessuna verifica fuori programmazione.")
        else:
            n_extra = len(df_extra_filtrato)
            n_extra_fat = int(df_extra_filtrato['fattura'].notna().sum())
            n_extra_a = int((df_extra_filtrato['tipo_impianto'] == TIPO_A).sum())
            n_extra_e = int((df_extra_filtrato['tipo_impianto'] == TIPO_E).sum())
            n_extra_nd = n_extra - n_extra_a - n_extra_e

            st.warning(f"Rilevate **{n_extra}** verifiche non presenti nelle programmate (T0).")
            x1, x2, x3, x4 = st.columns(4)
            x1.metric("Totale fuori programmazione", n_extra)
            x2.metric("🛗 Ascensori (/A)", n_extra_a)
            x3.metric("⚡ Messa a Terra (/E)", n_extra_e)
            x4.metric("Già con fattura", n_extra_fat)
            if n_extra_nd > 0:
                st.caption(f"{n_extra_nd} verifiche senza reparto assegnabile (colonna 'Protocollo' assente nel T1).")

            st.dataframe(
                df_extra_filtrato[['numero_verifica', 'protocollo', 'tipo_impianto', 'data_pianificata', 'fattura', 'data_t1']],
                use_container_width=True
            )

        st.markdown("---")
        st.subheader("📋 Dettaglio Completo Verifiche Filtrate")
        df_display = df_filtrato[['numero_verifica', 'data_pianificata', 'tipo_impianto', 'codice_impianto', 'fattura', 'stato', 'data_t0', 'data_t1', 'giorni_trascorsi', 'fascia_tempo', 'importo']]
        st.dataframe(df_display, use_container_width=True)

    else:
        st.info("Nessun dato salvato nel sistema. Effettua un primo caricamento nel Tab 1.")
