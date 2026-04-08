import streamlit as st
import pandas as pd
from geopy.geocoders import Nominatim
from geopy.extra.rate_limiter import RateLimiter
from io import BytesIO
import re

st.set_page_config(page_title="Geocodificatore Michele", layout="centered")
st.title("🌍 Geocodificatore Base Protection - Versione Ultra")

def pulizia_chirurgica(testo):
    if not testo or pd.isna(testo): return ""
    t = str(testo).upper()
    
    # 1. Gestione C/O (Presso) - Elimina il nome azienda sia se prima che dopo
    # Se "VIA ROMA 1 C/O PINCO SRL" -> resta "VIA ROMA 1"
    # Se "C/O PINCO SRL VIA ROMA 1" -> resta "VIA ROMA 1"
    if "C/O" in t:
        parti = t.split("C/O")
        # Cerchiamo quale parte contiene "VIA", "CORSO", "PIAZZA", ecc.
        t = parti[0] if any(x in parti[0] for x in ["VIA", "PIAZZA", "CORSO", "VIALE", "STRADA"]) else parti[-1]

    # 2. Rimuove tutto ciò che è tra parentesi (es. nomi province o note)
    t = re.sub(r'\(.*?\)', '', t)
    
    # 3. Pulizia sigle fastidiose
    disturbatori = [
        "VIA LOC.", "LOC.", "LOCALITÀ", "LOCALITA'", "SNC", "P.I.P.", "Z.I.", 
        "NR.", " N.", "ACCESSO CIVICO", "ZONA INDUSTRIALE", "KM", "STRADA STATALE"
    ]
    for p in disturbatori:
        t = t.replace(p, " ")

    # 4. Semplificazione drastica del civico (es. 8/10 o 863/A diventano 8 o 863)
    t = re.sub(r'(\d+)/[A-Z0-9/]+', r'\1', t)
    
    # 5. Espansione e pulizia finale
    t = t.replace("F.LLI", "FRATELLI").replace("C.DA", "CONTRADA").replace("P.ZZA", "PIAZZA")
    t = t.replace("V.LE", "VIALE").replace("C.SO", "CORSO").replace("C.", " ")
    
    # Rimuove punteggiatura e spazi doppi
    t = re.sub(r'[.,]', ' ', t)
    return " ".join(t.split())

if uploaded_file := st.file_uploader("Carica il file Excel", type=['xlsx']):
    df = pd.read_excel(uploaded_file, dtype=str)
    col = st.selectbox("Seleziona la colonna Indirizzo:", df.columns.tolist())

    if st.button("Avvia Geocodifica Totale"):
        geolocator = Nominatim(user_agent="michele_base_protection_v4")
        geocode = RateLimiter(geolocator.geocode, min_delay_seconds=1.2, timeout=10)

        progresso = st.progress(0)
        num_righe = len(df)
        risultati = []

        for i, row in df.iterrows():
            addr_orig = str(row[col])
            addr_clean = pulizia_chirurgica(addr_orig)
            location = None

            # --- STRATEGIA A 4 LIVELLI ---
            # 1. Prova l'indirizzo pulito completo
            location = geocode(f"{addr_clean}, Italy")

            # 2. Se fallisce, prova a togliere il numero civico (cerca solo la via)
            if not location:
                # Prende tutto tranne i numeri finali
                via_solo = re.sub(r'\d+', '', addr_clean).strip()
                # Prende l'ultima parola (che di solito è la città)
                citta_solo = addr_clean.split()[-1]
                location = geocode(f"{via_solo} {citta_solo}, Italy")

            # 3. Se fallisce, prova solo "Nome Via + Città"
            if not location:
                parti = addr_clean.split()
                if len(parti) >= 2:
                    location = geocode(f"{parti[0]} {parti[1]} {parti[-1]}, Italy")

            # 4. Ultima spiaggia: Cerca solo la Città (per non avere il buco nel file)
            if not location:
                citta_fallback = addr_clean.split()[-1]
                location = geocode(f"{citta_fallback}, Italy")

            if location:
                risultati.append({
                    'Latitudine': str(location.latitude).replace(',', '.'),
                    'Longitudine': str(location.longitude).replace(',', '.')
                })
            else:
                risultati.append({'Latitudine': "NON TROVATO", 'Longitudine': "NON TROVATO"})
            
            progresso.progress((i + 1) / num_righe)

        df['Latitudine'], df['Longitudine'] = [r['Latitudine'] for r in risultati], [r['Longitudine'] for r in risultati]
        
        st.success("✅ Finito! Ora abbiamo usato anche il fallback sulla città se la via era impossibile.")
        
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)
        st.download_button(label="📥 Scarica Risultati", data=output.getvalue(), file_name="mappa_clienti_base_v4.xlsx")
