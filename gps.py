import streamlit as st
import pandas as pd
from geopy.geocoders import Nominatim
from geopy.extra.rate_limiter import RateLimiter
from io import BytesIO
import re

st.set_page_config(page_title="Geocodificatore Michele", layout="centered")
st.title("🌍 Geocodificatore Base Protection - Master Fix")

def pulizia_totale(testo):
    if not testo: return ""
    t = str(testo).upper()
    
    # 1. Gestione C/O: prendiamo solo quello che c'è DOPO il C/O
    if "C/O" in t:
        t = t.split("C/O")[-1]
    
    # 2. Semplificazione Numeri Civici: trasforma "10/A" o "10/12" in "10"
    # Aiuta tantissimo a trovare la strada se il civico esatto non è mappato
    t = re.sub(r'(\d+)/[A-Z0-9/]+', r'\1', t)
    
    # 3. Rimozione parole inutili per il GPS
    parole_da_eliminare = [
        "VIA LOC.", "LOC.", "LOCALITÀ", "LOCALITA'", "SNC", 
        "P.I.P.", "Z.I.", "NR.", " N. ", "ACCESSO CIVICO", "ZONA INDUSTRIALE"
    ]
    for p in parole_da_eliminare:
        t = t.replace(p, " ")

    # 4. Espansione abbreviazioni
    t = t.replace("F.LLI", "FRATELLI")
    t = t.replace("C.DA", "CONTRADA")
    t = t.replace("P.ZZA", "PIAZZA")
    t = t.replace("V.LE", "VIALE")
    t = t.replace("C.SO", "CORSO")
    t = t.replace("C.", " ") # Rimuove iniziali puntate come C. Beschi
    
    # 5. Pulizia finale punteggiatura e spazi
    t = t.replace(",", " ").replace(".", " ")
    return " ".join(t.split())

uploaded_file = st.file_uploader("Carica il file Excel con i clienti", type=['xlsx'])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file, dtype=str) 
    colonna_selezionata = st.selectbox("Seleziona la colonna Indirizzo:", df.columns.tolist())

    if st.button("Avvia Geocodifica"):
        geolocator = Nominatim(user_agent="michele_base_ultimate")
        # Aumentiamo il timeout per indirizzi complessi
        geocode = RateLimiter(geolocator.geocode, min_delay_seconds=1.2, error_wait_seconds=5.0)

        progresso = st.progress(0)
        num_righe = len(df)
        risultati = []

        for i, row in df.iterrows():
            addr_orig = str(row[colonna_selezionata])
            
            # TENTATIVO 1: Indirizzo pulito "Master"
            addr_clean = pulizia_totale(addr_orig)
            location = geocode(f"{addr_clean}, Italy")

            # TENTATIVO 2 (Fallback): Se fallisce, proviamo solo Via + Città (senza civico)
            if not location:
                parti = addr_clean.split()
                if len(parti) > 3:
                    # Prende la via (primi due elementi) e la città (ultimo elemento)
                    tentativo_corto = f"{parti[0]} {parti[1]} {parti[-1]}, Italy"
                    location = geocode(tentativo_corto)

            if location:
                risultati.append({
                    'Latitudine': str(location.latitude).replace(',', '.'),
                    'Longitudine': str(location.longitude).replace(',', '.')
                })
            else:
                risultati.append({'Latitudine': "NON TROVATO", 'Longitudine': "NON TROVATO"})
            
            progresso.progress((i + 1) / num_righe)

        df['Latitudine'] = [r['Latitudine'] for r in risultati]
        df['Longitudine'] = [r['Longitudine'] for r in risultati]

        st.success("✅ Finito! Ora molti più indirizzi dovrebbero essere mappati.")
        
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)
        
        st.download_button(label="📥 Scarica Excel per Mappa", data=output.getvalue(), file_name="clienti_base_gps.xlsx")
