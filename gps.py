import streamlit as st
import pandas as pd
from geopy.geocoders import Nominatim
from geopy.extra.rate_limiter import RateLimiter
from io import BytesIO

st.set_page_config(page_title="Geocodificatore Michele", layout="centered")
st.title("🌍 Geocodificatore Base Protection - Ultra Fix")

uploaded_file = st.file_uploader("Scegli il tuo file Excel", type=['xlsx'])

def pulizia_profonda(testo):
    """Sostituisce le abbreviazioni comuni che bloccano Nominatim"""
    t = str(testo).upper()
    t = t.replace("F.LLI", "FRATELLI")
    t = t.replace("C.", "") # Spesso l'iniziale del nome blocca tutto, meglio toglierla
    t = t.replace("P.ZZA", "PIAZZA")
    t = t.replace("V.LE", "VIALE")
    t = t.replace(".", " ")
    t = t.replace(",", " ")
    return " ".join(t.split())

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file, dtype=str) 
    colonna_selezionata = st.selectbox("Seleziona la colonna Indirizzo:", df.columns.tolist())

    if st.button("Avvia Elaborazione"):
        geolocator = Nominatim(user_agent="michele_final_boss")
        geocode = RateLimiter(geolocator.geocode, min_delay_seconds=1.2)

        progresso = st.progress(0)
        num_righe = len(df)
        lats, lons = [], []

        for i, row in df.iterrows():
            addr_orig = str(row[colonna_selezionata])
            location = None

            # --- TENTATIVO 1: Originale ---
            location = geocode(f"{addr_orig}, Italy")

            # --- TENTATIVO 2: Pulizia sigle (F.lli -> Fratelli, etc.) ---
            if not location:
                addr_clean = pulizia_profonda(addr_orig)
                location = geocode(f"{addr_clean}, Italy")

            # --- TENTATIVO 3: Solo Via e Città (Rimuoviamo il civico se dà noia) ---
            # Se l'indirizzo è "VIA ROMA 10 00100 ROMA", proviamo "VIA ROMA ROMA"
            if not location:
                parti = addr_clean.split()
                if len(parti) > 4:
                    tentativo_breve = f"{parti[0]} {parti[1]} {parti[-1]}, Italy"
                    location = geocode(tentativo_breve)

            if location:
                lats.append(str(location.latitude).replace(',', '.'))
                lons.append(str(location.longitude).replace(',', '.'))
            else:
                lats.append("NON TROVATO")
                lons.append("NON TROVATO")
            
            progresso.progress((i + 1) / num_righe)

        df['Latitudine'] = lats
        df['Longitudine'] = lons

        st.success("✅ Elaborazione completata con tentativi multipli!")
        
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)
        
        st.download_button(label="📥 Scarica Excel Finale", data=output.getvalue(), file_name="clienti_base_finito.xlsx")
