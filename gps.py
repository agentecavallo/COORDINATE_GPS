import streamlit as st
import pandas as pd
from geopy.geocoders import Nominatim
from geopy.extra.rate_limiter import RateLimiter
from io import BytesIO
import re

st.set_page_config(page_title="Geocodificatore Michele", layout="centered")
st.title("🌍 Geocodificatore Base Protection")

uploaded_file = st.file_uploader("Scegli il tuo file Excel", type=['xlsx'])

def pulisci_indirizzo(testo):
    """Funzione per pulire l'indirizzo da caratteri che disturbano la ricerca"""
    if pd.isna(testo): return ""
    # Sostituisce "C. " con niente se non trova nulla, o prova a rimuovere punti e virgole
    t = str(testo).replace(",", " ").replace(".", " ")
    # Rimuove spazi doppi
    t = " ".join(t.split())
    return t

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file, dtype=str) 
    st.write("Dati caricati correttamente.")

    colonne = df.columns.tolist()
    colonna_selezionata = st.selectbox("Seleziona la colonna Indirizzo:", colonne)

    if st.button("Avvia Elaborazione"):
        geolocator = Nominatim(user_agent="michele_base_ultra_fix")
        geocode = RateLimiter(geolocator.geocode, min_delay_seconds=1.2, error_wait_seconds=5.0)

        progresso = st.progress(0)
        num_righe = len(df)
        risultati = []

        for i, row in df.iterrows():
            original_addr = str(row[colonna_selezionata])
            
            # TENTATIVO 1: Indirizzo originale + Italy
            query1 = f"{original_addr}, Italy"
            location = geocode(query1)

            # TENTATIVO 2: Se fallisce, puliamo l'indirizzo (via i punti e virgole)
            if not location:
                query2 = f"{pulisci_indirizzo(original_addr)}, Italy"
                location = geocode(query2)

            # TENTATIVO 3: Se fallisce ancora, proviamo a togliere il civico e tenere solo Via e Città
            if not location:
                # Esempio: "VIA C BESCHI 8 00125 ROMA" -> prende solo la parte testuale
                # Questo è un tentativo disperato per non lasciarti il "NON TROVATO"
                parti = query2.split()
                if len(parti) > 3:
                    query3 = " ".join(parti[:2]) + " " + " ".join(parti[-2:])
                    location = geocode(query3)

            if location:
                risultati.append({'Latitudine': location.latitude, 'Longitudine': location.longitude})
            else:
                risultati.append({'Latitudine': "NON TROVATO", 'Longitudine': "NON TROVATO"})
            
            progresso.progress((i + 1) / num_righe)

        df['Latitudine'] = [r['Latitudine'] for r in risultati]
        df['Longitudine'] = [r['Longitudine'] for r in risultati]

        st.success("✅ Elaborazione terminata!")
        
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)
        
        st.download_button(
            label="📥 Scarica Excel Ottimizzato",
            data=output.getvalue(),
            file_name="clienti_coordinate_fix.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
