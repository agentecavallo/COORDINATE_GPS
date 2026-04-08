import streamlit as st
import pandas as pd
from geopy.geocoders import Nominatim
from geopy.extra.rate_limiter import RateLimiter
from io import BytesIO

st.set_page_config(page_title="Geocodificatore Michele", layout="centered")
st.title("🌍 Geocodificatore Base Protection (Versione Punto)")

uploaded_file = st.file_uploader("Scegli il tuo file Excel", type=['xlsx'])

if uploaded_file is not None:
    # Leggiamo tutto come stringa per non perdere gli zeri dei telefoni
    df = pd.read_excel(uploaded_file, dtype=str) 
    
    colonne = df.columns.tolist()
    colonna_selezionata = st.selectbox("Seleziona la colonna Indirizzo:", colonne)

    if st.button("Avvia Elaborazione"):
        geolocator = Nominatim(user_agent="michele_base_dot_fix")
        geocode = RateLimiter(geolocator.geocode, min_delay_seconds=1.2)

        progresso = st.progress(0)
        num_righe = len(df)
        risultati_lat = []
        risultati_lon = []

        for i, row in df.iterrows():
            indirizzo = str(row[colonna_selezionata]) + ", Italy"
            try:
                location = geocode(indirizzo)
                if location:
                    # TRUCCO: Trasformiamo in stringa e forziamo il punto
                    lat = str(location.latitude).replace(',', '.')
                    lon = str(location.longitude).replace(',', '.')
                    risultati_lat.append(lat)
                    risultati_lon.append(lon)
                else:
                    risultati_lat.append("NON TROVATO")
                    risultati_lon.append("NON TROVATO")
            except:
                risultati_lat.append("ERRORE")
                risultati_lon.append("ERRORE")
            
            progresso.progress((i + 1) / num_righe)

        # Aggiungiamo le colonne al dataframe
        df['Latitudine'] = risultati_lat
        df['Longitudine'] = risultati_lon

        st.success("✅ Completato! Ora i punti GPS hanno il punto decimale.")
        
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)
        
        st.download_button(
            label="📥 Scarica Excel con PUNTI",
            data=output.getvalue(),
            file_name="clienti_base_punto.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
