import streamlit as st
import pandas as pd
from geopy.geocoders import Nominatim
from geopy.extra.rate_limiter import RateLimiter
import time

# Configurazione pagina Streamlit
st.set_page_config(page_title="Geocodificatore Michele", layout="centered")

st.title("🌍 Geocodificatore Indirizzi")
st.subheader("Strumento per Base Protection - Centro Italia")

st.info("""
Carica un file Excel con una colonna contenente gli indirizzi completi. 
Il sistema aggiungerà automaticamente Latitudine e Longitudine.
""")

# 1. Caricamento file tramite interfaccia
uploaded_file = st.file_uploader("Scegli il tuo file Excel (.xlsx)", type=['xlsx'])

if uploaded_file is not None:
    # Lettura anteprima del file
    df = pd.read_excel(uploaded_file)
    st.write("Anteprima dei dati caricati:")
    st.dataframe(df.head())

    # Selezione della colonna che contiene l'indirizzo
    colonne = df.columns.tolist()
    colonna_selezionata = st.selectbox("Seleziona la colonna che contiene l'indirizzo:", colonne)

    if st.button("Avvia Geocodifica Massiva"):
        # Inizializzazione geolocalizzatore
        geolocator = Nominatim(user_agent="michele_base_protection_app")
        # RateLimiter per evitare blocchi (1.5 secondi tra ogni richiesta)
        geocode = RateLimiter(geolocator.geocode, min_delay_seconds=1.5, error_wait_seconds=5.0)

        progresso = st.progress(0)
        status_text = st.empty()
        
        num_righe = len(df)
        risultati = []

        # Ciclo di geocodifica
        for i, row in df.iterrows():
            indirizzo = str(row[colonna_selezionata]) + ", Italy"
            status_text.text(f"Elaborazione riga {i+1} di {num_righe}: {indirizzo}")
            
            try:
                location = geocode(indirizzo)
                if location:
                    risultati.append({'Latitudine': location.latitude, 'Longitudine': location.longitude})
                else:
                    risultati.append({'Latitudine': None, 'Longitudine': None})
            except Exception as e:
                risultati.append({'Latitudine': None, 'Longitudine': None})
            
            # Aggiornamento barra progresso
            progresso.progress((i + 1) / num_righe)

        # Unione risultati al dataframe originale
        df_risultati = pd.concat([df, pd.DataFrame(risultati)], axis=1)

        st.success("✅ Elaborazione completata!")
        st.dataframe(df_risultati.head())

        # Bottone per scaricare il nuovo file Excel
        from io import BytesIO
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_risultati.to_excel(writer, index=False)
        
        st.download_button(
            label="📥 Scarica Excel con Coordinate",
            data=output.getvalue(),
            file_name="clienti_con_coordinate.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.warning("In attesa del caricamento del file Excel...")

# Nota a fondo pagina
st.markdown("---")
st.caption("Nota: Il processo richiede circa 1.5 secondi per ogni indirizzo per rispettare le policy gratuite di OpenStreetMap.")
