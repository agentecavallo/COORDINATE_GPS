import streamlit as st
import pandas as pd
from geopy.geocoders import Nominatim
from geopy.extra.rate_limiter import RateLimiter
from io import BytesIO

st.set_page_config(page_title="Geocodificatore Michele", layout="centered")
st.title("🌍 Geocodificatore Professionale")

uploaded_file = st.file_uploader("Scegli il tuo file Excel", type=['xlsx'])

if uploaded_file is not None:
    # FIX TELEFONI: Leggiamo tutto il file come stringhe per non perdere gli zeri
    df = pd.read_excel(uploaded_file, dtype=str) 
    
    st.write("Dati caricati (formattati come testo):")
    st.dataframe(df.head())

    colonne = df.columns.tolist()
    colonna_selezionata = st.selectbox("Seleziona la colonna Indirizzo:", colonne)

    if st.button("Avvia Elaborazione"):
        geolocator = Nominatim(user_agent="michele_base_fix")
        # Aumentiamo un po' il timeout per gli indirizzi difficili
        geocode = RateLimiter(geolocator.geocode, min_delay_seconds=1.5, error_wait_seconds=5.0)

        progresso = st.progress(0)
        num_righe = len(df)
        risultati = []

        for i, row in df.iterrows():
            # PULIZIA: Prendiamo l'indirizzo e togliamo eventuali spazi doppi
            raw_addr = str(row[colonna_selezionata]).strip()
            
            # Proviamo prima l'indirizzo così com'è
            query = f"{raw_addr}, Italy"
            
            try:
                location = geocode(query)
                
                # Se non lo trova, facciamo un secondo tentativo "semplificato"
                # (Esempio: togliamo 'F.LLI' o caratteri strani se necessario)
                if not location and "F.LLI" in query:
                    query_alt = query.replace("F.LLI", "Fratelli")
                    location = geocode(query_alt)

                if location:
                    risultati.append({'Latitudine': location.latitude, 'Longitudine': location.longitude})
                else:
                    risultati.append({'Latitudine': "NON TROVATO", 'Longitudine': "NON TROVATO"})
            except:
                risultati.append({'Latitudine': "ERRORE", 'Longitudine': "ERRORE"})
            
            progresso.progress((i + 1) / num_righe)

        # Uniamo i risultati
        df['Latitudine'] = [r['Latitudine'] for r in risultati]
        df['Longitudine'] = [r['Longitudine'] for r in risultati]

        st.success("✅ Completato!")
        
        # Salvataggio mantenendo il formato testo
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)
        
        st.download_button(
            label="📥 Scarica Excel Corretto",
            data=output.getvalue(),
            file_name="coordinate_clienti_base.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
