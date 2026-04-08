import pandas as pd
from geopy.geocoders import Nominatim
from geopy.extra.rate_limiter import RateLimiter

def geocodifica_colonna_unica(file_input, nome_colonna, file_output):
    # Carichiamo l'Excel
    df = pd.read_excel(file_input)
    
    # Inizializziamo il geolocalizzatore
    geolocator = Nominatim(user_agent="michele_geocoder_pro")
    
    # Aggiungiamo il ritardo per non essere bloccati (1.5 secondi per stare sicuri con 500 record)
    geocode = RateLimiter(geolocator.geocode, min_delay_seconds=1.5, error_wait_seconds=5.0)

    print(f"Elaborazione di {len(df)} indirizzi in corso...")

    # Applichiamo la ricerca direttamente sulla colonna esistente
    # Aggiungiamo ", Italy" alla fine per aiutare il sistema se non è specificato
    df['temp_location'] = df[nome_colonna].apply(lambda x: geocode(str(x) + ", Italy") if pd.notnull(x) else None)

    # Estraiamo i dati
    df['Latitudine'] = df['temp_location'].apply(lambda loc: loc.latitude if loc else None)
    df['Longitudine'] = df['temp_location'].apply(lambda loc: loc.longitude if loc else None)

    # Pulizia: rimuoviamo la colonna temporanea
    df.drop(columns=['temp_location'], inplace=True)

    # Salvataggio
    df.to_excel(file_output, index=False)
    print(f"Fatto! Risultati salvati in: {file_output}")

if __name__ == "__main__":
    # CAMBIA "Indirizzo Completo" con il nome esatto della tua colonna nell'Excel
    geocodifica_colonna_unica("clienti.xlsx", "Indirizzo Completo", "risultati_gps.xlsx")
