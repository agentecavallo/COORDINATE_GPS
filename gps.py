import pandas as pd
from geopy.geocoders import Nominatim
from geopy.extra.rate_limiter import RateLimiter
import time

def geocodifica_lista(file_input, file_output):
    # 1. Carichiamo il file Excel
    df = pd.read_excel(file_input)
    
    # Inizializziamo il geolocalizzatore
    geolocator = Nominatim(user_agent="michele_batch_geocoder")
    
    # 2. RateLimiter serve per non essere bloccati dal server:
    # aggiunge automaticamente un ritardo di 1 secondo tra una richiesta e l'altra
    geocode = RateLimiter(geolocator.geocode, min_delay_seconds=1)

    print(f"Inizio elaborazione di {len(df)} indirizzi...")

    # 3. Creiamo la stringa dell'indirizzo completo unendo le colonne
    # Modifica i nomi tra parentesi quadre se le tue colonne si chiamano diversamente
    df['indirizzo_full'] = (df['Indirizzo'].astype(str) + " " + 
                            df['Civico'].astype(str) + ", " + 
                            df['CAP'].astype(str) + " " + 
                            df['Citta'].astype(str) + " (" + 
                            df['Provincia'].astype(str) + "), Italy")

    # 4. Applichiamo la geocodifica
    # Questa riga crea una colonna 'location' con tutti i dati restituiti
    df['location'] = df['indirizzo_full'].apply(geocode)

    # 5. Estraiamo Latitudine e Longitudine dalla colonna 'location'
    df['Latitudine'] = df['location'].apply(lambda loc: loc.latitude if loc else None)
    df['Longitudine'] = df['location'].apply(lambda loc: loc.longitude if loc else None)

    # Rimuoviamo la colonna di supporto 'location' per pulizia
    df.drop(columns=['location', 'indirizzo_full'], inplace=True)

    # 6. Salviamo il risultato in un nuovo file Excel
    df.to_excel(file_output, index=False)
    print(f"Elaborazione completata! File salvato come: {file_output}")

if __name__ == "__main__":
    # Assicurati che il file 'clienti.xlsx' sia nella stessa cartella del programma
    try:
        geocodifica_lista("clienti.xlsx", "clienti_coordinati.xlsx")
    except Exception as e:
        print(f"Errore: {e}")
