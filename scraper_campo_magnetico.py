import glob
import ast
import requests
import pandas as pd
import matplotlib.pyplot as plt
#con questo script si creano file .xlsx in cui vengono aggiunti alle colonne relative a longitudine, latitudine e tempo le colonne contenenti i dati estrapolati dal sito ngdc.noa.gov relativi al campo magnetico terrestre che ci interessano. Vengono poi realizzati i grafici che mettono in relazione le coordinate con i valori di telta rispetto a partenza e/o arrivo.

info = ["inclination", "totalintensity", "horintensity"]

for file_path in glob.glob('*.xlsx'):
    print('Apro file', file_path)
    df = pd.read_excel(file_path)
    
    # Inizializzare colonne per i dati aggiuntivi
    for infoi in info:
        df[infoi] = pd.Series()
    
    # Richiesta dei dati per ogni riga
    for i in range(len(df)):
        response = requests.get("https://www.ngdc.noaa.gov/geomag-web/calculators/calculateIgrfwmm?key=EAU2y&lat1=" + str(df['lat'][i]) 
        + "&lon1=" + str(df['lon'][i]) 
        + "&model=IGRF&startYear=" + str(df['Year'][i]) 
        + "&endYear=" + str(df['Year'][i]) 
        + f"&resultFormat=json&startMonth={df['Month'][i]:02d}&startDay={df['Day'][i]}&endMonth={df['Month'][i]}&endDay={df['Day'][i]}")
        
        rc = response.content
        response_dict = ast.literal_eval(rc.decode('utf-8'))["result"][0]
        
        for infoi in info:
            df.loc[i, infoi] = response_dict[infoi]
    
    # Calcolare i delta per le variabili richieste
    for infoi in info:
        dato_iniziale = df.iloc[0][infoi]
        dato_finale = df.iloc[-1][infoi]
        
        delta_iniziale = df[infoi] - dato_iniziale
        delta_finale = dato_finale - df[infoi]
        
        df[f'delta {infoi} iniziale'] = delta_iniziale
        df[f'delta {infoi} finale'] = delta_finale
    
    # Salvare il file modificato
    output_name = file_path.split('.xlsx')[0] + "_MODIFICATO.xlsx"
    df.to_excel(output_name, index=False)
    print("Dati nuovi salvati in", output_name)
    
    # Creare grafici
    for infoi in info:
        plt.figure(figsize=(12, 6))
        
        plt.subplot(2, 2, 1)
        plt.plot(df.index, df[f'delta {infoi} iniziale'], label=f'Delta {infoi} Iniziale')
        plt.xlabel('Indice')
        plt.ylabel(f'Delta {infoi} Iniziale')
        plt.title(f'Delta {infoi} Iniziale')
        plt.legend()
        
        plt.subplot(2, 2, 2)
        plt.plot(df.index, df[f'delta {infoi} finale'], label=f'Delta {infoi} Finale', color='orange')
        plt.xlabel('Indice')
        plt.ylabel(f'Delta {infoi} Finale')
        plt.title(f'Delta {infoi} Finale')
        plt.legend()
        
        plt.tight_layout()
        plt.savefig(f"{file_path.split('.xlsx')[0]}_{infoi}_grafico.png")
        plt.close()
    
    print("Grafici creati e salvati.")
