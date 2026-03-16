import pandas as pd
import numpy as np
import random
import os
from datetime import datetime, timedelta

def main():
    # Setup
    output_dir = "Data_Kadea_RDC"
    os.makedirs(output_dir, exist_ok=True)

    np.random.seed(42)
    random.seed(42)

    # 26 Provinces of the DRC
    provinces_rdc = [
        "Kinshasa", "Kongo-Central", "Kwango", "Kwilu", "Mai-Ndombe",
        "Kasai", "Kasai-Central", "Kasai-Oriental", "Lomami", "Sankuru",
        "Maniema", "Sud-Kivu", "Nord-Kivu", "Ituri", "Haut-Uele",
        "Bas-Uele", "Tshopo", "Equateur", "Nord-Ubangi", "Sud-Ubangi",
        "Mongala", "Tshuapa", "Tanganyika", "Haut-Lomami", "Lualaba", "Haut-Katanga"
    ]

    technologies = ['3G', '4G', '5G']
    equipements = ['Nokia', 'Huawei', 'Ericsson', 'ZTE']
    forfaits = ['Bloqué 5Go', 'Illimité 50Go', 'Premium 5G 150Go', 'Pro Monde']

    # Pre-generate global IDs to ensure VLOOKUP works across fragmented files
    global_cell_ids = [f"ANT-{i:04d}" for i in range(1, 1001)] # 1000 antennas
    global_client_ids = [f"CLI-{i:05d}" for i in range(1, 2001)] # 2000 clients

    # --- 1. Generate 25 Logs_Reseau files (Max 50 rows each) ---
    start_date = datetime(2023, 10, 1)

    for file_idx in range(1, 26):
        num_rows = random.randint(30, 50)

        dates = [start_date + timedelta(days=random.randint(0, 30), hours=random.randint(0, 23), minutes=random.randint(0, 59)) for _ in range(num_rows)]
        durations = [random.randint(5, 120) for _ in range(num_rows)]

        data = {
            'Date_Heure_Panne': dates,
            'Cell_ID': [random.choice(global_cell_ids) for _ in range(num_rows)],
            'Duree_Panne_Min': durations,
            'Statut_Resolution': [random.choice(['Résolu', 'En Cours', 'Non Résolu']) for _ in range(num_rows)]
        }
        df = pd.DataFrame(data)

        # Intentional Errors
        if random.random() < 0.3: # 30% chance to have a duplicate row
            df = pd.concat([df, df.sample(1)]).reset_index(drop=True)

        if random.random() < 0.3: # 30% chance for missing Cell_ID
            missing_idx = random.randint(0, len(df)-1)
            df.loc[missing_idx, 'Cell_ID'] = np.nan

        if random.random() < 0.2: # 20% chance for lower case ID
            lower_idx = random.randint(0, len(df)-1)
            if pd.notna(df.loc[lower_idx, 'Cell_ID']):
                df.loc[lower_idx, 'Cell_ID'] = df.loc[lower_idx, 'Cell_ID'].lower()

        if random.random() < 0.1: # 10% chance for string date format
            df['Date_Heure_Panne'] = df['Date_Heure_Panne'].astype(object)
            date_idx = random.randint(0, len(df)-1)
            df.loc[date_idx, 'Date_Heure_Panne'] = df.loc[date_idx, 'Date_Heure_Panne'].strftime('%d/%m/%Y %H:%M')

        filename = f"{output_dir}/Logs_Reseau_Part_{file_idx:02d}.xlsx"
        df.to_excel(filename, index=False)


    # --- 2. Generate 10 Registre_Maintenance files (Max 50 rows each) ---
    # We distribute the 1000 global antennas across 10 files (we'll limit to 50 max to match the prompt)
    # This means not all antennas will exist in the registry, naturally causing #N/A !

    used_cells = random.sample(global_cell_ids, 500) # Only 500 antennas have registry data
    cell_chunks = [used_cells[i:i + 50] for i in range(0, 500, 50)] # Split into chunks of 50

    for file_idx in range(1, 11):
        chunk = cell_chunks[file_idx - 1]
        num_rows = len(chunk)

        data = {
            'ID_Antenne': chunk,
            'Province_RDC': [random.choice(provinces_rdc) for _ in range(num_rows)],
            'Latitude': [round(random.uniform(-13.0, 5.0), 6) for _ in range(num_rows)], # RDC Approx latitudes
            'Longitude': [round(random.uniform(12.0, 31.0), 6) for _ in range(num_rows)], # RDC Approx longitudes
            'Technologie': [random.choice(technologies) for _ in range(num_rows)],
            'Equipementier': [random.choice(equipements) for _ in range(num_rows)]
        }
        df = pd.DataFrame(data)

        # Intentional Errors
        if random.random() < 0.3: # 30% chance for missing province
            missing_idx = random.randint(0, len(df)-1)
            df.loc[missing_idx, 'Province_RDC'] = np.nan

        if random.random() < 0.4: # 40% chance for space in ID (breaks VLOOKUP)
            space_idx = random.randint(0, len(df)-1)
            df.loc[space_idx, 'ID_Antenne'] = df.loc[space_idx, 'ID_Antenne'] + " "

        filename = f"{output_dir}/Registre_Maintenance_Part_{file_idx:02d}.xlsx"
        df.to_excel(filename, index=False)


    # --- 3. Generate 15 Clients_Churn files (Max 50 rows each) ---
    used_clients = random.sample(global_client_ids, 750)
    client_chunks = [used_clients[i:i + 50] for i in range(0, 750, 50)] # Split into chunks of 50

    for file_idx in range(1, 16):
        chunk = client_chunks[file_idx - 1]
        num_rows = len(chunk)

        data = {
            'ID_Client': chunk,
            'Province_Client': [random.choice(provinces_rdc) for _ in range(num_rows)],
            'Anciennete_Mois': [random.randint(1, 120) for _ in range(num_rows)],
            'Type_Forfait': [random.choice(forfaits) for _ in range(num_rows)],
            'Duree_Cumulee_Panne_Min': [random.randint(0, 500) for _ in range(num_rows)],
            'Facture_Mensuelle': [round(random.uniform(10.0, 100.0), 2) for _ in range(num_rows)],
            'Plaintes_Service_Client': [random.randint(0, 5) for _ in range(num_rows)]
        }
        df = pd.DataFrame(data)

        # Intentional Errors
        if random.random() < 0.2: # Negative ancienty
            neg_idx = random.randint(0, len(df)-1)
            df.loc[neg_idx, 'Anciennete_Mois'] = -df.loc[neg_idx, 'Anciennete_Mois']

        if random.random() < 0.3: # Text in numeric column
            df['Facture_Mensuelle'] = df['Facture_Mensuelle'].astype(object)
            text_idx = random.randint(0, len(df)-1)
            df.loc[text_idx, 'Facture_Mensuelle'] = str(df.loc[text_idx, 'Facture_Mensuelle']).replace('.', ',') + "€"

        filename = f"{output_dir}/Clients_Churn_Part_{file_idx:02d}.xlsx"
        df.to_excel(filename, index=False)

    print(f"Successfully generated 50 fragmented Excel files in {output_dir}/")

if __name__ == "__main__":
    main()
