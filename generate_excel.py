import pandas as pd
import numpy as np
import random
from datetime import datetime, timedelta

def generate_logs_reseau(num_rows=850):
    np.random.seed(42)
    random.seed(42)

    cell_ids = [f"ANT-{i:04d}" for i in range(1, 201)]

    # Generate random dates over the last month
    start_date = datetime(2023, 10, 1)
    dates = [start_date + timedelta(days=random.randint(0, 30), hours=random.randint(0, 23), minutes=random.randint(0, 59)) for _ in range(num_rows)]

    durations = [random.randint(5, 120) for _ in range(num_rows)]

    data = {
        'Date_Heure_Panne': dates,
        'Cell_ID': [random.choice(cell_ids) for _ in range(num_rows)],
        'Duree_Panne_Min': durations,
        'Statut_Resolution': [random.choice(['Résolu', 'En Cours', 'Non Résolu']) for _ in range(num_rows)]
    }

    df = pd.DataFrame(data)

    # Introduce intentional errors
    # 1. Duplicates
    df = pd.concat([df, df.sample(15)]).reset_index(drop=True)

    # 2. Missing values in Cell_ID
    missing_indices = random.sample(range(len(df)), 10)
    for idx in missing_indices:
        df.loc[idx, 'Cell_ID'] = np.nan

    # 3. Mixed formats in Cell_ID (some lower case)
    mixed_indices = random.sample(range(len(df)), 20)
    for idx in mixed_indices:
        if pd.notna(df.loc[idx, 'Cell_ID']):
            df.loc[idx, 'Cell_ID'] = df.loc[idx, 'Cell_ID'].lower()

    # 4. Inconsistent dates (string formats instead of datetime for a few)
    date_err_indices = random.sample(range(len(df)), 5)
    df['Date_Heure_Panne'] = df['Date_Heure_Panne'].astype(object)
    for idx in date_err_indices:
        df.loc[idx, 'Date_Heure_Panne'] = df.loc[idx, 'Date_Heure_Panne'].strftime('%d/%m/%Y %H:%M')

    # Ensure Date_Heure_Panne column is converted safely for Excel (handle mixed types)
    # We will let to_excel handle it, but it's better to export as strings if mixed, or keep datetime and let openpyxl fail or succeed.
    # To be safe and simulate bad data, let's cast the whole column to string to simulate mixed formats that need parsing

    df.to_excel('Logs_Reseau.xlsx', index=False)
    print(f"Logs_Reseau.xlsx generated with {len(df)} rows.")

def generate_registre_maintenance(num_rows=250):
    np.random.seed(42)
    random.seed(42)

    cell_ids = [f"ANT-{i:04d}" for i in range(1, 251)]
    regions = ['Ile-de-France', 'PACA', 'Auvergne-Rhone-Alpes', 'Occitanie', 'Nouvelle-Aquitaine', 'Bretagne']
    technologies = ['3G', '4G', '5G']
    equipements = ['Nokia', 'Huawei', 'Ericsson']

    data = {
        'ID_Antenne': cell_ids,
        'Region': [random.choice(regions) for _ in range(num_rows)],
        'Latitude': [round(random.uniform(41.0, 51.0), 6) for _ in range(num_rows)],
        'Longitude': [round(random.uniform(-5.0, 8.0), 6) for _ in range(num_rows)],
        'Technologie': [random.choice(technologies) for _ in range(num_rows)],
        'Equipementier': [random.choice(equipements) for _ in range(num_rows)]
    }

    df = pd.DataFrame(data)

    # Intentional errors
    # 1. Missing regions
    missing_indices = random.sample(range(len(df)), 15)
    for idx in missing_indices:
        df.loc[idx, 'Region'] = np.nan

    # 2. Spaces in IDs to mess up VLOOKUP
    space_indices = random.sample(range(len(df)), 10)
    for idx in space_indices:
        df.loc[idx, 'ID_Antenne'] = df.loc[idx, 'ID_Antenne'] + " "

    df.to_excel('Registre_Maintenance.xlsx', index=False)
    print(f"Registre_Maintenance.xlsx generated with {len(df)} rows.")

def generate_clients_churn(num_rows=1000):
    np.random.seed(42)
    random.seed(42)

    client_ids = [f"CLI-{i:05d}" for i in range(1, num_rows + 1)]
    forfaits = ['Bloqué 5Go', 'Illimité 50Go', 'Premium 5G 150Go', 'Pro Monde']

    data = {
        'ID_Client': client_ids,
        'Anciennete_Mois': [random.randint(1, 120) for _ in range(num_rows)],
        'Type_Forfait': [random.choice(forfaits) for _ in range(num_rows)],
        'Duree_Cumulee_Panne_Min': [random.randint(0, 500) for _ in range(num_rows)],
        'Facture_Mensuelle': [round(random.uniform(10.0, 100.0), 2) for _ in range(num_rows)],
        'Plaintes_Service_Client': [random.randint(0, 5) for _ in range(num_rows)]
    }

    df = pd.DataFrame(data)

    # Intentional errors
    # Negative values in ancienty
    neg_indices = random.sample(range(len(df)), 5)
    for idx in neg_indices:
        df.loc[idx, 'Anciennete_Mois'] = -df.loc[idx, 'Anciennete_Mois']

    # Text in numeric column
    text_indices = random.sample(range(len(df)), 5)
    df['Facture_Mensuelle'] = df['Facture_Mensuelle'].astype(object)
    for idx in text_indices:
        df.loc[idx, 'Facture_Mensuelle'] = str(df.loc[idx, 'Facture_Mensuelle']).replace('.', ',') + "€"

    df.to_excel('Clients_Churn.xlsx', index=False)
    print(f"Clients_Churn.xlsx generated with {len(df)} rows.")

if __name__ == "__main__":
    generate_logs_reseau()
    generate_registre_maintenance()
    generate_clients_churn()
