import pandas as pd
import numpy as np

def create_dirty_dataset():
    # Liste des 26 provinces de la RDC
    provinces = [
        "Kinshasa", "Kongo Central", "Kwango", "Kwilu", "Mai-Ndombe",
        "Kasaï", "Kasaï-Central", "Kasaï-Oriental", "Lomami", "Sankuru",
        "Maniema", "Sud-Kivu", "Nord-Kivu", "Ituri", "Haut-Uele", "Tshopo",
        "Bas-Uele", "Nord-Ubangi", "Mongala", "Sud-Ubangi", "Équateur",
        "Tshuapa", "Tanganyika", "Haut-Lomami", "Lualaba", "Haut-Katanga"
    ]

    # Génération de données factices pour Kadea Telco (Antennes / Cell Towers)
    np.random.seed(42)
    num_records = 500

    data = {
        "Tower_ID": [f"TWR-{np.random.randint(1000, 9999)}" for _ in range(num_records)],
        "Province": np.random.choice(provinces, num_records),
        "Status": np.random.choice(["Active", "Maintenance", "Offline", "active", " OFF-LINE ", ""], num_records),
        "Signal_Strength_dBm": np.random.normal(-85, 15, num_records),
        "Technician_Name": np.random.choice(["Jean", "Paul", "Marie", "Luc", "Alice", np.nan], num_records),
        "Install_Date": pd.to_datetime(np.random.choice(pd.date_range("2015-01-01", "2023-01-01"), num_records)).astype(str)
    }

    df = pd.DataFrame(data)

    # INTRODUCTION INTENTIONNELLE D'ERREURS (Intentional Errors pour démontrer le nettoyage / l'utilité du RDBMS)

    # 1. Duplicates
    duplicates = df.sample(20)
    df = pd.concat([df, duplicates], ignore_index=True)

    # 2. Missing Values (NaNs)
    df.loc[df.sample(30).index, 'Province'] = np.nan
    df.loc[df.sample(40).index, 'Signal_Strength_dBm'] = np.nan

    # 3. Mixed Formats dans les dates (Simulation d'erreurs humaines de tableur)
    bad_dates_idx = df.sample(25).index
    df.loc[bad_dates_idx, 'Install_Date'] = "Inconnue"
    bad_dates_idx2 = df.sample(25).index
    df.loc[bad_dates_idx2, 'Install_Date'] = "12/31/2020" # Format US mélangé

    # 4. Inconsistance de casse dans les statuts déjà introduite ("active", " OFF-LINE ", "")

    # Sauvegarde en Excel
    filename = "Kadea_Telco_RDC_Dataset.xlsx"
    df.to_excel(filename, index=False)
    print(f"Dataset sale généré avec succès : {filename}")

if __name__ == "__main__":
    create_dirty_dataset()
