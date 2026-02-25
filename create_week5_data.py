from openpyxl import Workbook
import random

def create_week5_dataset():
    wb = Workbook()

    # --- Sheet 1: Ventes Massives (Pour justifier Power BI) ---
    ws_ventes = wb.active
    ws_ventes.title = "Ventes_Globales_2023"
    ws_ventes.append(["Date", "Produit", "Categorie", "Region", "Vendeur", "Quantite", "Prix_Unitaire", "Total"])

    produits = {
        "PC Portable": ("Informatique", 800),
        "Ecran 27": ("Informatique", 200),
        "Souris": ("Accessoires", 25),
        "Clavier": ("Accessoires", 40),
        "Imprimante": ("Bureautique", 150),
        "Chaise": ("Mobilier", 120),
        "Bureau": ("Mobilier", 250)
    }

    regions = ["Nord", "Sud", "Est", "Ouest", "Paris"]
    vendeurs = ["Alice", "Bob", "Charlie", "David", "Emma"]

    # Generate 1000 rows (Big enough to show Excel limits for visuals)
    for i in range(1, 1001):
        date = f"2023-{random.randint(1,12):02d}-{random.randint(1,28):02d}"
        prod_name = random.choice(list(produits.keys()))
        cat, prix = produits[prod_name]
        region = random.choice(regions)
        vendeur = random.choice(vendeurs)
        qte = random.randint(1, 10)
        total = qte * prix

        ws_ventes.append([date, prod_name, cat, region, vendeur, qte, prix, total])

    # --- Sheet 2: Objectifs (Pour comparaison) ---
    ws_obj = wb.create_sheet("Objectifs_Vendeurs")
    ws_obj.append(["Vendeur", "Objectif_Annuel"])
    for v in vendeurs:
        ws_obj.append([v, random.randint(50000, 150000)])

    filename = "dataset_S5_bi.xlsx"
    wb.save(filename)
    print(f"Created {filename}")

if __name__ == "__main__":
    create_week5_dataset()
