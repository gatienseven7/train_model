# PROJET FIL ROUGE : KADEA TELCO
## Guide Technique et Quizz

Bienvenue dans le guide de réalisation du projet Kadea Telco. Ce document vous accompagne étape par étape dans la réalisation des 4 phases de ce défi complexe.

---

## 🛠️ Phase 1 : Le chaos des Logs et le défi du croisement

### Objectif
Nettoyer le fichier `Logs_Reseau.xlsx` et le croiser avec `Registre_Maintenance.xlsx`.

### Interface et Instructions (Excel & Google Sheets)
1. **Nettoyage des données (Data Cleansing)**
   - *Excel UI*: Allez dans l'onglet **Data** (Données) > **Remove Duplicates** (Supprimer les doublons).
   - *Google Sheets UI*: Allez dans **Data** > **Data cleanup** > **Remove duplicates**.

   *[ Insérer la capture d'écran de l'outil de suppression des doublons ici ]*

2. **Croisement de données (VLOOKUP / RECHERCHEV)**
   - Dans le fichier Logs, créez une colonne `Region`.
   - Utilisez la formule pour récupérer la région depuis le Registre de Maintenance en utilisant le `Cell_ID`.
   - **Attention** : Gérez les erreurs `#N/A` pour les ID d'antennes qui n'ont pas de correspondance (ex: avec `SIERREUR` ou `IFERROR`).

   *[ Insérer la capture d'écran de la formule VLOOKUP saisie dans la barre de formule ici ]*

---

## 📊 Phase 2 : L'offensive et la synthèse macroscopique

### Objectif
Créer un Tableau Croisé Dynamique pour résumer les temps de panne par Région et Technologie.

### Interface et Instructions
1. **Création du TCD (Pivot Table)**
   - *Excel UI*: **Insert** > **PivotTable**.
   - *Google Sheets UI*: **Insert** > **Pivot table**.

   *[ Insérer la capture d'écran de la fenêtre de création de Pivot Table ici ]*

2. **Configuration**
   - **Rows (Lignes)** : `Region` puis `Technologie`.
   - **Values (Valeurs)** : `Duree_Panne_Min` configuré en **Sum** (Somme) ET en **Average** (Moyenne).

---

## 💡 Phase 3 : Le labyrinthe de la Rétention

### Objectif
Dans le fichier `Clients_Churn.xlsx`, créer une colonne `Taux_Remise` basée sur des conditions complexes.

### Instructions Logiques
Vous devez imbriquer les conditions suivantes en utilisant `IF` (SI), `AND` (ET), et `OR` (OU) :
- Si le client a plus de 12 mois d'ancienneté **ET** possède le forfait "Premium 5G 150Go", il reçoit 20% de remise.
- Sinon, si le client a subi plus de 120 min de panne cumulée **OU** a fait plus de 3 plaintes, il reçoit 10% de remise.
- Dans tous les autres cas, la remise est de 0%.

   *[ Insérer la capture d'écran de la longue formule logique dans la cellule ici ]*

---

## 🚀 Phase 4 : La libération via la Business Intelligence (Power BI)

### Objectif
Importer vos données nettoyées dans Power BI et créer un tableau de bord interactif.

### Interface et Instructions
1. **Importation (Get Data)**
   - *Power BI UI*: Cliquez sur **Get Data** (Obtenir les données) > **Excel workbook**.

   *[ Insérer la capture d'écran du menu Get Data de PowerBI ici ]*

2. **Modélisation (Model View)**
   - Allez dans la vue modèle (**Model view** sur la barre latérale gauche).
   - Créez les relations entre la table Logs et la table Maintenance via l'ID de l'antenne.

   *[ Insérer la capture d'écran de la vue relationnelle avec le lien actif ici ]*

---

## 📝 Quizz de validation des compétences

**Q1 : Quelle fonction est la plus appropriée pour rapatrier le nom d'une région à partir d'un identifiant unique (ID) stocké dans un autre tableau ?**
- A) SUMIFS
- B) VLOOKUP (RECHERCHEV)
- C) CONCATENATE
- D) COUNTBLANK
> **Réponse correcte : B**
> *Explication : VLOOKUP permet de rechercher une valeur dans la première colonne d'une matrice et de renvoyer la valeur d'une autre colonne sur la même ligne.*

**Q2 : Lors de la création d'un Tableau Croisé Dynamique (Pivot Table) pour calculer la durée MOYENNE de panne, quel paramètre devez-vous ajuster dans le champ des valeurs (Values) ?**
- A) Modifier l'agrégation de 'Sum' à 'Average' dans les paramètres du champ de valeur (Value Field Settings).
- B) Déplacer le champ de 'Values' vers 'Rows'.
- C) Créer un segment (Slicer).
- D) Modifier le format du nombre en pourcentage.
> **Réponse correcte : A**
> *Explication : Par défaut, les valeurs numériques sont sommées (Sum). Il faut explicitement demander le calcul de la moyenne (Average).*

**Q3 : Quelle combinaison de fonctions permet d'évaluer si un client remplit DEUX conditions obligatoires simultanément pour obtenir une remise ?**
- A) IF et OR (SI et OU)
- B) SUM et IF (SOMME et SI)
- C) IF et AND (SI et ET)
- D) IFERROR (SIERREUR)
> **Réponse correcte : C**
> *Explication : La fonction AND (ET) vérifie que tous les arguments sont VRAIS, ce qui est parfait pour l'imbriquer comme condition d'un IF (SI).*

**Q4 : Dans Power BI, comment s'appelle l'étape où l'on relie deux tables via une colonne commune (comme l'ID Antenne) ?**
- A) Data Cleansing (Nettoyage des données)
- B) Data Visualization (Visualisation des données)
- C) Data Modeling (Modélisation des données)
- D) DAX Formatting
> **Réponse correcte : C**
> *Explication : La modélisation (Model View) est l'endroit où l'on définit l'architecture relationnelle des données (tables de faits et de dimensions).*
