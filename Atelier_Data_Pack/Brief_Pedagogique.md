# Brief Pédagogique : Atelier "Nettoyage & Gestion des Données"

**Titre du module :** Gestion de Données Fondamentale (Data Wrangling & SQL Basics)
**Durée estimée :** 3h (1h Théorie + 2h Pratique)
**Public cible :** Débutant à Intermédiaire (Pas de prérequis techniques forts).

---

## 🎯 Objectifs Pédagogiques

À la fin de cet atelier, l'apprenant sera capable de :

1.  **Identifier** les anomalies courantes dans un jeu de données (doublons, valeurs manquantes, formats incohérents).
2.  **Appliquer** des techniques de nettoyage sur Excel (Filtres, Recherche/Remplacement) et Power Query (Transformation de types, Split column).
3.  **Distinguer** les concepts de Base de Données Relationnelle (SQL) et Non-Relationnelle (NoSQL).
4.  **Exécuter** les opérations fondamentales CRUD (Create, Read, Update, Delete) via des requêtes SQL simples.
5.  **Comprendre** l'importance de la qualité des données (Principe "Garbage In, Garbage Out").

---

## 📝 Description du Projet Pratique

**Contexte :**
Vous êtes Data Analyst Junior chez "TelcoNet", un opérateur téléphonique fictif. Le service marketing vous envoie un fichier Excel contenant la liste des nouveaux abonnés du mois dernier. Ils veulent lancer une campagne SMS, mais le fichier est inexploitable : noms mélangés, dates au format américain, doublons...

**Mission :**
1.  **Nettoyer** le fichier `dataset_clients_raw.xlsx` pour obtenir une liste propre et unique de clients actifs.
2.  **Structurer** ces données pour qu'elles soient prêtes à être importées dans la base de données de l'entreprise.
3.  **Simuler** l'insertion et la mise à jour de ces clients dans une base SQL via DB Fiddle.

**Livrables attendus de l'étudiant :**
*   Le fichier Excel nettoyé (`dataset_clients_clean.xlsx`).
*   Une capture d'écran de sa requête SQL `SELECT` fonctionnelle montrant les clients filtrés par solde.

---

## 📊 Grille d'Évaluation (KPIs)

| Compétence | Indicateur de Réussite (KPI) | Points |
| :--- | :--- | :--- |
| **Qualité des Données (Excel)** | 0 Doublon restant dans le fichier final. | 20 |
| **Nettoyage (Excel/Power Query)** | La colonne "Nom_Client" est correctement séparée en "Nom" et "Prénom". | 20 |
| **Formatage (Excel)** | Toutes les dates sont au format uniforme (JJ/MM/AAAA) et reconnues comme dates par Excel. | 15 |
| **Compréhension SQL (CRUD)** | La requête `INSERT` insère correctement les données avec les bons types (String vs Number). | 15 |
| **Logique de Requête (SQL)** | La requête `SELECT` utilise correctement une clause `WHERE` pour filtrer (ex: Solde > 0). | 15 |
| **Rigueur (Bonnes Pratiques)** | Le fichier rendu ne contient pas de lignes vides parasites ni de colonnes inutiles. | 15 |

**Note Totale : /100**

---

## 💡 Contenu Sommaire de l'Atelier

1.  **Introduction (15 min) :** Présentation PPTX. Rappel des enjeux (Garbage In, Garbage Out).
2.  **Démonstration Excel (45 min) :**
    *   Les pièges des fichiers CSV/Excel bruts.
    *   Démonstration des fonctions de base (Tri vs Filtre, Doublons).
    *   Introduction à Power Query pour les cas complexes (Dates, Séparateurs).
3.  **Pause (10 min)**
4.  **Démonstration SQL/NoSQL (45 min) :**
    *   Concept de Table vs Document.
    *   Live coding sur DB Fiddle : Création de table, Insertion, Lecture.
    *   Comparaison visuelle avec un document JSON (MongoDB).
5.  **Atelier Pratique (45 min) :** Les étudiants réalisent la mission "TelcoNet" en autonomie avec le support du `ReadMeWeek3.md`.
6.  **Q&A et Synthèse (15 min).**
