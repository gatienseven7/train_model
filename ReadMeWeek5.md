# Semaine 5 : Introduction à la Business Intelligence (Transition)

**Objectif :** Dépasser les limites d'Excel et découvrir l'interactivité avec Power BI (ou Looker Studio).

---

## 🚀 Partie 1 : Pourquoi la BI ? (Théorie)

Excel est génial, mais...
*   Il ralentit avec 1 million de lignes.
*   Il n'est pas "temps réel".
*   Il faut refaire les graphiques chaque mois.

**La BI (Business Intelligence), c'est :**
1.  **Connecter** (On ne copie-colle pas, on se connecte à la source).
2.  **Transformer** (Le nettoyage se fait une seule fois, comme dans Power Query).
3.  **Visualiser** (C'est interactif : je clique sur "Paris", tout le rapport filtre sur Paris).

---

## 🛠️ Partie 2 : Installation & Connexion

Nous utiliserons **Microsoft Power BI Desktop** (Gratuit, Windows uniquement).
*Si vous êtes sur Mac, utilisez **Google Looker Studio** (100% Web).*

### Étape 1 : Obtenir les Données
1.  Ouvrez Power BI Desktop.
2.  Cliquez sur **Importer des données à partir d'Excel**.
3.  Sélectionnez le fichier `dataset_S5_bi.xlsx`.
4.  Cochez la feuille **Ventes_Globales_2023**.
5.  Cliquez sur **Charger** (Load).

> **Sur Looker Studio :**
> Créer un rapport vide > Connecter à Google Sheets (Il faut d'abord importer le fichier Excel dans un Google Sheet).

---

## 📊 Partie 3 : "Glisser-Déposer" (Le Premier Dashboard)

Power BI fonctionne par "Drag & Drop". C'est comme un jeu de Lego.

### Exercice 1 : Chiffre d'Affaires par Vendeur (Histogramme)
1.  A droite, dans le panneau **Champs** (Fields), cochez `Vendeur`.
2.  Cochez aussi `Total` (Power BI comprend que c'est une somme).
3.  Power BI crée automatiquement un graphique !
4.  Changez le type de visuel (dans le panneau **Visualisations**) pour mettre un **Histogramme groupé**.

### Exercice 2 : Répartition par Catégorie (Anneau)
1.  Cliquez dans le vide (sur la page blanche) pour désélectionner le premier graph.
2.  Cochez `Categorie` et `Total`.
3.  Choisissez le visuel **Anneau** (Donut chart).

### Exercice 3 : La Magie de l'Interactivité
1.  Cliquez sur la part "Informatique" de votre Anneau.
2.  Regardez l'Histogramme des vendeurs : **Il bouge !**
3.  Il ne montre plus que les ventes informatiques pour chaque vendeur.
4.  C'est ça, la puissance de la BI.

### Exercice 4 : La Carte (Map)
1.  Cliquez dans le vide.
2.  Cochez `Region` (ou Ville si disponible).
3.  Choisissez le visuel **Carte** (Map).
4.  Vous voyez vos ventes géographiquement.

---

## ✅ Checklist de Transition

*   [ ] J'ai compris la différence entre un Fichier Excel (statique) et un Rapport BI (dynamique).
*   [ ] J'ai réussi à connecter mon fichier Excel à Power BI / Looker Studio.
*   [ ] J'ai créé 3 visuels différents sur la même page.
*   [ ] J'ai testé l'interaction (Cliques sur un graph pour filtrer les autres).

**Bravo ! Vous venez d'entrer dans le monde de la Data Analyse moderne.**
