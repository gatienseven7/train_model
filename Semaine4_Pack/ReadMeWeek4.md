# Semaine 4 : Logique & Visualisation de Données (Guide Pratique)

**Objectif :** Transformer des chiffres bruts en décisions (Logique SI) et en histoires visuelles (Graphiques).

---

## 🧠 Partie 1 : La Logique Conditionnelle (SI / IF)

L'ordinateur est bête. Il faut lui dire quoi faire.
La fonction **SI** (IF en anglais) est la base de toute l'informatique :
> *"Si cette condition est VRAIE, fais ceci. Sinon, fais cela."*

### Exercice 1 : Admis ou Recalé ? (Feuille "Notes_Examens")
Nous avons une liste d'élèves avec leur moyenne. Nous voulons écrire automatiquement "Admis" ou "Recalé".

1.  Ouvrez le fichier `dataset_S4_logique_viz.xlsx`.
2.  Allez sur la feuille **Notes_Examens**.
3.  Cliquez en **E2** (Sous "Resultat_Attendu").
4.  Écrivez la formule :
    *   **Excel (Français)** : `=SI(D2>=10; "Admis"; "Recalé")`
    *   **Google Sheets (Anglais/Français)** : `=IF(D2>=10; "Admis"; "Recalé")`
5.  Validez. Étirez la formule vers le bas (double-clic sur le petit carré en bas à droite de la cellule).

> **Analyse :**
> *   `D2>=10` : C'est le **Test**. Est-ce que la moyenne est supérieure ou égale à 10 ?
> *   `"Admis"` : C'est la **Valeur si Vrai**.
> *   `"Recalé"` : C'est la **Valeur si Faux**.

---

## 📊 Partie 2 : Choisir le Bon Graphique

Une image vaut 1000 mots, mais une mauvaise image ment 1000 fois.

### Règle d'Or :
*   **Comparer des quantités** (Qui a vendu le plus ?) -> **Histogramme (Barres)**.
*   **Voir une évolution** (Comment les ventes changent mois par mois ?) -> **Courbe (Ligne)**.
*   **Voir une répartition** (Quelle part du budget pour le Loyer ?) -> **Camembert (Secteurs)**. *Attention : À éviter s'il y a trop de parts !*

---

## 🎨 Partie 3 : Création de Graphiques (Atelier)

### Exercice 2 : Le Meilleur Vendeur (Histogramme)
1.  Allez sur la feuille **Ventes_Mensuelles**.
2.  Sélectionnez les colonnes **Vendeur** et **Total_Trimestre** (Maintenez Ctrl pour sélectionner deux colonnes non adjacentes si besoin).
3.  **Excel** : Onglet **Insertion** > **Histogramme** (Premier icône de barres).
4.  **Google Sheets** : Menu **Insertion** > **Graphique**.
5.  Admirez le résultat. Qui est le meilleur ? (La barre la plus haute).
6.  **Important :** Ajoutez un titre ! "Ventes Totales par Vendeur (Q1)". Un graphique sans titre ne veut rien dire.

### Exercice 3 : L'Évolution des Ventes (Courbe)
1.  Toujours sur **Ventes_Mensuelles**.
2.  Sélectionnez tout le tableau (Vendeur + Jan/Fev/Mars).
3.  **Excel** : Insertion > Graphique Recommandé > **Courbe**.
4.  **Google Sheets** : Type de graphique > **Courbe**.
5.  Vous voyez maintenant la tendance pour chaque vendeur.

### Exercice 4 : La Répartition du Budget (Camembert)
1.  Allez sur la feuille **Budget_Projet**.
2.  Sélectionnez la colonne **Categorie** et **Depense_Reelle**.
3.  **Excel** : Insertion > **Graphique en secteurs** (Camembert 2D).
4.  **Google Sheets** : Type de graphique > **Secteurs**.
5.  Ajoutez les étiquettes de données (Clic droit sur le camembert > Ajouter des étiquettes) pour voir les pourcentages.

---

## ✅ Checklist de Validation

*   [ ] J'ai utilisé la fonction `=SI()` pour automatiser une décision.
*   [ ] Je sais faire la différence entre l'axe X (Horizontal - Catégories) et l'axe Y (Vertical - Valeurs).
*   [ ] J'ai créé un Histogramme pour comparer.
*   [ ] J'ai créé un Camembert pour montrer des proportions.
*   [ ] J'ai toujours mis un **Titre** et des **Légendes** à mes graphiques.

**Bravo ! Vous savez maintenant faire parler les chiffres.**
