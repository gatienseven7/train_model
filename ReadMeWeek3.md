# Atelier Technique S3 : De l'Excel au SQL/NoSQL (Guide Pratique)

Ce document est votre guide pas à pas pour l'atelier. Il vous accompagnera dans la transformation de données brutes et "sales" en informations exploitables, en passant d'Excel aux bases de données.

---

## 🛠️ Partie 1 : Préparation de l'Environnement

Avant de commencer, assurez-vous d'avoir les outils suivants :

1.  **Microsoft Excel** (Version 2016 ou plus récente recommandée pour Power Query).
2.  Le fichier de données : `dataset_clients_raw.xlsx` (Fourni dans le ZIP).
3.  Un navigateur web pour accéder aux outils SQL/NoSQL en ligne.

### Étape 1 : Ouvrir le fichier brut
*   Double-cliquez sur `dataset_clients_raw.xlsx`.
*   Observez les données. Vous remarquerez des problèmes typiques :
    *   Des cases vides.
    *   Des doublons (Lignes répétées).
    *   Des formats de dates différents (YYYY-MM-DD vs DD/MM/YYYY).
    *   Des noms mal écrits.

---

## 🧹 Partie 2 : Nettoyage de Données sur Excel (Niveau 1 - Basique)

Objectif : Nettoyer la liste des clients pour une campagne Telco.

### 2.1. Supprimer les Doublons
Les doublons faussent les analyses (on compte deux fois le même client).

1.  Sélectionnez toute votre table (Ctrl+A).
2.  Allez dans l'onglet **Données** (Data).
3.  Cliquez sur l'icône **Supprimer les doublons** (Remove Duplicates).

> **[Capture d'écran : Ruban Excel > Onglet Données > Groupe Outils de données > Bouton Supprimer les doublons]**
> *L'icône ressemble à deux colonnes bleu/blanc avec une petite croix rouge.*

4.  Une fenêtre s'ouvre. Assurez-vous que toutes les colonnes sont cochées.
5.  Cliquez sur **OK**. Excel vous dira combien de lignes ont été supprimées.

### 2.2. Rechercher et Remplacer (Correction rapide)
On voit que certains montants ont des virgules et d'autres des points (29.99 vs 29,99), ce qui empêche les calculs.

1.  Sélectionnez la colonne **Montant_Forfait**.
2.  Appuyez sur **Ctrl + H** (ou Accueil > Rechercher et sélectionner > Remplacer).
3.  Dans "Rechercher" (Find what), tapez : `,` (virgule).
4.  Dans "Remplacer par" (Replace with), tapez : `.` (point).
5.  Cliquez sur **Remplacer tout** (Replace All).

> **[Capture d'écran : Fenêtre "Rechercher et remplacer"]**
> *Champ Rechercher : , | Champ Remplacer par : .*

### 2.3. Filtrer les Données Vides
Nous voulons supprimer les clients qui n'ont pas de numéro de téléphone.

1.  Sélectionnez la ligne d'en-tête (Ligne 1).
2.  Allez dans **Données** > **Filtrer** (Filter). Des petites flèches apparaissent sur chaque colonne.
3.  Cliquez sur la flèche de la colonne **Telephone**.
4.  Décochez tout, et ne cochez que **(Vides)** ou **(Blanks)** tout en bas.
5.  Les lignes vides apparaissent. Sélectionnez ces lignes (sur les numéros de ligne à gauche), faites **Clic Droit > Supprimer la ligne** (Delete Row).
6.  Retournez dans le filtre et faites **Effacer le filtre** (Clear Filter) pour revoir vos données propres.

---

## 🚀 Partie 3 : Nettoyage Avancé avec Power Query (Niveau 2 - Pro)

Power Query permet d'automatiser ce nettoyage. C'est l'outil secret des pros de la data.

### 3.1. Charger les données dans Power Query
1.  Sélectionnez vos données dans Excel.
2.  Allez dans **Données** > **À partir de tableau ou d'une plage** (From Table/Range).
3.  Une nouvelle fenêtre s'ouvre : C'est l'éditeur Power Query.

> **[Capture d'écran : Ruban Données > Groupe Récupérer et transformer des données > Bouton "À partir de tableau/plage"]**

### 3.2. Uniformiser les Dates
Power Query détecte parfois mal les dates mixtes (FR vs US).

1.  Cliquez sur l'icône "ABC/123" à gauche du titre de la colonne **Date_Inscription**.
2.  Choisissez **Date**.
3.  Si des erreurs apparaissent (`Error`), annulez l'étape (croix rouge à droite dans "Étapes appliquées").
4.  Faites **Clic Droit** sur la colonne > **Modifier le type** > **Utilisant les paramètres régionaux...** (Using Locale).
5.  Choisissez **Date** et **Anglais (États-Unis)** ou **Français (France)** selon ce qui corrige vos données.

### 3.3. Fractionner une colonne (Split)
Imaginons que la colonne "Nom_Client" contient "Dupont Jean" et on veut séparer Nom et Prénom.

1.  Sélectionnez la colonne **Nom_Client**.
2.  Allez dans l'onglet **Accueil** > **Fractionner la colonne** (Split Column) > **Par délimiteur** (By Delimiter).
3.  Choisissez **Espace**.
4.  Cliquez sur **OK**. Vous avez maintenant deux colonnes. Renommez-les "Nom" et "Prénom".

> **[Capture d'écran : Ruban Power Query > Onglet Accueil > Bouton Fractionner la colonne > Par délimiteur]**

### 3.4. Charger le résultat
1.  Cliquez sur le bouton tout à gauche **Fermer et charger** (Close & Load).
2.  Une nouvelle feuille Excel se crée avec vos données toutes propres !

---

## 💾 Partie 4 : Introduction aux Bases de Données (SQL vs NoSQL)

Maintenant que nos données sont propres, nous allons voir comment les gérer dans une vraie base de données.

### Outil utilisé : DB Fiddle (SQL)
Allez sur [https://www.db-fiddle.com/](https://www.db-fiddle.com/) et choisissez **PostgreSQL 15**.

### 4.1. CREATE (Créer la structure)
Contrairement à Excel où on écrit direct, en SQL il faut définir le "moule" (la table).

Collez ceci dans la partie GAUCHE (Schema SQL) :

```sql
CREATE TABLE Clients (
    id SERIAL PRIMARY KEY,
    nom VARCHAR(50),
    email VARCHAR(100),
    solde DECIMAL(10, 2)
);
```

> **Explication :** On crée une boîte "Clients" avec des étiquettes précises (Texte, Décimal...).

### 4.2. INSERT (Ajouter des données - Create du CRUD)
Toujours à GAUCHE, en dessous :

```sql
INSERT INTO Clients (nom, email, solde) VALUES
('Jean Dupont', 'jean@email.com', 29.99),
('Sophie Martin', 'sophie@test.fr', 45.00),
('Lucas Bernard', 'lucas@yahoo.com', 100.50);
```

Cliquez sur **RUN** en haut. Rien ne s'affiche ? C'est normal ! Vous avez juste stocké les données.

### 4.3. SELECT (Lire les données - Read du CRUD)
Maintenant, interrogeons la base. Dans la partie DROITE (Query SQL) :

**Exemple 1 : Tout voir**
```sql
SELECT * FROM Clients;
```
*Cliquez sur RUN. Vous voyez votre tableau.*

**Exemple 2 : Filtrer (Le "Filtre" d'Excel)**
```sql
SELECT * FROM Clients WHERE solde > 40;
```
*Affiche uniquement Sophie et Lucas.*

**Exemple 3 : Trier (Le "Tri" d'Excel)**
```sql
SELECT * FROM Clients ORDER BY solde DESC;
```
*Trie du plus riche au moins riche.*

> **[Capture d'écran : DB Fiddle avec le code SQL à gauche et le résultat tableau à droite]**

### 4.4. UPDATE (Mettre à jour - Update du CRUD)
Jean a payé sa facture, son solde change. (Partie DROITE) :

```sql
UPDATE Clients SET solde = 19.99 WHERE nom = 'Jean Dupont';
SELECT * FROM Clients; -- Pour vérifier
```

### 4.5. DELETE (Supprimer - Delete du CRUD)
Sophie résilie son abonnement.

```sql
DELETE FROM Clients WHERE nom = 'Sophie Martin';
SELECT * FROM Clients; -- Sophie a disparu
```

---

## 🍃 Partie 5 : NoSQL (MongoDB) - La souplesse

Le NoSQL stocke les données comme des documents (fiches), pas des tableaux.
Outil : Essayons de visualiser le concept JSON.

Dans une base NoSQL (comme MongoDB), notre client Jean ressemblerait à ceci :

```json
{
  "_id": 1,
  "nom": "Jean Dupont",
  "contact": {
    "email": "jean@email.com",
    "tel": "0612345678"
  },
  "historique_achats": ["Forfait A", "Option B"]
}
```

> **Différence clé :** Dans SQL, pour ajouter "historique_achats", il aurait fallu créer une autre table complexe. Ici, on l'écrit juste dans le document !

### Comparaison des commandes

| Action | SQL | NoSQL (MongoDB) |
| :--- | :--- | :--- |
| **Créer** | `INSERT INTO...` | `db.clients.insertOne({...})` |
| **Lire** | `SELECT * FROM...` | `db.clients.find({})` |
| **Modifier** | `UPDATE...` | `db.clients.updateOne(...)` |
| **Supprimer** | `DELETE FROM...` | `db.clients.deleteOne(...)` |

---

## ✅ Checklist de Fin d'Atelier

*   [ ] J'ai nettoyé mon fichier Excel (Doublons, Vides).
*   [ ] J'ai utilisé Power Query pour séparer une colonne.
*   [ ] J'ai exécuté ma première requête SQL `SELECT`.
*   [ ] J'ai compris la différence entre une ligne (SQL) et un document (NoSQL).

**Bravo ! Vous avez fait vos premiers pas de Data Engineer.**
