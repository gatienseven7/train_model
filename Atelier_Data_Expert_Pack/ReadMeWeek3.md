# Atelier Technique Data Expert : Excel Avancé & SQL/NoSQL (Guide Complet)

Ce document est le manuel officiel pour l'atelier de 3h30. Il couvre des techniques de Data Wrangling allant du débutant à l'expert.

---

## 🛠️ Partie 1 : Environnement

1.  **Excel** : Version 2016+ (Pour Power Query, Power Pivot).
2.  **Google Sheets** : Compte Google actif.
3.  **Fichier** : `dataset_clients_advanced.xlsx` (Contient 3 feuilles : Clients, Offres, Logs).

---

## 🧹 Partie 2 : Excel - Maîtrise Totale du Tri & Filtre

### 2.1. Le Tri Multi-Niveaux (Custom Sort)
On veut trier les clients par **Région**, puis par **Offre**, puis par **Nom**.

1.  Sélectionnez tout le tableau "Clients_Brut" (Ctrl+A).
2.  Onglet **Données** > **Trier** (Gros bouton carré).
3.  Ajoutez des niveaux :
    *   Trier par : **Region** (A à Z).
    *   Ajouter un niveau : **Code_Offre** (A à Z).
    *   Ajouter un niveau : **Nom_Complet** (A à Z).
4.  OK. Vos données sont parfaitement hiérarchisées.

### 2.2. Filtres Avancés (Criteria Range)
Le filtre standard (flèches) est limité. Le filtre avancé permet des conditions complexes (ex: "Région = IDF OU Offre = FIBRE" ET "Prix > 30").

1.  Copiez les en-têtes de votre tableau (Ligne 1) et collez-les à côté (ex: colonne M).
2.  Sous "Region" (M2), écrivez `IDF`.
3.  Sous "Statut_Paiement" (O2), écrivez `Retard`.
4.  Allez dans **Données** > **Avancé** (À côté de l'entonnoir).
5.  **Plage** : Votre tableau complet ($A$1:$F$500).
6.  **Zone de critères** : Votre petit tableau à côté ($M$1:$O$2).
7.  OK. Excel ne montre que les parisiens en retard de paiement.

### 2.3. Slicers (Segment)
Pour rendre le filtrage interactif (Dashboard).

1.  Transformez votre plage en Tableau officiel : **Insertion** > **Tableau** (Ctrl+L).
2.  Onglet **Création de tableau** > **Insérer un segment**.
3.  Cochez **Region** et **Statut**.
4.  Cliquez sur les boutons pour filtrer instantanément !

---

## 🔗 Partie 3 : Excel - Croisement de Données (VLOOKUP & XLOOKUP)

Vous avez les codes offres dans "Clients_Brut" mais les PRIX sont dans la feuille "Ref_Offres". Il faut les rapatrier.

### 3.1. RECHERCHEV (VLOOKUP) - La Classique
Dans une nouvelle colonne "Prix" sur la feuille Clients :

```excel
=RECHERCHEV(D2; Ref_Offres!$A$2:$C$6; 2; FAUX)
```
*   **D2** : Ce que je cherche (Code Offre).
*   **Ref_Offres!A:C** : Où je cherche (La table des prix).
*   **2** : Je veux la 2ème colonne (Le Prix).
*   **FAUX** : Je veux une correspondance exacte.

### 3.2. RECHERCHEX (XLOOKUP) - La Moderne (Office 365)
Plus robuste, ne casse pas si on insère des colonnes.

```excel
=RECHERCHEX(D2; Ref_Offres!A:A; Ref_Offres!B:B; "Pas trouvé")
```
*   Cherche D2 dans la colonne A (Codes).
*   Renvoie la valeur correspondante de la colonne B (Prix).

---

## ⚡ Partie 4 : Power Query - L'Arme Absolue

Pour nettoyer les dates mixtes (FR/US) et fusionner sans formules.

### 4.1. Charger et Fusionner (Merge Queries)
1.  Chargez les deux tables ("Clients" et "Ref_Offres") dans Power Query (**Données > À partir de tableau**).
2.  Dans Power Query, cliquez sur la requête "Clients".
3.  Ruban **Accueil** > **Fusionner des requêtes** (Merge Queries).
4.  Choisissez la table "Ref_Offres" en bas.
5.  Cliquez sur la colonne commune "Code_Offre" dans les deux tables.
6.  Validez. Une colonne "Table" apparaît. Cliquez sur la double flèche pour "Développer" et cochez "Prix".
    *   *Magie : Plus besoin de VLOOKUP !*

### 4.2. Langage M (Pour les Dates complexes)
Parfois, l'interface ne suffit pas.
Dans la barre de formule Power Query, pour créer une date propre à partir d'un texte "YYYYMMDD" :

```powerquery
Date.FromText([Date_Souscription])
```

Pour une condition personnalisée (Colonne conditionnelle) :
```powerquery
if [Statut] = "Retard" then "A Relancer" else "OK"
```

---

## 🤖 Partie 5 : Automatisation VBA (Excel CRUD)

Transformons Excel en mini-application.
Faites **Alt + F11** pour ouvrir l'éditeur VBA.
**Insertion > Module**. Collez ce code :

### 5.1. CREATE (Ajouter un client via Formulaire)

```vba
Sub AjouterClient()
    Dim ws As Worksheet
    Set ws = Sheets("Clients_Brut")

    ' Trouver la première ligne vide
    Dim derniereLigne As Long
    derniereLigne = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1

    ' Lire les valeurs des cases B1, B2, B3 (votre formulaire)
    ws.Cells(derniereLigne, 1).Value = Range("B1").Value ' ID
    ws.Cells(derniereLigne, 2).Value = Range("B2").Value ' Nom
    ws.Cells(derniereLigne, 3).Value = Now ' Date Auto

    MsgBox "Client ajouté avec succès !", vbInformation

    ' Vider le formulaire
    Range("B1:B2").ClearContents
End Sub
```

### 5.2. READ (Rechercher un client)
```vba
Sub TrouverClient()
    Dim idCherche As String
    idCherche = InputBox("Entrez l'ID Client (ex: C001)")

    Dim c As Range
    Set c = Sheets("Clients_Brut").Range("A:A").Find(idCherche, LookIn:=xlValues)

    If Not c Is Nothing Then
        MsgBox "Trouvé : " & c.Offset(0, 1).Value & " (" & c.Offset(0, 5).Value & ")", vbOKOnly
    Else
        MsgBox "Client introuvable.", vbExclamation
    End If
End Sub
```

Assignez ces macros à des boutons (Formes rectangulaires > Clic droit > Affecter une macro).

---

## 🌐 Partie 6 : Google Sheets - Le Cloud Power

Les fonctions sont souvent identiques, mais certaines sont uniques et ultra-puissantes.

### 6.1. La fonction QUERY (Le SQL dans Sheets !)
C'est LA fonction qui tue le game. Elle permet d'écrire du pseudo-SQL directement dans une cellule.

Dans une nouvelle feuille GSheet :
```excel
=QUERY(Clients_Brut!A:F; "SELECT B, D, F WHERE F = 'Retard' ORDER BY B"; 1)
```
*   Traduction : Sélectionne les colonnes Nom (B), Offre (D) et Statut (F) où le statut est "Retard", trié par Nom.
*   C'est dynamique : Si la source change, le résultat change !

### 6.2. FILTER (Array Formula)
Plus simple que Query :
```excel
=FILTER(A:F; F:F="Retard"; E:E="IDF")
```
Renvoie toutes les lignes où Statut=Retard ET Region=IDF.

---

## 🗄️ Partie 7 : SQL & NoSQL (CRUD Avancé)

### Outils :
*   SQL : [DB Fiddle (PostgreSQL)](https://www.db-fiddle.com/)
*   NoSQL : [MongoDB Atlas](https://www.mongodb.com/cloud/atlas)

### 7.1. SQL : Jointures (JOIN)
Comme le VLOOKUP, mais en base de données.

```sql
-- Création des 2 tables
CREATE TABLE Offres (code VARCHAR(20), prix DECIMAL);
INSERT INTO Offres VALUES ('FIBRE', 39.99), ('ADSL', 29.99);

CREATE TABLE Clients (nom VARCHAR(50), code_offre VARCHAR(20));
INSERT INTO Clients VALUES ('Jean', 'FIBRE'), ('Paul', 'ADSL');

-- La Jointure (INNER JOIN)
SELECT Clients.nom, Offres.prix
FROM Clients
INNER JOIN Offres ON Clients.code_offre = Offres.code;
```

### 7.2. NoSQL : Structure Imbriquée
Pas de jointure en NoSQL ! On imbrique.

```javascript
db.clients.insertOne({
  "nom": "Jean",
  "offre": {
    "code": "FIBRE",
    "prix": 39.99,
    "details": ["TV", "Internet", "Tel"]
  }
})
```
Pour chercher tous ceux qui ont la TV :
```javascript
db.clients.find({"offre.details": "TV"})
```

---

## ✅ Checklist Expert

*   [ ] Je sais faire un filtre avancé avec plage de critères sur Excel.
*   [ ] Je maîtrise XLOOKUP pour croiser deux feuilles.
*   [ ] J'ai utilisé Power Query pour fusionner (Merge) deux tables sans formules.
*   [ ] J'ai créé un bouton VBA pour automatiser une tâche simple.
*   [ ] J'ai testé la fonction `=QUERY()` sur Google Sheets.
*   [ ] Je comprends la différence entre un JOIN SQL et l'imbrication NoSQL.

**Félicitations, vous êtes prêt pour le Big Data !**
