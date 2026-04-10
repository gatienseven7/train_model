# 🧠 Pop Quiz : Module 2 - Week 1 (Relational Database Foundations & Schema Design)

Voici une série de questions "Pop Quiz" à poser à l'oral à la fin de chaque bloc pour tester la compréhension en temps réel des étudiants, tout en gardant notre approche *Smart-Fun* et le *bilinguisme stratégique*.

---

## 🔴 BLOC 1 : CORE CONCEPTS

**Question 1 : Quel est le principal problème qui survient quand on utilise Excel comme base de données pour un million de lignes et plus ?**
- A) Il devient soudainement payant et très cher.
- B) Il fait face au "Excel Ceiling" : l'application suffoque, les performances chutent et il ne peut plus stocker de données de façon sécurisée.
- C) Les couleurs des cellules s'effacent mystérieusement.
- D) Il refuse d'enregistrer des données qui ne sont pas en anglais.

> **Réponse Correcte : B**
> **Explication :** C'est le fameux *Excel Ceiling*. Excel est un outil incroyable pour l'analyse, mais ce n'est pas un *RDBMS*. Au-delà d'un certain volume, le logiciel manque de *Scalability* et met en péril la *Data Integrity*.

**Question 2 : Pourquoi une banque confierait-elle son système à un RDBMS plutôt qu'à une base de données NoSQL ou un tableur ?**
- A) Pour l'atomicité des transactions et garantir la *Data Integrity* (aucune opération n'est laissée à moitié faite).
- B) Parce que c'est plus joli visuellement.
- C) Parce que ça permet de mettre des emojis dans les mots de passe.
- D) Pour permettre de stocker des documents au format texte non structuré.

> **Réponse Correcte : A**
> **Explication :** Un *RDBMS* (Relational Database Management System) offre une rigidité salvatrice. Il s'assure que si vous transférez 100$, ils sont bien débités d'un côté ET crédités de l'autre, ou l'opération entière est annulée. C'est la base de la *Data Integrity*.

---

## 🔴 BLOC 2 : TECHNICAL VOCAB

**Question 3 : Dans le jargon de l'architecte de données, comment appelle-t-on une ligne dans une table ?**
- A) Un *Row-bot*
- B) Un *Field*
- C) Un *Record*
- D) Un *Slot*

> **Réponse Correcte : C**
> **Explication :** Une ligne représente une occurrence unique de votre objet du monde réel, on l'appelle un *Record*. Les colonnes qui décrivent ce record sont les *Fields*.

**Question 4 : Quel est le rôle d'une *Foreign Key* (FK) ?**
- A) Elle sert à encrypter la base de données avec des algorithmes étrangers.
- B) Elle pointe vers la *Primary Key* (PK) d'une autre table pour créer un lien relationnel entre les données.
- C) C'est un identifiant généré aléatoirement qu'on ne peut utiliser qu'une fois.
- D) C'est le mot de passe pour accéder au serveur.

> **Réponse Correcte : B**
> **Explication :** La *Foreign Key* est le ciment du modèle relationnel. Elle permet d'associer, par exemple, un client (identifié par sa *Primary Key*) à sa commande correspondante sans avoir à dupliquer toutes les informations du client.

---

## 🔴 BLOC 3 : WORKSHOP

**Question 5 : Que se passe-t-il si vous tentez d'insérer un *Record* dans la table `Orders` avec un `Client_ID` (en *Foreign Key*) qui n'existe pas dans la table `Customers` ?**
- A) Le SGBD crée automatiquement un client fantôme pour vous dépanner.
- B) Le SGBD rejette l'insertion et génère une erreur pour préserver la *Data Integrity*.
- C) Le SGBD plante et vous devez redémarrer votre ordinateur.
- D) La commande est enregistrée dans une table de brouillon.

> **Réponse Correcte : B**
> **Explication :** C'est tout l'intérêt du modèle relationnel ! Les contraintes empêchent la création d'informations orphelines. Si le client n'existe pas, la base refuse catégoriquement l'ajout pour garantir la *Data Integrity*.

---

## 🔴 BLOC 4 : EXERCISE (Design Thinking)

**Question 6 : Lors de la conception d'un *Conceptual Data Model* (CDM), comment gère-t-on une relation "Many-to-Many" (par exemple : Un auteur peut écrire plusieurs livres, et un livre peut avoir plusieurs auteurs) ?**
- A) On met tous les noms d'auteurs séparés par des virgules dans le *Field* `Auteur` de la table `Books`.
- B) On duplique le livre autant de fois qu'il y a d'auteurs.
- C) On crée une table intermédiaire (ex: `Book_Authors`) qui liera les *Primary Keys* des tables `Books` et `Authors`.
- D) On choisit l'auteur principal et on ignore les autres pour simplifier le *Database Schema*.

> **Réponse Correcte : C**
> **Explication :** Mettre plusieurs valeurs dans un même champ (séparées par des virgules) brise les règles de normalisation ! Pour gérer une cardinalité complexe ("Many-to-Many"), on doit créer une table de jonction qui contient les *Foreign Keys* des deux entités. C'est l'essence même du *Database Schema* design.
