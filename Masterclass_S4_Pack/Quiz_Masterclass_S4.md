# Quiz de Validation : De la Donnée à l'Action (Masterclass S4)

**Objectif :** Valider la compréhension des concepts clés de logique conditionnelle et de Data Storytelling abordés lors de la Masterclass.

---

### Question 1 : Le rôle de la fonction logique (SI/IF)
Dans le contexte d'un tableau de bord professionnel, quel est l'apport principal d'une fonction logique `=SI()` par rapport à un simple calcul mathématique (ex: Somme) ?
*   A) Elle permet de créer des graphiques plus colorés.
*   B) Elle remplace l'analyse humaine en permettant de trier les données par ordre alphabétique.
*   C) Elle automatise la prise de décision et génère des alertes métier (reporting par exception).
*   D) Elle compresse la taille du fichier Excel.

### Question 2 : Le choix du visuel (Comparaison)
Vous devez présenter à votre Direction Générale le chiffre d'affaires de vos 5 meilleurs commerciaux sur l'année écoulée pour désigner le "Vendeur de l'Année". Quel graphique choisissez-vous ?
*   A) Un graphique en secteurs (Camembert).
*   B) Un histogramme (Graphique en barres).
*   C) Une courbe.
*   D) Un nuage de points.

### Question 3 : Le choix du visuel (Évolution)
Vous souhaitez analyser la saisonnalité des ventes de votre produit phare mois par mois sur les 3 dernières années pour anticiper un pic de demande. Quel est le format le plus pertinent ?
*   A) Une courbe continue.
*   B) Un graphique en secteurs (Camembert) avec 36 parts.
*   C) Un tableau de données brut sans mise en forme.
*   D) Un graphique en radar.

### Question 4 : L'anatomie d'un graphique (X vs Y)
Lorsqu'on construit un graphique en deux dimensions pour raconter une histoire métier (Data Storytelling), comment se répartissent généralement les rôles entre les axes ?
*   A) L'axe X donne la Valeur (la métrique, ex: le CA), l'axe Y donne le Contexte (le temps, les catégories).
*   B) Les deux axes doivent toujours afficher des dates pour que le graphique soit valide.
*   C) L'axe X donne le Contexte (la segmentation, le temps : le "Qui" ou "Quand"), l'axe Y donne la Valeur métier (la performance : le "Combien").
*   D) L'axe Y sert uniquement à donner un titre au graphique.

### Question 5 : La Règle d'Edward Tufte (Le Data-Ink Ratio)
Lors de la conception d'un tableau de bord, vous appliquez le principe du "Data-Ink Ratio" (Maximiser l'encre dédiée aux données). Laquelle de ces actions **contredit** ce principe ?
*   A) Supprimer les lignes de quadrillage denses en arrière-plan du graphique.
*   B) Mettre l'histogramme en 3D avec une ombre portée pour le rendre "plus moderne".
*   C) Retirer une légende inutile si le titre du graphique explique déjà de quoi il s'agit.
*   D) S'assurer que les axes partent bien de zéro pour un graphique en barres.

---
---

## 🎯 Réponses (Pour le Formateur)

**Q1 : Réponse C.**
*Explication :* La fonction SI transforme un tableur statique en un moteur de règles. Elle permet de mettre en évidence (ex: via le mot "ALERTE" ou "OK") une situation nécessitant une action, sans que l'humain ait à lire chaque ligne.

**Q2 : Réponse B.**
*Explication :* Pour comparer des éléments de même nature (catégories) de manière claire, la longueur des barres d'un histogramme est l'outil le plus précis pour l'œil humain. Le camembert est fait pour les proportions d'un tout (100%).

**Q3 : Réponse A.**
*Explication :* La courbe (ou graphique en ligne) est conçue spécifiquement pour montrer la continuité et les tendances temporelles. Elle permet de repérer immédiatement la saisonnalité (les "pics" et les "creux").

**Q4 : Réponse C.**
*Explication :* L'axe horizontal (Abscisses - X) est le terrain de jeu, le contexte (les mois, les noms des commerciaux). L'axe vertical (Ordonnées - Y) est la jauge de performance, la valeur mesurée.

**Q5 : Réponse B.**
*Explication :* La 3D et les ombres portées n'apportent aucune information supplémentaire (elles n'affichent pas de "donnée"). Pire, la perspective fausse souvent la lecture des valeurs réelles. Le Data-Ink Ratio impose de supprimer tout ce qui est purement décoratif.