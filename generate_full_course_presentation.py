from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

# Couleurs Kadea Academy
KADEA_RED = RGBColor(237, 28, 36)
ANTHRACITE = RGBColor(40, 40, 40)
WHITE = RGBColor(255, 255, 255)

def set_shape_text(shape, text, color=ANTHRACITE, font_size=Pt(16), bold=False, alignment=PP_ALIGN.LEFT):
    if not shape.has_text_frame:
        return
    text_frame = shape.text_frame
    text_frame.clear()
    p = text_frame.paragraphs[0]
    p.text = text
    p.font.color.rgb = color
    p.font.size = font_size
    p.font.bold = bold
    p.font.name = 'Calibri'
    p.alignment = alignment

def add_slide_with_content(prs, title_text, text_blocks, viz_text=""):
    slide = prs.slides.add_slide(prs.slide_layouts[5])

    # Title
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(1))
    set_shape_text(title_box, title_text, KADEA_RED, Pt(26), True)

    # Content
    content_box = slide.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(5), Inches(5))
    tf = content_box.text_frame
    tf.word_wrap = True

    for block in text_blocks:
        p = tf.add_paragraph()
        p.text = block['title']
        p.font.bold = True
        p.font.size = Pt(18)
        p.font.color.rgb = ANTHRACITE if not block.get('highlight') else KADEA_RED

        for bullet in block['bullets']:
            p_bullet = tf.add_paragraph()
            p_bullet.text = bullet
            p_bullet.font.size = Pt(14)
            p_bullet.level = 1

    # Visual Placeholder
    if viz_text:
        viz_box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(6.0), Inches(1.5), Inches(3.5), Inches(5))
        viz_box.fill.solid()
        viz_box.fill.fore_color.rgb = WHITE
        viz_box.line.color.rgb = KADEA_RED
        viz_box.line.width = Pt(2)
        set_shape_text(viz_box, viz_text, ANTHRACITE, Pt(12), False, PP_ALIGN.CENTER)

def create_full_presentation():
    prs = Presentation()

    # Slide 1: Titre
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    set_shape_text(slide.shapes.title, "Bases de données : Modélisation et SQL", KADEA_RED, Pt(32), True, PP_ALIGN.CENTER)
    set_shape_text(slide.placeholders[1], "Cours Complet, de la Conception à l'Analyse", ANTHRACITE, Pt(24), False, PP_ALIGN.CENTER)

    # Slide 2: Généralités
    add_slide_with_content(
        prs,
        "1. Généralités sur les Bases de Données",
        [
            {"title": "Définition", "bullets": ["Un ensemble de données stockées et structurées informatiquement pour faciliter la consultation et la modification."]},
            {"title": "Le Rôle du SGBD", "bullets": ["Le Système de Gestion de Bases de Données est le logiciel intermédiaire.", "Il gère la cohérence des données, les accès multiples (concurrents) et la sécurité.", "Il assure la reprise sur incident (ex: annulation après une coupure de courant)."]},
        ],
        "[VISUEL KADEA]\nSchéma: Utilisateur -> SGBD -> Fichiers de la base de données sur disque dur."
    )

    # Slide 3: Historique
    add_slide_with_content(
        prs,
        "2. Les Différents Types de SGBD (Historique)",
        [
            {"title": "Années 1960", "bullets": ["Hiérarchiques : données en arborescence.", "Réseaux : arborescence avec raccourcis."]},
            {"title": "Années 1970 - Modèle Relationnel", "bullets": ["1970 : Edgar F. Codd invente le modèle relationnel.", "Données organisées en tableaux liés entre eux. Langage SQL standardisé en 1974."]},
            {"title": "Années 2000 - NoSQL", "highlight": True, "bullets": ["Not Only SQL : Bases Clé/Valeur, Orientées Documents, Graphes.", "Pour le Big Data et la très haute disponibilité."]}
        ],
        "[VISUEL KADEA]\nFrise chronologique des SGBD, mettant en avant l'apparition du Modèle Relationnel et du NoSQL."
    )

    # Slide 4: Transactions ACID
    add_slide_with_content(
        prs,
        "3. La Notion de Transaction et l'ACIDité",
        [
            {"title": "La Transaction", "bullets": ["Une séquence indivisible d'actions (ex: virement bancaire = débit + crédit)."]},
            {"title": "Les Principes ACID", "highlight": True, "bullets": [
                "Atomicité : La transaction est exécutée en entier ou annulée totalement.",
                "Cohérence : La base passe d'un état valide à un autre état valide.",
                "Isolation : Deux transactions simultanées n'interfèrent pas.",
                "Durabilité : Une transaction validée survit aux pannes système."
            ]}
        ],
        "[VISUEL KADEA]\nIcône d'un coffre-fort avec les 4 piliers de l'ACID."
    )

    # Slide 5: Le Modèle Relationnel (Concepts)
    add_slide_with_content(
        prs,
        "4. Le Modèle Relationnel : Fondements",
        [
            {"title": "Relations (Tables)", "bullets": ["L'information est stockée dans des tables constituées de colonnes (attributs) et de lignes (enregistrements).", "Chaque colonne possède un domaine de valeurs précis (type, longueur, contraintes)."]},
            {"title": "Le Système de Clés", "highlight": True, "bullets": [
                "Clé Primaire : Identifiant unique d'une ligne dans une table.",
                "Clé Étrangère : Colonne pointant vers la clé primaire d'une autre table, créant le lien."
            ]}
        ],
        "[VISUEL KADEA]\nIllustration de deux tables (Ville et Personne). Un rayon laser relie la Clé Primaire 'Code_Ville' à la Clé Étrangère de la table Personne."
    )

    # Slide 6: Le Modèle Relationnel (Cardinalités)
    add_slide_with_content(
        prs,
        "5. Associations et Cardinalités",
        [
            {"title": "Types de Liens", "bullets": [
                "1-1 : Un pays a une seule capitale.",
                "1-N : Un pays possède plusieurs villes, une ville appartient à un seul pays."
            ]},
            {"title": "Le Cas Complexe : Plusieurs-à-Plusieurs (N-M)", "highlight": True, "bullets": [
                "Ex: Une personne travaille dans plusieurs entreprises, une entreprise emploie plusieurs personnes.",
                "Solution : Créer une 'Table d'Association' (ex: Travail) contenant les clés étrangères des deux tables."
            ]}
        ],
        "[VISUEL KADEA]\nSchéma expliquant la résolution d'une relation N-M par une table intermédiaire."
    )

    # Slide 7: Algèbre Relationnelle
    add_slide_with_content(
        prs,
        "6. L'Algèbre Relationnelle",
        [
            {"title": "Opérations sur 1 table", "bullets": [
                "Sélection : Filtre les lignes selon une condition.",
                "Projection : Ne conserve que certaines colonnes.",
                "Renommage : Modifie le nom d'une colonne."
            ]},
            {"title": "Opérations sur 2 tables", "highlight": True, "bullets": [
                "Produit Cartésien : Combine tous les éléments de deux tables.",
                "Jointure : Combine des éléments répondant à un critère de liaison."
            ]}
        ],
        "[VISUEL KADEA]\nInfographie montrant le concept visuel de la Projection (colonnes verticales en surbrillance) et de la Sélection (lignes horizontales en surbrillance)."
    )

    # Slide 8: Modélisation UML
    add_slide_with_content(
        prs,
        "7. Modélisation Conceptuelle avec UML",
        [
            {"title": "Diagramme de Classes UML", "bullets": [
                "Le nom de la relation en haut.",
                "La liste des colonnes en bas : nom_colonne: type(longueur).",
                "Les Clés Primaires sont soulignées, les Clés Étrangères précédées d'un #."
            ]},
            {"title": "Agrégation et Composition", "highlight": True, "bullets": [
                "Agrégation : Lien simple (ex: une maison est une agrégation de murs).",
                "Composition : Agrégation forte, la destruction du tout détruit la partie (ex: école et cycles scolaires)."
            ]}
        ],
        "[VISUEL KADEA]\nExemple de formalisme UML strict avec une table d'association."
    )

    # Slide 9: SQL DDL
    add_slide_with_content(
        prs,
        "8. Le Langage SQL : Définition des Données",
        [
            {"title": "Architecture Client-Serveur", "bullets": ["Le client (ex: DBeaver) envoie des requêtes SQL au serveur (ex: PostgreSQL)."]},
            {"title": "Création de Tables (DDL)", "highlight": True, "bullets": [
                "Syntaxe : CREATE TABLE nom (colonne1 type, colonne2 type);",
                "Contraintes : NOT NULL, PRIMARY KEY, UNIQUE, CHECK(condition)."
            ]},
            {"title": "Modification", "bullets": ["ALTER TABLE : Ajouter, supprimer des colonnes ou des contraintes.", "DROP TABLE : Supprimer une table."]}
        ],
        "[VISUEL KADEA]\nCapture d'écran de l'interface DBeaver connectée à une base PostgreSQL."
    )

    # Slide 10: SQL DML
    add_slide_with_content(
        prs,
        "9. Le Langage SQL : Manipulation des Données",
        [
            {"title": "Insertion", "bullets": ["INSERT INTO table (col1, col2) VALUES (val1, val2);"]},
            {"title": "Modification", "bullets": ["UPDATE table SET col1 = val1 WHERE condition;"]},
            {"title": "Suppression", "bullets": ["DELETE FROM table WHERE condition;"]},
            {"title": "Qualité de Données (LinkedIn Learning)", "highlight": True, "bullets": [
                "Le SQL est indispensable pour le nettoyage (Data Wrangling).",
                "Permet de gérer massivement les formats mixtes ou les données manquantes (NULL)."
            ]}
        ],
        "[VISUEL KADEA]\nCode snippet : Une requête UPDATE corrigeant une faute de frappe dans un dataset sale."
    )

    # Slide 11: SQL DQL (Sélection)
    add_slide_with_content(
        prs,
        "10. Le Langage SQL : Requêtes d'Analyse (1/2)",
        [
            {"title": "Projection et Sélection", "bullets": [
                "SELECT col1, col2 FROM table WHERE condition;",
                "Utilisation de l'opérateur LIKE avec '%' pour les recherches textuelles."
            ]},
            {"title": "La Jointure (JOIN)", "highlight": True, "bullets": [
                "Pour relier les données de plusieurs tables.",
                "SELECT * FROM tableA JOIN tableB ON tableA.id = tableB.fk_id;"
            ]}
        ],
        "[VISUEL KADEA]\nUn schéma d'ensembles de Venn montrant le fonctionnement d'une Jointure (l'intersection)."
    )

    # Slide 12: SQL DQL (Agrégation)
    add_slide_with_content(
        prs,
        "11. Le Langage SQL : Requêtes d'Analyse (2/2)",
        [
            {"title": "Fonctions d'Agrégation", "bullets": [
                "COUNT(), AVG(), SUM(), MIN(), MAX().",
                "Le regroupement via GROUP BY pour analyser par catégorie."
            ]},
            {"title": "Fonctions de Fenêtrage (LinkedIn Learning)", "highlight": True, "bullets": [
                "Pour l'analyse statistique avancée sur PostgreSQL.",
                "Syntaxe : OVER(), PARTITION BY().",
                "Permet de calculer des totaux cumulés ou des rangs sans réduire le nombre de lignes (contrairement à GROUP BY)."
            ]}
        ],
        "[VISUEL KADEA]\nTableau de bord minimaliste montrant des résultats agrégés par région."
    )

    prs.save('Cours_Complet_BDD_et_SQL.pptx')
    print("Présentation complète en français générée avec succès.")

if __name__ == "__main__":
    create_full_presentation()
