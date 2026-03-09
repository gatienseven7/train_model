from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE

def create_presentation():
    prs = Presentation()

    # Slides content based on user input
    slides_data = [
        {
            "title": "Semaine 5 - Transition d'Excel vers la Business Intelligence Moderne.",
            "content": [
                "L'objectif de cette semaine est de faire le pont entre le tableur classique et les outils décisionnels interactifs.",
                "Vous découvrirez les concepts de la BI, la puissance de l'interactivité et réaliserez votre première visualisation guidée."
            ]
        },
        {
            "title": "Prendre des décisions basées sur des faits, un concept théorisé dès 1856.",
            "content": [
                "En 1856, Richard Miller Devens utilise le terme \"Business Intelligence\" pour décrire un banquier qui a devancé ses concurrents en analysant des données empiriques plutôt que de se fier à son instinct.",
                "En 1958, Hans Peter Luhn (IBM) théorise la BI comme un système automatique capable d'appréhender les liens entre les faits et de les diffuser aux décideurs."
            ]
        },
        {
            "title": "Excel est excellent pour la saisie, mais limité pour le traitement massif et le partage.",
            "content": [
                "Limite de volume : Excel sature autour d'un million de lignes, là où un outil comme Power BI peut analyser des millions, voire des milliards de lignes en direct.",
                "Sécurité et Partage : Fini l'envoi de fichiers lourds par email ; la BI moderne utilise le Cloud (comme Power BI Service) pour un partage sécurisé avec des vues filtrées par utilisateur (RLS).",
                "Maintenance : L'actualisation manuelle via des macros est remplacée par une actualisation automatique et planifiée."
            ]
        },
        {
            "title": "La BI transforme les données opérationnelles en informations stratégiques.",
            "content": [
                "Les systèmes opérationnels (ex: RH, ventes) maintiennent l'entreprise en vie au quotidien.",
                "La BI, elle, ne génère pas de profit direct mais aide à la stratégie via 4 phases : la Collecte (via des outils ETL), l'Intégration (dans un Data Warehouse), la Distribution (Data Marts) et la Restitution (Cubes OLAP / Tableaux de bord)."
            ]
        },
        {
            "title": "Deux grandes écoles pour structurer un Entrepôt de Données (Data Warehouse).",
            "content": [
                "L'approche Top-Down (Bill Inmon) : Cartographie globale et données normalisées (schéma en flocons). Garantit une grande cohérence, mais est très complexe et coûteuse à mettre en place.",
                "L'approche Bottom-Up (Ralph Kimball) : Orientée processus métiers (Data Marts dénormalisés en étoile). Très rapide à mettre en œuvre, mais nécessite une maintenance rigoureuse pour éviter les incohérences."
            ]
        },
        {
            "title": "La BI \"en Libre-Service\" a démocratisé l'accès à la donnée.",
            "content": [
                "Dans la BI traditionnelle, seuls les experts IT généraient les rapports, créant des files d'attente et de la frustration pour les dirigeants.",
                "Aujourd'hui, la BI moderne (Power BI, Looker Studio) permet à des profils non-informaticiens d'explorer librement la donnée et de créer eux-mêmes leurs visuels."
            ]
        },
        {
            "title": "Appliquer le Storytelling (Framework McKinsey) pour convaincre.",
            "content": [
                "Le but d'un tableau de bord n'est pas d'exister, mais de faire prendre des décisions à la direction.",
                "Utilisez la structure SCR : Situation (le contexte des données), Complication (le problème identifié, ex: hausse des annulations), Résolution (les recommandations).",
                "Appliquez la structure Dot-Dash : un titre d'action qui donne la conclusion, soutenu par des graphiques qui prouvent cette conclusion."
            ]
        },
        {
            "title": "La clarté avant tout : ne recréez pas Excel dans un outil BI.",
            "content": [
                "Le piège classique est de vouloir afficher 15 graphiques statiques sur une même page.",
                "Privilégiez l'interactivité : un seul visuel bien construit peut remplacer plusieurs graphiques statiques grâce aux \"paramètres de champ\" permettant à l'utilisateur de choisir ce qu'il veut observer (ex: chiffre d'affaires vs taux d'annulation)."
            ]
        },
        {
            "title": "Faire parler la donnée d'un simple clic.",
            "content": [
                "Contrairement à Excel, cliquer sur un élément d'un graphique (ex: une région) filtre instantanément tous les autres visuels de la page.",
                "Les outils modernes permettent même d'interroger les données en langage naturel (ex: \"Quel est le produit le plus vendu ?\") comme sur un moteur de recherche."
            ]
        },
        {
            "title": "Le rôle du Consultant BI est de créer de la valeur pour chaque département.",
            "content": [
                "Logistique : Anticiper les besoins, éviter les ruptures de stocks et identifier les goulots d'étranglement.",
                "Marketing & Ventes : Prédire les comportements des clients, cibler les campagnes et adapter les stratégies par région.",
                "Ressources Humaines : Optimiser la présence, suivre le taux de rotation et les besoins en formation."
            ]
        },
        {
            "title": "À vous de jouer : La magie du glisser-déposer.",
            "content": [
                "Étape 1 : Ouverture de l'outil (Power BI Desktop ou Looker Studio) et connexion au fichier Excel travaillé en Semaine 4.",
                "Étape 2 : Refaire le graphique de la semaine dernière en quelques secondes par un simple glisser-déposer, et tester les filtres interactifs."
            ]
        },
        {
            "title": "Du Cloud à l'Intelligence Artificielle : d'une analyse descriptive à prédictive.",
            "content": [
                "La migration massive vers le Cloud (AWS, Microsoft, Google) facilite le stockage, réduit les coûts et améliore l'accessibilité.",
                "L'IA et le Machine Learning viennent compléter la BI : la BI décrit le passé et le présent, l'IA prédit les tendances futures et propose des actions automatisées."
            ]
        }
    ]

    title_slide_layout = prs.slide_layouts[0]
    bullet_slide_layout = prs.slide_layouts[1]

    # Add main title slide
    slide = prs.slides.add_slide(title_slide_layout)
    title = slide.shapes.title
    subtitle = slide.placeholders[1]
    title.text = "De la donnée brute à la décision stratégique"
    subtitle.text = "Semaine 5\nTransition d'Excel vers la Business Intelligence Moderne"

    # Decorate title slide
    # Corporate Red rectangle on the left
    left = Inches(0)
    top = Inches(0)
    width = Inches(0.5)
    height = prs.slide_height
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor(200, 0, 0) # Corporate Red
    shape.line.fill.background()

    for slide_data in slides_data:
        slide = prs.slides.add_slide(bullet_slide_layout)
        shapes = slide.shapes

        title_shape = shapes.title
        body_shape = shapes.placeholders[1]

        title_shape.text = slide_data["title"]

        # Style Title
        title_shape.text_frame.paragraphs[0].font.color.rgb = RGBColor(200, 0, 0)
        title_shape.text_frame.paragraphs[0].font.bold = True

        tf = body_shape.text_frame
        tf.text = slide_data["content"][0]

        for point in slide_data["content"][1:]:
            p = tf.add_paragraph()
            p.text = point
            p.level = 0

        # Decorate content slide
        # Corporate Red rectangle on the top left
        left = Inches(0)
        top = Inches(0)
        width = Inches(2)
        height = Inches(0.1)
        shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
        shape.fill.solid()
        shape.fill.fore_color.rgb = RGBColor(200, 0, 0) # Corporate Red
        shape.line.fill.background()

        # Small circle at the bottom right
        left = prs.slide_width - Inches(0.5)
        top = prs.slide_height - Inches(0.5)
        width = Inches(0.3)
        height = Inches(0.3)
        shape = slide.shapes.add_shape(MSO_SHAPE.OVAL, left, top, width, height)
        shape.fill.solid()
        shape.fill.fore_color.rgb = RGBColor(200, 0, 0) # Corporate Red
        shape.line.fill.background()

    prs.save('Semaine_5_Transition_BI.pptx')

if __name__ == '__main__':
    create_presentation()
