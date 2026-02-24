from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

# --- Configuration ---
FILENAME = "Atelier_Technique_Data.pptx"
TITLE_FONT = "Arial"
BODY_FONT = "Calibri"
COLOR_BLUE = RGBColor(0, 51, 153) # Deep Corporate Blue
COLOR_WHITE = RGBColor(255, 255, 255)
COLOR_ORANGE = RGBColor(255, 140, 0) # Accent Color
COLOR_GRAY = RGBColor(80, 80, 80)
COLOR_LIGHT_GRAY = RGBColor(240, 240, 240)

# --- Helper Functions ---

def create_slide(prs, layout_index=1):
    try:
        slide_layout = prs.slide_layouts[layout_index]
        slide = prs.slides.add_slide(slide_layout)
        return slide
    except IndexError:
        slide_layout = prs.slide_layouts[0]
        slide = prs.slides.add_slide(slide_layout)
        return slide

def add_title(slide, text, font_size=32, color=COLOR_BLUE):
    if slide.shapes.title:
        title = slide.shapes.title
        title.text = text
        title.text_frame.paragraphs[0].font.name = TITLE_FONT
        title.text_frame.paragraphs[0].font.size = Pt(font_size)
        title.text_frame.paragraphs[0].font.color.rgb = color
        title.text_frame.paragraphs[0].alignment = PP_ALIGN.LEFT

def add_content_text(slide, points, left=Inches(0.5), top=Inches(1.5), width=Inches(9), height=Inches(5)):
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    for i, point in enumerate(points):
        p = tf.add_paragraph()
        p.text = point
        p.font.name = BODY_FONT
        p.font.size = Pt(20)
        p.font.color.rgb = COLOR_GRAY
        p.space_after = Pt(10)
        p.level = 0

def draw_shape(slide, shape_type, left, top, width, height, color=COLOR_BLUE, text=""):
    shape = slide.shapes.add_shape(shape_type, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.color.rgb = COLOR_GRAY
    if text:
        shape.text = text
        shape.text_frame.paragraphs[0].font.color.rgb = COLOR_WHITE
        shape.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    return shape

def add_code_block(slide, code_text, left, top, width, height):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = COLOR_LIGHT_GRAY
    shape.line.color.rgb = COLOR_GRAY

    tf = shape.text_frame
    p = tf.paragraphs[0]
    p.text = code_text
    p.font.name = "Consolas"
    p.font.size = Pt(14)
    p.font.color.rgb = RGBColor(0, 0, 0)
    p.alignment = PP_ALIGN.LEFT

def add_placeholder_image(slide, description, left, top, width, height):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor(200, 200, 200) # Gray placeholder
    shape.line.dash_style = 4 # Dashed line

    tf = shape.text_frame
    p = tf.paragraphs[0]
    p.text = f"[Capture d'écran : {description}]"
    p.font.color.rgb = RGBColor(100, 100, 100)
    p.alignment = PP_ALIGN.CENTER

# --- Main Script ---

def generate_presentation():
    prs = Presentation()

    # 1. Slide de Titre
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    title = slide.shapes.title
    subtitle = slide.placeholders[1]

    title.text = "Atelier Technique : Data Expert & SQL"
    title.text_frame.paragraphs[0].font.name = TITLE_FONT
    title.text_frame.paragraphs[0].font.size = Pt(40)
    title.text_frame.paragraphs[0].font.color.rgb = COLOR_BLUE

    subtitle.text = "Excel Expert (VBA, Power Query) & CRUD SQL/NoSQL"
    subtitle.text_frame.paragraphs[0].font.name = BODY_FONT
    subtitle.text_frame.paragraphs[0].font.size = Pt(24)
    subtitle.text_frame.paragraphs[0].font.color.rgb = COLOR_ORANGE

    # 2. Agenda Expert
    slide = create_slide(prs, 5)
    add_title(slide, "Au Programme (3h30)")
    add_content_text(slide, [
        "1. Tri & Filtres Avancés (Multi-critères, Slicers)",
        "2. Croiser les données : VLOOKUP vs Power Query Merge",
        "3. Automatisation : Intro VBA & Google AppScript",
        "4. Le langage M (Power Query) pour les cas complexes",
        "5. SQL Jointures (JOIN) vs NoSQL Imbrication",
        "6. Atelier Pratique : Gestion de Clients Telco"
    ])

    # 3. Excel : Tri Expert
    slide = create_slide(prs, 5)
    add_title(slide, "Tri Multi-Niveaux")
    add_content_text(slide, [
        "Le Tri simple (A-Z) ne suffit pas.",
        "Cas réel : Trier par Région, PUIS par Offre, PUIS par Nom.",
        "Outil : Données > Trier (Custom Sort).",
        "Astuce : On peut trier par couleur ou icône !"
    ], width=Inches(6))

    draw_shape(slide, MSO_SHAPE.DOWN_ARROW, Inches(7), Inches(2), Inches(1), Inches(3), COLOR_BLUE, "1. Region\n2. Offre\n3. Nom")

    # 4. Filtres Avancés
    slide = create_slide(prs, 5)
    add_title(slide, "Filtres Avancés (Criteria Range)")
    add_content_text(slide, [
        "Pour des conditions complexes (ET / OU).",
        "Nécessite une 'Zone de critères' séparée.",
        "Permet d'extraire les données vers une autre feuille.",
        "Alternative moderne : Fonction FILTER() (Office 365)."
    ])

    add_placeholder_image(slide, "Boîte de dialogue 'Filtre Avancé' avec Plage et Critères", Inches(1), Inches(4.5), Inches(8), Inches(2.5))

    # 5. Power Query : Au-delà du clic
    slide = create_slide(prs, 5)
    add_title(slide, "Power Query : Langage M")
    add_content_text(slide, [
        "Derrière chaque clic, Power Query écrit du code M.",
        "On peut l'éditer pour des calculs complexes.",
        "Exemple : Date.FromText([DateString])",
        "Exemple : Text.Select([Phone], {'0'..'9'})"
    ], width=Inches(5))

    add_code_block(slide, "let\n  Source = Excel.CurrentWorkbook(){[Name='Table1']}[Content],\n  ChangedType = Table.TransformColumnTypes(Source,{{'Date', type date}}),\n  FilteredRows = Table.SelectRows(ChangedType, each ([Montant] > 100))\nin\n  FilteredRows", Inches(5.5), Inches(2), Inches(4), Inches(4))

    # 6. VBA vs AppScript
    slide = create_slide(prs, 5)
    add_title(slide, "Automatisation : Le Duel")

    draw_shape(slide, MSO_SHAPE.RECTANGLE, Inches(0.5), Inches(2), Inches(4), Inches(4), COLOR_BLUE, "VBA (Excel)\n\nAncien mais robuste.\nLocal (Desktop).\nLangage : Visual Basic.\nIdéal pour formulaires UserForm.")
    draw_shape(slide, MSO_SHAPE.RECTANGLE, Inches(5), Inches(2), Inches(4), Inches(4), COLOR_ORANGE, "AppScript (Google)\n\nModerne (Cloud).\nWeb (Browser).\nLangage : JavaScript.\nIdéal pour connecteurs (Gmail, Drive).")

    # 7. Google Sheets : La fonction QUERY
    slide = create_slide(prs, 5)
    add_title(slide, "Google Sheets : QUERY()")
    add_content_text(slide, [
        "La puissance du SQL directement dans le tableur.",
        "Plus flexible que les TCD (Tableaux Croisés Dynamiques).",
        "Syntaxe : =QUERY(Données; 'SELECT A, SUM(B) GROUP BY A'; 1)"
    ], width=Inches(9))

    add_code_block(slide, "=QUERY(A1:D100; \n  \"SELECT A, AVG(C) \n   WHERE D = 'Actif' \n   GROUP BY A \n   LABEL AVG(C) 'Moyenne'\")", Inches(1), Inches(4), Inches(8), Inches(2))

    # 8. SQL Avancé : JOIN
    slide = create_slide(prs, 5)
    add_title(slide, "SQL : Les Jointures (JOIN)")

    # Visual representations of tables
    t1 = draw_shape(slide, MSO_SHAPE.RECTANGLE, Inches(1), Inches(3), Inches(2), Inches(1.5), COLOR_BLUE, "Table Clients\n(ID_Offre)")
    t2 = draw_shape(slide, MSO_SHAPE.RECTANGLE, Inches(5), Inches(3), Inches(2), Inches(1.5), COLOR_GRAY, "Table Offres\n(ID_Offre)")

    # Connector
    slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(3), Inches(3.75), Inches(5), Inches(3.75))

    add_content_text(slide, ["INNER JOIN : Prend l'intersection des deux tables basées sur une clé commune."], Inches(1), Inches(5), Inches(8))

    # 9. NoSQL : Imbrication
    slide = create_slide(prs, 5)
    add_title(slide, "NoSQL : Pas de JOIN !")
    add_content_text(slide, [
        "En NoSQL, on évite les relations complexes.",
        "On préfère imbriquer les données (Embedding).",
        "Avantage : Lecture ultra-rapide (1 seul accès disque).",
        "Inconvénient : Duplication de données."
    ])

    add_code_block(slide, "{\n  'nom': 'Jean',\n  'adresse': { \n    'rue': 'Main St', \n    'ville': 'Paris' \n  }\n}", Inches(6), Inches(2.5), Inches(3), Inches(3))

    # 10. Conclusion & Ressources
    slide = create_slide(prs, 5)
    add_title(slide, "Ressources & Outils")
    add_content_text(slide, [
        "DB Fiddle (SQL Training)",
        "MongoDB Atlas (NoSQL Cloud)",
        "Stack Overflow (VBA/Excel Help)",
        "Documentation Google AppScript"
    ])

    prs.save(FILENAME)
    print(f"Presentation saved as {FILENAME}")

if __name__ == "__main__":
    generate_presentation()
