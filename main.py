import sys
import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
import os
import hashlib

from pptx.dml.color import RGBColor

RAINBOW_COLORS = [
    RGBColor(255, 0, 0),
    RGBColor(255, 127, 0),
    RGBColor(255, 255, 0),
    RGBColor(127, 255, 0),
    RGBColor(0, 255, 0),
    RGBColor(0, 255, 127),
    RGBColor(0, 255, 255),
    RGBColor(0, 127, 255),
    RGBColor(0, 0, 255),
    RGBColor(139, 0, 255),
]

if len(sys.argv) < 2:
    print("Usage: python generate_slide.py <fichier_excel>")
    sys.exit(1)

excel_path = sys.argv[1]
if not os.path.exists(excel_path):
    print(f"Fichier introuvable: {excel_path}")
    sys.exit(1)

df = pd.read_excel(excel_path)

# Remplace squad_to_color() par ce dictionnaire mappé
squads_uniques = sorted(df["squad"].dropna().unique())
squad_color_map = {
    squad: RAINBOW_COLORS[i % len(RAINBOW_COLORS)]
    for i, squad in enumerate(squads_uniques)
}

expected_columns = {"Produit", "solution", "planification", "squad"}
if not expected_columns.issubset(df.columns):
    print("Le fichier Excel doit contenir : Produit, solution, planification, squad")
    sys.exit(1)

positions = {
    "T1/2025": Inches(1.0),
    "T2/2025": Inches(2.5),
    "T3/2025": Inches(4.0),
    "T4/2025": Inches(5.5),
    "T1/2026": Inches(7.0),
    "T2/2026": Inches(8.5),
    "T3/2026": Inches(10.0),
    "T4/2026": Inches(11.5),
}

pptx_path = "exemple_chronogramme.pptx"
prs = Presentation(pptx_path)
slide = prs.slides[0]

height = Inches(0.3)
width = Inches(1.0)
base_top = Inches(2.0)
vspace = Inches(0.4)
ligne_par_trimestre = {}

# Génération des boîtes principales
for _, row in df.iterrows():
    produit = str(row["Produit"])
    solution = str(row["solution"])
    trimestre = str(row["planification"]).strip()
    squad = str(row["squad"]).strip()

    if trimestre not in positions:
        print(f"⚠️ Trimestre inconnu ignoré : {trimestre}")
        continue

    left = positions[trimestre]
    ligne = ligne_par_trimestre.get(trimestre, 0)
    top = base_top + vspace * ligne
    ligne_par_trimestre[trimestre] = ligne + 1

    textbox = slide.shapes.add_shape(
        autoshape_type_id=1,
        left=left,
        top=top,
        width=width,
        height=height
    )
    textbox.text = f"{produit}-{solution}"

    # Appliquer style
    fill = textbox.fill
    fill.solid()
    fill.fore_color.rgb = squad_color_map[squad]

    # 🔶 Si mosart = 1 → contour orange
    if row.get("mosart", 0) == 1:
        textbox.line.width = Pt(2.5)
        textbox.line.color.rgb = RGBColor(255, 102, 0)

    else:
        textbox.line.color.rgb = RGBColor(0, 0, 0)

    # Texte centré blanc
    text_frame = textbox.text_frame
    for paragraph in text_frame.paragraphs:
        paragraph.alignment = PP_ALIGN.CENTER
        for run in paragraph.runs:
            run.font.size = Pt(8)
            run.font.bold = True
            run.font.color.rgb = RGBColor(255, 255, 255)

    # 🐳 Icône Kubernetes à gauche si full kube = 1
    if row.get("full kube", 0) == 1:
        slide.shapes.add_picture(
            "kubernetes.png",
            left=left - Inches(0.15),
            top=top,
            width=Inches(0.3),
            height=Inches(0.3)
        )

    # ⚡ Icône éclair rouge si critique = oui
    if str(row.get("critique", "")).strip().lower() == "oui":
        slide.shapes.add_picture(
            "eclair.png",
            left=left + width - Inches(0.15),
            top=top + Inches(0.02),
            width=Inches(0.3),
            height=Inches(0.3)
        )

    text_frame = textbox.text_frame
    for paragraph in text_frame.paragraphs:
        paragraph.alignment = PP_ALIGN.CENTER
        for run in paragraph.runs:
            run.font.size = Pt(8)
            run.font.bold = True
            run.font.color.rgb = RGBColor(255, 255, 255)

# ✅ Ajouter la légende en bas à gauche
squads_uniques = sorted(df["squad"].dropna().unique())
legend_left = Inches(0.5)
legend_top_base = Inches(6.0)
legend_height = Inches(0.4)
legend_width = Inches(3.0)
legend_vspace = Inches(0.45)

for i, squad in enumerate(squads_uniques):
    top = legend_top_base + i * legend_vspace
    box = slide.shapes.add_shape(
        autoshape_type_id=1,
        left=legend_left,
        top=top,
        width=legend_width,
        height=legend_height
    )
    box.text = squad

    fill = box.fill
    fill.solid()
    fill.fore_color.rgb = squad_color_map[squad]
    box.line.color.rgb = RGBColor(0, 0, 0)

    text_frame = box.text_frame
    for paragraph in text_frame.paragraphs:
        for run in paragraph.runs:
            run.font.size = Pt(10)
            run.font.bold = True
            run.font.color.rgb = RGBColor(255, 255, 255)

# ✅ Enregistrement
prs.save("chronogramme_genere.pptx")
print("✅ Fichier généré avec légende : chronogramme_genere.pptx")
