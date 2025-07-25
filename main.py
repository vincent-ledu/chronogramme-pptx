import sys
import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
import os
from pptx.dml.color import RGBColor
import json
import argparse

parser = argparse.ArgumentParser(description="Génère un slide chronogramme à partir d’un fichier Excel.")
parser.add_argument("excel_file", help="Chemin vers le fichier Excel contenant les données.")
parser.add_argument("--config", default="config.json", help="Chemin vers le fichier JSON de configuration (défaut: config.json).")

args = parser.parse_args()
excel_path = args.excel_file
config_path = args.config

if not os.path.exists(config_path):
    print(f"❌ Fichier de configuration introuvable : {config_path}")
    sys.exit(1)

with open(config_path, "r", encoding="utf-8") as f:
    config = json.load(f)

col_produit = config["colonne_produit"]
col_solution = config["colonne_solution"]
col_planif = config["colonne_planification"]
col_squad = config["colonne_squad"]
col_kube = config["colonne_full_kube"]
col_mosart = config["colonne_mosart"]
col_critique = config["colonne_critique"]

def hex_to_rgbcolor(hex_code):
    hex_code = hex_code.lstrip("#")
    return RGBColor(int(hex_code[0:2], 16), int(hex_code[2:4], 16), int(hex_code[4:6], 16))

if len(sys.argv) < 2:
    print("Usage: python generate_slide.py <fichier_excel>")
    sys.exit(1)

excel_path = sys.argv[1]
if not os.path.exists(excel_path):
    print(f"Fichier introuvable: {excel_path}")
    sys.exit(1)

df = pd.read_excel(excel_path)

# Remplace squad_to_color() par ce dictionnaire mappé
squads_uniques = sorted(df[col_squad].dropna().unique())
squad_colors_raw = config.get("squad_colors", {})
squad_color_map = {
    squad: hex_to_rgbcolor(hex_code)
    for squad, hex_code in squad_colors_raw.items()
}
expected_columns = {col_produit, col_solution, col_planif, col_squad}
if not expected_columns.issubset(df.columns):
    print(f"Le fichier Excel doit contenir : {expected_columns}")
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

height = Inches(0.22)
width = Inches(1.0)
base_top = Inches(1.5)
vspace = Inches(0.27)
ligne_par_trimestre = {}

# Génération des boîtes principales
for _, row in df.iterrows():
    produit = str(row[col_produit])
    solution = str(row[col_solution])
    trimestre = str(row[col_planif]).strip()
    squad = str(row[col_squad]).strip()
    full_kube = row.get(col_kube, 0)
    mosart = row.get(col_mosart, 0)
    critique = str(row.get(col_critique, "")).strip().lower()


    if trimestre not in positions:
        print(f"⚠️ {produit}-{solution} : Trimestre inconnu ignoré : {trimestre}")
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
    color = squad_color_map.get(squad, RGBColor(200, 200, 200))  # Gris si inconnu
    fill.fore_color.rgb = color

    # 🔶 Si mosart != "" → contour orange
    if str(row.get(col_mosart, "")).strip().startswith("lot"):
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

    # 🐳 Icône Kubernetes à gauche si full kube = oui
    if str(row.get(col_kube, "")).strip().lower() == "oui":
        slide.shapes.add_picture(
            "kubernetes.png",
            left=left - Inches(0.15),
            top=top,
            width=Inches(0.22),
            height=Inches(0.22)
        )

    # ⚡ Icône éclair rouge si critique = oui
    if str(row.get(col_critique, "")).strip().lower() == "oui":
        slide.shapes.add_picture(
            "eclair.png",
            left=left + width - Inches(0.15),
            top=top + Inches(0.02),
            width=Inches(0.22),
            height=Inches(0.22)
        )

    text_frame = textbox.text_frame
    for paragraph in text_frame.paragraphs:
        paragraph.alignment = PP_ALIGN.CENTER
        for run in paragraph.runs:
            run.font.size = Pt(8)
            run.font.bold = True
            run.font.color.rgb = RGBColor(255, 255, 255)

# ✅ Ajouter la légende en bas à gauche
squads_uniques = sorted(df[col_squad].dropna().unique())
legend_left = Inches(0.5)
legend_top_base = Inches(5.5)
legend_height = Inches(0.3)
legend_width = Inches(3.0)
legend_vspace = Inches(0.32)

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
