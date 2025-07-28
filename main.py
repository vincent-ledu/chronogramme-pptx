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
import re
from pptx.util import Inches

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

if len(sys.argv) < 2:
    print("Usage: python generate_slide.py <fichier_excel>")
    sys.exit(1)

excel_path = sys.argv[1]
if not os.path.exists(excel_path):
    print(f"Fichier introuvable: {excel_path}")
    sys.exit(1)

col_produit = config["colonne_produit"]
col_solution = config["colonne_solution"]
col_planif = config["colonne_planification"]
col_tribue = config["colonne_tribue"]
col_squad = config["colonne_squad"]
col_kube = config["colonne_full_kube"]
col_z = config["colonne_full_z"]
col_mosart = config["colonne_mosart"]
col_critique = config["colonne_critique"]
col_validate = config["colonne_validate"]
col_decom = config["colonne_decom"]

icone_kube = config.get("icone_kube", "kubernetes.png")
icone_critique = config.get("icone_critique", "eclair.png")
icone_check = config.get("icone_check", "check.png")
icone_z = config.get("icone_z", "z.png")

def hex_to_rgbcolor(hex_code):
    hex_code = hex_code.lstrip("#")
    return RGBColor(int(hex_code[0:2], 16), int(hex_code[2:4], 16), int(hex_code[4:6], 16))


# Remplace squad_to_color() par ce dictionnaire mappé
#squads_uniques = sorted(df[col_squad].dropna().unique())
squad_colors_raw = config.get("squad_colors", {})
squad_color_map = {
    squad: hex_to_rgbcolor(hex_code)
    for squad, hex_code in squad_colors_raw.items()
}

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

df = pd.read_excel(excel_path)
tribues = df[col_tribue].dropna().unique()
expected_columns = {col_produit, col_solution, col_planif, col_squad}
if not expected_columns.issubset(df.columns):
    print(f"Le fichier Excel doit contenir : {expected_columns}")
    sys.exit(1)


# Cloner les lignes critiques à l’année suivante
def dupliquer_ligne_critique(df):
    lignes_critiques = df[df[col_critique].astype(str).str.strip().str.lower() == "oui"].copy()
    def trimestre_plus_1(val):
        match = re.match(r"(T[1-4])/(\d{4})", str(val).strip())
        if match:
            tri, annee = match.groups()
            return f"{tri}/{int(annee)+1}"
        return val
    lignes_critiques[col_planif] = lignes_critiques[col_planif].apply(trimestre_plus_1)
    return pd.concat([df, lignes_critiques], ignore_index=True)


for tribue in tribues:
    df_tribue = df[df[col_tribue] == tribue].copy()
    df_tribue = dupliquer_ligne_critique(df_tribue)

    # ➤ ici, appliquer :
    # - duplication critique (df_tribue = dupliquer_ligne_critique(df_tribue, ...))
    # planifier à n+1 si critique
    # - gestion planification
    # - génération du slide

    pptx_path = "exemple_chronogramme.pptx"
    prs = Presentation(pptx_path)
    slide = prs.slides[0]

    # Slide alternatif pour planification absente ou invalide
    slide_invalide = prs.slides[1]
    positions_valides = set(positions.keys())

    height = Inches(0.22)
    width = Inches(1.0)
    base_top = Inches(1.5)
    vspace = Inches(0.27)
    ligne_par_trimestre = {}
    lignes_inconnues = 0


    # ... génération du contenu comme avant (boîtes, couleurs, icônes, etc.) ...
    # Génération des boîtes principales
    for _, row in df_tribue.iterrows():
        produit = str(row.get(col_produit, ""))
        solution = str(row.get(col_solution, ""))
        trimestre = str(row.get(col_planif, "")).strip()
        squad = str(row.get(col_squad, "")).strip()
        full_kube = str(row.get(col_kube, "")).strip().lower() == "oui"
        mosart = str(row.get(col_mosart, "")).strip().lower().startswith("lot")
        decom = str(row.get(col_decom, "")).strip().lower() == "oui"
        critique = str(row.get(col_critique, "")).strip().lower() == "oui"
        validate = str(row.get(col_validate, "")).strip().lower() == "oui"
        full_z = str(row.get(col_z, "")).strip().lower() == "oui"

        color = squad_color_map.get(squad, RGBColor(160, 160, 160))

        if trimestre not in positions:
            print(f"⚠️ {produit}-{solution} : Trimestre inconnu ignoré : {trimestre}")

        if trimestre in positions_valides:
            left = positions[trimestre]
            ligne = ligne_par_trimestre.get(trimestre, 0)
            top = base_top + vspace * ligne
            ligne_par_trimestre[trimestre] = ligne + 1
            target_slide = slide
        else:
            left = Inches(1.0)
            top = Inches(1.0 + lignes_inconnues * 0.4)
            lignes_inconnues += 1
            target_slide = slide_invalide

        textbox = target_slide.shapes.add_shape(
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
        fill.fore_color.rgb = color

        line = textbox.line
        line.color.rgb = RGBColor(0, 0, 0)
        # 🔶 Si mosart startwith = "lot" → contour orange
        if mosart:
            line.width = Pt(2.5)
            line.color.rgb = RGBColor(255, 102, 0)
        
        # Si decom = "oui" → contour gris
        if decom:
            textbox.line.width = Pt(2.5)
            textbox.line.color.rgb = RGBColor(80, 80, 80)

        # Texte centré blanc
        text_frame = textbox.text_frame
        for paragraph in text_frame.paragraphs:
            paragraph.alignment = PP_ALIGN.CENTER
            for run in paragraph.runs:
                run.font.size = Pt(8)
                run.font.bold = True
                run.font.color.rgb = RGBColor(255, 255, 255)

        # 🐳 Icône Kubernetes à gauche si full kube = oui
        if full_kube:
            target_slide.shapes.add_picture(
                icone_kube,
                left=left - Inches(0.15),
                top=top,
                width=Inches(0.22),
                height=Inches(0.22)
            )

        # Z Icône Z à gauche si full Z = oui
        if full_z:
            target_slide.shapes.add_picture(
                icone_z,
                left=left - Inches(0.15),
                top=top,
                width=Inches(0.22),
                height=Inches(0.22)
            )

        # ⚡ Icône éclair rouge si critique = oui
        if critique:
            target_slide.shapes.add_picture(
                icone_critique,
                left=left + width - Inches(0.15),
                top=top + Inches(0.02),
                width=Inches(0.22),
                height=Inches(0.22)
            )

        # Icône check à au milieu si validé = oui
        if validate:
            target_slide.shapes.add_picture(
                icone_check,
                left=left + Inches(0.15),
                top=top,
                width=Inches(0.22),
                height=Inches(0.22)
            )


    # ✅ Ajouter la légende en bas à gauche
    squads_uniques = sorted(df_tribue[col_squad].dropna().unique())
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
        box.fill.solid()
        box.fill.fore_color.rgb = squad_color_map[squad]
        box.line.color.rgb = RGBColor(0, 0, 0)

        text_frame = box.text_frame
        for paragraph in text_frame.paragraphs:
            for run in paragraph.runs:
                run.font.size = Pt(10)
                run.font.bold = True
                run.font.color.rgb = RGBColor(255, 255, 255)
    
    
    nom_fichier = f"chronogramme_{tribue.replace(' ', '_')}.pptx"
    prs.save(nom_fichier)
    print(f"✅ Fichier généré pour {tribue} : {nom_fichier}")
