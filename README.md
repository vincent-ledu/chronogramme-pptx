# 📊 Générateur de Chronogrammes PowerPoint

Ce projet permet de générer automatiquement des slides PowerPoint chronogrammes à partir d’un fichier Excel structuré. Chaque slide représente une _tribue_, avec les produits, solutions et leurs états d'avancement visuellement représentés par des boîtes et des icônes.

---

## 🚀 Fonctionnalités

- ✅ Génération d'un fichier `.pptx` par tribue
- 🧱 Regroupement des lignes par `Produit` + `Solution` + `Type`
- 🔁 Duplication des lignes critiques sur l’année suivante
- 🎨 Couleurs automatiques par `Squad`
- 🔍 Filtrage des données `NR` / `NA` dans la colonne `réalisé`
- 🧼 Nettoyage et harmonisation des données (texte et booléens)
- 🧾 Légende des squads
- 🎯 Icônes conditionnelles :
  - 🐳 Kubernetes
  - ⚡ Critique
  - 🪵 Décommissionné
  - ✅ Validé ou tout réalisé
  - 🔨 Reconstruction / Restauration / Resynchro
  - 🟣 Full Z

---

## 📁 Structure des fichiers

```
.
├── generate_chronogrammes.py      # Script principal
├── config.json                    # Fichier de configuration
├── planning.xlsx                  # Fichier Excel source
├── exemple_chronogramme.pptx     # Modèle PowerPoint
├── /output                        # Dossier de sortie des fichiers générés
└── /icones                        # Dossier des fichiers PNG des icônes
```

---

## 🧾 Configuration (`config.json`)

Extrait d'exemple :

```json
{
  "colonne_produit": "Produit",
  "colonne_solution": "solution",
  "colonne_planification": "planification",
  "colonne_tribue": "tribue",
  "colonne_squad": "squad",
  "colonne_full_kube": "full kube",
  "colonne_full_z": "full z",
  "colonne_mosart": "mosart",
  "colonne_critique": "critique",
  "colonne_validate": "validé",
  "colonne_decom": "decom",
  "colonne_type": "type",
  "colonne_realise": "réalisé",
  "icone_kube": "icones/kubernetes.png",
  "icone_critique": "icones/eclair.png",
  "icone_check_green": "icones/check_green.png",
  "icone_check_blue": "icones/check_blue.png",
  "icone_z": "icones/z.png",
  "icone_reconstruction": "icones/reconstruction.png",
  "icone_restauration": "icones/restauration.png",
  "icone_resynchro": "icones/resynchro.png",
  "squad_colors": {
    "squad 1": "#92D050",
    "squad 2": "#00B050",
    "squad 3": "#00B0F0"
  }
}
```

---

## ✅ Utilisation

1. Vérifie que tous les fichiers nécessaires sont en place
2. Exécute le script :

```bash
python generate_chronogrammes.py planning.xlsx
```

3. Les fichiers `.pptx` seront générés par défaut dans `.`

Options:

- `--config config.json`: Indique le fichier de configuration à utiliser
- `--out "c:\temp\chronogramme"`: Indique le répertoire de sortie à utiliser pour les fichiers PPTX
- `--template "c:\temp\template.pptx"`: Indique le modèle powerpoint à utiliser

---

## 🧪 Conseils

- Les valeurs `"NR"` ou `"NA"` dans la colonne `"réalisé"` sont ignorées
- Le champ `"type"` doit correspondre aux icônes disponibles
- Les trimestres doivent être au format `T1/2025`, `T4/2026`, etc.

---

## 🛠️ Dépendances

- Python 3.10+
- `pandas`
- `python-pptx`

Installation :

```bash
pip install -r requirements.txt
```
