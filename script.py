from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
import pandas as pd
from datetime import datetime
import os
import re


# === Chemins ===
excel_path = "data/tableau.xlsx"
pdf_template = "data/modele facture pré remplie pro-forma UK.pdf"
output_dir = "output_coord_bg"
os.makedirs(output_dir, exist_ok=True)

# === Charger et nettoyer les colonnes ===
df = pd.read_excel(excel_path)
df.columns = (
    df.columns
    .str.strip()
    .str.replace("’", "'", regex=False)
    .str.replace("é", "e")
    .str.replace("è", "e")
    .str.replace("ê", "e")
    .str.replace("à", "a")
    .str.replace("ç", "c")
    .str.replace(r"\s+", "_", regex=True)
    .str.lower()
)


# === Fonction pour sécuriser les accès aux champs ===
def safe_str(value, default=""):
    return str(value) if pd.notna(value) else default

def clean_text(value, sep=", "):
    if pd.isna(value):
        return ""
    text = str(value)
    # Remplacer les caractères "carré noir", "non imprimables", ou similaires
    text = text.replace("■", sep)
    text = re.sub(r"[^\x20-\x7EÀ-ÿ]", "", text)  # Supprime tout caractère non imprimable ASCII étendu
    return text.strip()


# === Fonction d'adresse propre ===
def build_address(row):
    champs = ['complement_de_voie', 'adresse', 'lieu_dit/boite_postale', 'codepostal', 'ville']
    return ', '.join([safe_str(row.get(c)) for c in champs if pd.notna(row.get(c))])

# === Génération des PDFs ===
for index, row in df.iterrows():

    ref = safe_str(row.get('reference_pli', 'sans_ref'))
    buffer_pdf_path = os.path.join(output_dir, f"{ref}_temp.pdf")
    final_pdf_path = os.path.join(output_dir, f"{ref}.pdf")

    # Création du PDF temporaire avec les textes positionnés
    c = canvas.Canvas(buffer_pdf_path, pagesize=A4)
    c.setFont("Helvetica", 10)

    # x:0 a gauche
    # y:0 en bas

    # Haut de page : numéro de facture
    c.drawString(415, 731, ref)

    # Informations destinataire
    nom_complet = f"{safe_str(row.get('nom')).upper()} {safe_str(row.get('prenom')).upper()}"

    c.drawString(77, 700, "Bip&Go")
    c.drawString(77, 688, "Echangeur des Essarts")
    c.drawString(77, 676, "76 530 Grand Couronne")
    c.drawString(244, 639, "750 535 288 00021")
    c.drawString(217, 619, "09 70 80 87 65")
    c.drawString(167, 599, "bipandgologistique@sanef.com")

    c.drawString(67, 540, nom_complet)
    c.drawString(67, 530, safe_str(row.get('codepostal')))
    c.drawString(117, 530, safe_str(row.get('ville')))
    c.drawString(67, 520, clean_text(row.get('adresse')))
    c.drawString(307, 510, safe_str(row.get('pays')))
    c.drawString(217, 492, safe_str(row.get('telmobile')))
    c.drawString(167, 472, safe_str(row.get('email')))

    # Description et infos douanières
    # c.drawString(77, 691, build_address(row))

    c.drawString(117, 380, safe_str(row.get('description')))
    c.drawString(307, 380, safe_str(row.get('quantite')))
    c.drawString(367, 380, safe_str(row.get('n°tarifaire_du_sh')))
    c.drawString(467, 380, safe_str(row.get('pays_d\'origine')))

    c.drawString(380, 210, safe_str(row.get('valeur_(en_€)')))

    c.drawString(77, 95, "Grand Couronne")
    date_ajd = datetime.now().strftime("%d/%m/%Y")
    c.drawString(155, 95, date_ajd)

    c.save()

    # Fusion avec le PDF modèle
    background = PdfReader(pdf_template)
    overlay = PdfReader(buffer_pdf_path)
    writer = PdfWriter()
    background_page = background.pages[0]
    background_page.merge_page(overlay.pages[0])
    writer.add_page(background_page)

    with open(final_pdf_path, "wb") as f:
        writer.write(f)

    os.remove(buffer_pdf_path)
