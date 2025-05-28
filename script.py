from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
import pandas as pd
import os

# Chemins
excel_path = "data/tableau.xlsx"
pdf_template = "data/modele facture pré remplie pro-forma UK.pdf"
output_dir = "output_coord_bg"
os.makedirs(output_dir, exist_ok=True)

# Charger données Excel
df = pd.read_excel(excel_path)
df.columns = df.columns.str.strip().str.replace("’", "'")

# Fonction d'adresse propre
def build_address(row):
    fields = ['Complément de voie', 'Voie', 'Lieu dit/Boite postale', 'Code postal', 'Commune']
    return ', '.join([str(row.get(col, '')).strip() for col in fields if pd.notna(row.get(col))])

# Pour chaque ligne Excel, on génère un PDF
for _, row in df.iterrows():
    ref = str(row['Référence pli'])
    buffer_pdf_path = os.path.join(output_dir, f"{ref}_temp.pdf")
    final_pdf_path = os.path.join(output_dir, f"{ref}.pdf")

    # Créer le PDF temporaire avec texte positionné
    c = canvas.Canvas(buffer_pdf_path, pagesize=A4)
    c.setFont("Helvetica", 10)

    # Écriture des champs aux coordonnées spécifiques
    c.drawString(415, 733, str(ref))  # Numéro de facture
    c.drawString(309, 381, str(row['Quantité']))  # Quantité
    c.drawString(68, 382, str(row['Description']))  # Description

    c.drawString(66, 100, build_address(row))  # Adresse destinataire
    c.drawString(66, 280, f"{row['Civilité']} {row['Nom']} {row['Prénom']}")  # Nom complet
    c.drawString(206, 362, str(row['E-mail']))  # Email
    c.drawString(206, 348, str(row['Téléphone']))  # Téléphone
    c.drawString(310, 375, str(row['Poids net (en kg)']))  # Poids
    c.drawString(360, 375, f"{row['Valeur (en €)']} €")  # Valeur

    c.save()

    # Fusion du PDF modèle + texte
    background = PdfReader(pdf_template)
    overlay = PdfReader(buffer_pdf_path)
    writer = PdfWriter()

    background_page = background.pages[0]
    overlay_page = overlay.pages[0]
    background_page.merge_page(overlay_page)
    writer.add_page(background_page)

    # Enregistrer le PDF final
    with open(final_pdf_path, "wb") as f:
        writer.write(f)

    # Supprimer le fichier temporaire
    os.remove(buffer_pdf_path)
