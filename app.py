import streamlit as st
import pandas as pd
from datetime import datetime
from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from io import BytesIO
import tempfile
import os
import re
import zipfile
from copy import deepcopy

st.title("Générateur de PDF à partir d'un Excel")

uploaded_excel = st.file_uploader("Téléverse ton fichier Excel", type=["xlsx"])
template_pdf = st.file_uploader("Téléverse le modèle PDF (vierge)", type=["pdf"])

def safe_str(value, default=""):
    return str(value) if pd.notna(value) else default

def clean_text(value, sep=", "):
    if pd.isna(value):
        return ""
    text = str(value)
    text = text.replace("■", sep)
    text = re.sub(r"[^\x20-\x7EÀ-ÿ]", "", text)
    return text.strip()

if uploaded_excel and template_pdf:
    df = pd.read_excel(uploaded_excel)

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

    with st.spinner("📄 Génération des fichiers PDF..."):
        output_dir = tempfile.mkdtemp()
        zip_path = os.path.join(output_dir, "documents.zip")

        with zipfile.ZipFile(zip_path, "w") as zipf:
            for _, row in df.iterrows():
                ref = safe_str(row.get('reference_pli', 'sans_ref'))
                final_pdf_path = os.path.join(output_dir, f"{ref}.pdf")

                packet = BytesIO()
                c = canvas.Canvas(packet, pagesize=A4)
                c.setFont("Helvetica", 10)

                date_ajd = datetime.now().strftime("%d/%m/%Y")
                nom_complet = f"{safe_str(row.get('nom')).upper()} {safe_str(row.get('prenom')).upper()}"

                c.drawString(415, 731, ref)
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

                c.drawString(117, 380, safe_str(row.get('description')))
                c.drawString(307, 380, safe_str(row.get('quantite')))
                c.drawString(367, 380, safe_str(row.get('n°tarifaire_du_sh')))
                c.drawString(467, 380, safe_str(row.get("pays_d'origine")))
                c.drawString(380, 210, safe_str(row.get('valeur_(en_€)')))

                c.drawString(77, 95, "Grand Couronne")
                c.drawString(155, 95, date_ajd)
                c.save()
                packet.seek(0)

                overlay = PdfReader(packet)
                background = PdfReader(template_pdf)
                writer = PdfWriter()

                base_page = deepcopy(background.pages[0])
                overlay_page = overlay.pages[0]
                base_page.merge_page(overlay_page)

                for _ in range(3):
                    writer.add_page(deepcopy(base_page))

                with open(final_pdf_path, "wb") as f:
                    writer.write(f)

                zipf.write(final_pdf_path, arcname=os.path.basename(final_pdf_path))

        with open(zip_path, "rb") as f:
            st.success("✅ Tous les PDF ont été générés et compressés avec succès !")
            st.download_button(
                label="📦 Télécharger tous les PDF (.zip)",
                data=f,
                file_name="pdf_export.zip",
                mime="application/zip"
            )
