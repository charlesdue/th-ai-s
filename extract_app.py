import os
import re
import fitz  # PyMuPDF
import pandas as pd
import streamlit as st
import tempfile
import zipfile
from io import BytesIO

def extraire_infos_from_text(texte, pdf_name="inconnu.pdf"):
    imputation_matches = re.findall(
        r"(?<=Imputation budgétaire)(?:\n.*?)*?(\d{3,}-\S+)[\n\s]+(P\d+)[\n\s]+(\d{2})[\n\s]+(\d{7})[\n\s]+(\d{10})[\n\s]+(A\d+)",
        texte, re.MULTILINE
    )

    def cap(p):
        match = re.search(p, texte)
        return match.group(1).strip() if match else ""

    communs = {
        "pdf": pdf_name,
        "commande": cap(r"Bon de commande\s*(\d+)"),
        "date_emission": cap(r"Date d[\u2019']?\u00e9mission\s*[:\-]?[\s\n]*(\d{2}\.\d{2}\.\d{4})"),
        "designation": cap(r"00010\s+(.*?)\s+\d"),
        "date_livraison": cap(r"Date de livraison[:\s\n]+(\d{2}\.\d{2}\.\d{4})"),
        "montant_ht": cap(r"Montant HT\s*[:\-]?[\s\n]*([\d\.,]+)"),
        "montant_tva": cap(r"Montant TVA\s*[:\-]?[\s\n]*([\d\.,]+)"),
        "montant_ttc": cap(r"Montant TTC\s*[:\-]?[\s\n]*([\d\.,]+)"),
        "marche": cap(r"March[eé]\s+n[\u00b0:]?[\s\n]*(\d{8,})"),
        "adresse_ligne_1": cap(r"\n\s*([A-Z][A-Z\s]+)\n\s*\d{1,3}\s+RUE"),
    }

    resultats = []
    if imputation_matches:
        for bloc in imputation_matches:
            bloc_infos = communs.copy()
            bloc_infos.update({
                "compte_budgetaire": bloc[0],
                "rubrique": bloc[1],
                "structure_organique": bloc[2],
                "autorisation_programme": bloc[3],
                "destination": bloc[4],
                "n_ec": bloc[5],
                "element_otp": bloc[5],
            })
            resultats.append(bloc_infos)
    else:
        communs.update({
            "compte_budgetaire": "",
            "rubrique": "",
            "structure_organique": "",
            "autorisation_programme": "",
            "destination": "",
            "n_ec": "",
            "element_otp": "",
        })
        resultats.append(communs)

    return resultats

# Streamlit app
st.title("Extraction automatique de bons de commande PDF")
zip_file = st.file_uploader("Upload un fichier ZIP contenant des PDF", type="zip")

if zip_file:
    all_data = []
    with st.spinner("Traitement en cours..."):
        with tempfile.TemporaryDirectory() as tmpdir:
            zip_path = os.path.join(tmpdir, "uploaded.zip")
            with open(zip_path, "wb") as f:
                f.write(zip_file.read())

            with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                zip_ref.extractall(tmpdir)

            for root, dirs, files in os.walk(tmpdir):
                for file in files:
                    if file.lower().endswith(".pdf"):
                        pdf_path = os.path.join(root, file)
                        doc = fitz.open(pdf_path)
                        texte = "".join([page.get_text() for page in doc])
                        resultats = extraire_infos_from_text(texte, pdf_name=file)
                        all_data.extend(resultats)

        df = pd.DataFrame(all_data)
        st.success("Extraction terminée !")
        st.dataframe(df)

        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='Extraction')
        output.seek(0)

        st.download_button(
            label="Télécharger le fichier Excel",
            data=output,
            file_name="resultats_extraction.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.info("Chargez un fichier .zip contenant des fichiers PDF.")
