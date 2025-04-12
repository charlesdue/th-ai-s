import os
import re
import fitz  # PyMuPDF
import pandas as pd
import streamlit as st
import tempfile

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
uploaded_files = st.file_uploader("Upload un ou plusieurs PDF", type="pdf", accept_multiple_files=True)

if uploaded_files:
    all_data = []
    with st.spinner("Traitement en cours..."):
        for uploaded_file in uploaded_files:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
                tmp.write(uploaded_file.read())
                tmp_path = tmp.name

            doc = fitz.open(tmp_path)
            texte = "".join([page.get_text() for page in doc])
            resultats = extraire_infos_from_text(texte, pdf_name=uploaded_file.name)
            all_data.extend(resultats)

        df = pd.DataFrame(all_data)
        st.success("Extraction terminée !")
        st.dataframe(df)
        st.download_button("Télécharger le CSV", data=df.to_csv(index=False), file_name="resultats_extraction.csv", mime="text/csv")
else:
    st.info("Chargez un ou plusieurs fichiers PDF pour commencer.")
