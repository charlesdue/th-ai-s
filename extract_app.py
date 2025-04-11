import os
import re
import fitz  # PyMuPDF
import pandas as pd
import streamlit as st


def extraire_infos(pdf_path, export_txt=False, txt_dir=None):
    doc = fitz.open(pdf_path)
    texte = ""
    for page in doc:
        texte += page.get_text()

    if export_txt:
        if txt_dir:
            os.makedirs(txt_dir, exist_ok=True)
            txt_path = os.path.join(txt_dir, os.path.splitext(os.path.basename(pdf_path))[0] + ".txt")
        else:
            txt_path = os.path.splitext(pdf_path)[0] + ".txt"
        with open(txt_path, "w", encoding="utf-8") as f:
            f.write(texte)

    # Recherche de la section d'imputation budgétaire complète
    imputation_matches = re.findall(
        r"(?<=Imputation budgétaire)(?:\n.*?)*?(\d{3,}-\S+)[\n\s]+(P\d+)[\n\s]+(\d{2})[\n\s]+(\d{7})[\n\s]+(\d{10})[\n\s]+(A\d+)",
        texte, re.MULTILINE
    )

    def extract_communs():
        def cap(p):
            match = re.search(p, texte)
            return match.group(1).strip() if match else ""

        return {
            "pdf": os.path.basename(pdf_path),
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
    communs = extract_communs()

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
                "element_otp": bloc[5],  # si diff, à corriger
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


def parcourir_dossier(dossier, export_txt=False):
    donnees = []
    txt_dir = os.path.join(dossier, "txt_extraits") if export_txt else None
    for racine, sous_dossiers, fichiers in os.walk(dossier):
        profondeur = racine[len(dossier):].count(os.sep)
        if profondeur > 1:
            continue

        for f in fichiers:
            if f.lower().endswith(".pdf") and "preuve_" not in f.lower():
                chemin_pdf = os.path.join(racine, f)
                blocs = extraire_infos(chemin_pdf, export_txt=export_txt, txt_dir=txt_dir)
                donnees.extend(blocs)

    return donnees


# Streamlit app
st.title("Extraction automatique de bons de commande PDF")
dossier = st.text_input("Chemin du dossier contenant les PDF :")
export_txt = st.checkbox("Sauvegarder aussi les textes extraits (.txt)", value=True)

if dossier and os.path.isdir(dossier):
    if st.button("Lancer l'extraction"):
        with st.spinner("Extraction en cours..."):
            resultats = parcourir_dossier(dossier, export_txt=export_txt)
            df = pd.DataFrame(resultats)
            csv_path = os.path.join(dossier, "resultats_extraction.csv")
            df.to_csv(csv_path, index=False)
            st.success("Extraction terminée !")
            st.download_button("Télécharger le CSV", data=df.to_csv(index=False), file_name="resultats_extraction.csv", mime="text/csv")
else:
    st.info("Veuillez entrer un chemin de dossier valide.")