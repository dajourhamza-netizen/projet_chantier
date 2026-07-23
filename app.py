import os
import io
import zipfile
import subprocess
import tempfile
import pandas as pd
import streamlit as st
from docxtpl import DocxTemplate

# ==========================================
# CONFIGURATION DE LA PAGE STREAMLIT
# ==========================================
st.set_page_config(
    page_title="Suivi Chantier - Demandes d'Intervention",
    page_icon="🏗️",
    layout="wide"
)

st.title("🏗️ Suivi Chantier - Générateur de Demandes d'Intervention (DI)")
st.write("---")

# Répertoire racine du projet (GitHub)
DOSSIER_CHANTIER = os.path.dirname(os.path.abspath(__file__))

# ==========================================
# FONCTIONS UTILITAIRES & HELPER FUNCTIONS
# ==========================================

def get_col_val(row, *possible_cols):
    """
    Extrait la valeur d'une colonne en testant plusieurs noms possibles 
    sans se soucier des majuscules/minuscules ou des espaces.
    """
    row_cols_clean = {str(k).strip().upper(): k for k in row.index}
    for col in possible_cols:
        col_clean = str(col).strip().upper()
        if col_clean in row_cols_clean:
            val = row[row_cols_clean[col_clean]]
            if pd.notna(val):
                return str(val).strip()
    return ""


def obtenir_modele_word(nature_travaux, dossier_chantier):
    """
    LIAISON AUTOMATIQUE : Associe la nature des travaux
    au bon fichier .docx présent sur GitHub.
    """
    if not nature_travaux:
        return None
    
    cle = str(nature_travaux).strip().upper()
    
    # Mapping exact avec vos fichiers .docx sur GitHub
    mapping_fichiers = {
        "ARASE DE PST": "ARASE DE PST.docx",
        "ARASE DE TERRASSEMENT": "ARASE DE TERRASSEMENT.docx",
        "ASSISE DE REMBLAI": "ASSISE DE REMBLAI.docx",
        "COUCHE DE FORME": "COUCHE DE FORME.docx",
        "DEGAGEMENT D'EMPRISE": "DEGAGEMENT D'EMPRISE.docx",
        "DÉCAPAGE": "DÉCAPAGE.docx",
        "DECAPAGE": "DÉCAPAGE.docx",
        "REMBLAI PST": "REMBLAI PST.docx",
        "REMBLAIS CONTIGUS": "REMBLAIS CONTIGUS.docx",
        "REMBLAIS DE FOUILLE": "REMBLAIS DE FOUILLE.docx",
        "REMBLAIS": "REMBLAIS.docx",
    }
    
    # 1. Correspondance exacte via le dictionnaire
    if cle in mapping_fichiers:
        chemin = os.path.join(dossier_chantier, mapping_fichiers[cle])
        if os.path.exists(chemin):
            return chemin

    # 2. Recherche tolérante si le nom du fichier sur GitHub est proche
    if os.path.exists(dossier_chantier):
        fichiers = [f for f in os.listdir(dossier_chantier) if f.endswith('.docx')]
        for f in fichiers:
            nom_sans_ext = os.path.splitext(f)[0].strip().upper()
            if nom_sans_ext == cle or cle in nom_sans_ext:
                return os.path.join(dossier_chantier, f)
            
    return None


def convertir_docx_en_pdf(docx_bytes):
    """Convertit le fichier DOCX en PDF via LibreOffice si disponible."""
    try:
        with tempfile.TemporaryDirectory() as tmpdir:
            in_path = os.path.join(tmpdir, "document.docx")
            with open(in_path, "wb") as f:
                f.write(docx_bytes)
            
            cmd = ["libreoffice", "--headless", "--convert-to", "pdf", "--outdir", tmpdir, in_path]
            subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, check=True, timeout=30)
            
            out_path = os.path.join(tmpdir, "document.pdf")
            if os.path.exists(out_path):
                with open(out_path, "rb") as f:
                    return f.read()
    except Exception:
        pass
    return None


def generer_docx_et_pdf_bytes(modele_path, contexte):
    """Remplit le modèle Word Jinja2 et produit les bytes DOCX et PDF."""
    doc = DocxTemplate(modele_path)
    doc.render(contexte)
    
    docx_io = io.BytesIO()
    doc.save(docx_io)
    docx_bytes = docx_io.getvalue()
    
    pdf_bytes = convertir_docx_en_pdf(docx_bytes)
    return docx_bytes, pdf_bytes


def creer_zip_fichiers(liste_fichiers):
    """Crée un fichier ZIP contenant l'ensemble des documents générés."""
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        for item in liste_fichiers:
            if item.get('bytes_docx'):
                zip_file.writestr(f"DOCX/{item['nom_docx']}", item['bytes_docx'])
            if item.get('bytes_pdf'):
                zip_file.writestr(f"PDF/{item['nom_pdf']}", item['bytes_pdf'])
    return zip_buffer.getvalue()


# ==========================================
# 1. CHARGEMENT ET FILTRAGE DES DONNÉES EXCEL
# ==========================================
fichier_excel = os.path.join(DOSSIER_CHANTIER, "suivi .xlsx")
if not os.path.exists(fichier_excel):
    fichier_excel = os.path.join(DOSSIER_CHANTIER, "suivi.xlsx")

if not os.path.exists(fichier_excel):
    st.error("❌ Fichier Excel `suivi.xlsx` introuvable sur le dépôt GitHub.")
    st.stop()

try:
    df = pd.read_excel(fichier_excel)
    # Formater la colonne DATE
    col_date = next((c for c in df.columns if "DATE" in str(c).upper()), None)
    if col_date:
        df[col_date] = pd.to_datetime(df[col_date], errors='coerce').dt.strftime('%d/%m/%Y')
    else:
        st.error("❌ Impossible de trouver la colonne 'DATE' dans le fichier Excel.")
        st.stop()
except Exception as e:
    st.error(f"❌ Erreur lors de la lecture du fichier Excel : {e}")
    st.stop()

# Barre latérale - Sélection de la date
st.sidebar.header("📅 Sélection du Jour")
dates_disponibles = [d for d in df[col_date].dropna().unique() if d != ""]
date_choisie = st.sidebar.selectbox("Choisissez une date de suivi :", options=dates_disponibles)

# Filtrage du DataFrame
df_jour = df[df[col_date] == date_choisie].copy()

st.subheader(f"📌 Travaux enregistrés pour le : `{date_choisie}` ({len(df_jour)} ligne(s))")
st.dataframe(df_jour, use_container_width=True)

st.write("---")

# ==========================================
# 2. GENERATION DES DOCUMENTS
# ==========================================
st.header("⚙️ Option de Génération des Demandes d'Intervention")

tab1, tab2 = st.tabs(["📄 Option 1 : Fiches Individuelles (Liaison Automatique)", "📑 Option 2 : DI Consolidation Journalière"])

# ------------------------------------------
# TAB 1 : FICHES INDIVIDUELLES (LIAISON)
# ------------------------------------------
with tab1:
    st.markdown("##### 📄 **Fiches Individuelles liées aux modèles Word par Nature de Travaux**")
    st.info("Chaque ligne de l'Excel est liée à son modèle spécifique `.docx` selon la colonne **TITRE DE LA NATURE DES TRAVAUX**.")

    if st.button(f"⚡ Générer toutes les fiches du {date_choisie}", type="primary", use_container_width=True):
        if df_jour.empty:
            st.warning("⚠️ Aucune donnée pour cette date.")
        else:
            try:
                with st.spinner("⏳ Recherche des modèles Word et génération des fiches..."):
                    fichiers_generes = []
                    fichiers_introuvables = []

                    for idx, row in df_jour.iterrows():
                        nature = get_col_val(row, "TITRE DE LA NATURE DES TRAVAUX", "NATURE")
                        
                        # Liaison dynamique
                        modele_path = obtenir_modele_word(nature, DOSSIER_CHANTIER)
                        
                        if not modele_path:
                            fichiers_introuvables.append(nature or "Nature inconnue")
                            continue

                        contexte = {
                            'DATE': get_col_val(row, "DATE") or date_choisie,
                            'NATURE': nature,
                            'ACTIVITE': get_col_val(row, "ACTIVITÉ RÉALISÉE", "ACTIVITE"),
                            'PARTIE': get_col_val(row, "PARTIE D'OUVRAGE", "PARTIE"),
                            'SITUATION': get_col_val(row, "SITUATION", "PK"),
                            'ESSAI': get_col_val(row, "ÉSSAI/ CONTRÔLE RÉALISÉE", "ESSAI"),
                            'OBSERVATION': get_col_val(row, "OBSERVATION", "OBS")
                        }

                        docx_b, pdf_b = generer_docx_et_pdf_bytes(modele_path, contexte)
                        nom_clean = f"{nature}_{idx+1}".replace(" ", "_").replace("'", "").replace("/", "-")

                        fichiers_generes.append({
                            'nom_docx': f"{nom_clean}.docx",
                            'bytes_docx': docx_b,
                            'nom_pdf': f"{nom_clean}.pdf",
                            'bytes_pdf': pdf_b
                        })

                    if fichiers_introuvables:
                        st.warning(f"⚠️ Modèle `.docx` non trouvé pour : {', '.join(set(fichiers_introuvables))}")

                    if fichiers_generes:
                        st.success(f"✅ {len(fichiers_generes)} fiche(s) générée(s) avec succès !")
                        zip_bytes = creer_zip_fichiers(fichiers_generes)
                        
                        st.download_button(
                            "📦 Télécharger toutes les fiches (ZIP)",
                            data=zip_bytes,
                            file_name=f"Fiches_DI_{date_choisie.replace('/', '-')}.zip",
                            mime="application/zip",
                            use_container_width=True
                        )
            except Exception as e:
                st.error(f"❌ Erreur lors de la génération : {e}")

# ------------------------------------------
# TAB 2 : DI CONSOLIDATION JOURNALIÈRE
# ------------------------------------------
with tab2:
    st.markdown("##### 📑 **Demande d'Intervention globale regroupant tout le tableau du jour**")
    st.info("Utilise le modèle Word général `Demande d'intervention.docx` pour regrouper toutes les lignes de la journée.")

    # Recherche du modèle journalier
    noms_modeles_di = [
        "Demande d'intervention.docx",
        "Demande_intervention.docx",
        "Demande d intervention.docx",
        "DI.docx"
    ]
    modele_di_global = None
    for nom in noms_modeles_di:
        p = os.path.join(DOSSIER_CHANTIER, nom)
        if os.path.exists(p):
            modele_di_global = p
            break

    if st.button(f"📑 Générer la DI globale du {date_choisie}", type="secondary", use_container_width=True):
        if not modele_di_global:
            st.error("❌ Impossible de trouver le modèle `Demande d'intervention.docx` sur GitHub.")
        else:
            try:
                with st.spinner("⏳ Génération du document récapitulatif journalier..."):
                    liste_activites = []
                    
                    for _, row in df_jour.iterrows():
                        date_row = get_col_val(row, "DATE") or date_choisie
                        nature = get_col_val(row, "TITRE DE LA NATURE DES TRAVAUX", "NATURE")
                        partie = get_col_val(row, "PARTIE D'OUVRAGE", "PARTIE")
                        situation = get_col_val(row, "SITUATION", "PK")
                        activite = get_col_val(row, "ACTIVITÉ RÉALISÉE", "ACTIVITE")
                        essai = get_col_val(row, "ÉSSAI/ CONTRÔLE RÉALISÉE", "ESSAI")

                        # Fusion propre des colonnes
                        act_nat = f"{activite} - {nature}".strip(" -") if activite and nature else activite or nature
                        partie_sit = f"{partie} / {situation}".strip(" /") if partie and situation else partie or situation

                        liste_activites.append({
                            'DATE': date_row,
                            'ACTIVITE_NATURE': act_nat,
                            'PARTIE_SITUATION': partie_sit,
                            'ESSAI': essai
                        })

                    contexte_global = {
                        'DATE': date_choisie,
                        'TRAVAUX': liste_activites
                    }

                    docx_b, pdf_b = generer_docx_et_pdf_bytes(modele_di_global, contexte_global)
                    date_clean = date_choisie.replace("/", "-")

                    st.success(f"✅ DI Globale générée avec {len(liste_activites)} ligne(s) !")
                    
                    c_down1, c_down2 = st.columns(2)
                    with c_down1:
                        st.download_button(
                            "📝 Télécharger DOCX (DI Globale)",
                            data=docx_b,
                            file_name=f"DI_Globale_{date_clean}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            use_container_width=True
                        )
                    with c_down2:
                        if pdf_b:
                            st.download_button(
                                "📕 Télécharger PDF (DI Globale)",
                                data=pdf_b,
                                file_name=f"DI_Globale_{date_clean}.pdf",
                                mime="application/pdf",
                                use_container_width=True
                            )
                        else:
                            st.info("ℹ️ La version PDF n'est pas disponible (LibreOffice non présent).")
            except Exception as e:
                st.error(f"❌ Erreur lors de la génération : {e}")
