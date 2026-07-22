import streamlit as st
import pandas as pd
from datetime import datetime
import os
import re
import io
import unicodedata
import zipfile
import subprocess
import tempfile
import shutil

from docxtpl import DocxTemplate, RichText
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter

# ==========================================
# 0. CONFIGURATION & STYLES (BATISCRIPT UI)
# ==========================================
st.set_page_config(
    page_title="BATISCRIPT - Suivi Chantier",
    page_icon="🏗️",
    layout="wide"
)

# Custom CSS matching Batiscript UI design exactly
st.markdown("""
<style>
    /* Main Canvas Background */
    .stApp {
        background-color: #f8f9fc !important;
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
    }

    /* Adjust Streamlit Top Header to prevent topbar clipping */
    header[data-testid="stHeader"] {
        background-color: transparent !important;
        z-index: 1 !important;
    }

    /* Container padding fixed so topbar is fully visible */
    .block-container {
        padding-top: 3.5rem !important;
        padding-bottom: 2rem !important;
    }

    /* Top Bar - Dark Navy Header */
    .batiscript-topbar {
        background-color: #1c2237;
        color: #ffffff;
        padding: 14px 24px;
        border-radius: 8px;
        margin-bottom: 25px;
        display: flex;
        justify-content: space-between;
        align-items: center;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.05);
    }
    .topbar-brand {
        font-size: 18px;
        font-weight: 900;
        letter-spacing: 1px;
        color: #ffffff;
        display: flex;
        align-items: center;
        gap: 10px;
    }
    .topbar-nav {
        display: flex;
        gap: 20px;
        font-size: 13px;
        color: #a0aec0;
        font-weight: 500;
    }
    .topbar-nav span {
        cursor: pointer;
        transition: color 0.2s;
    }
    .topbar-nav span.active {
        color: #ffffff;
        font-weight: 700;
        border-bottom: 2px solid #5b58e6;
        padding-bottom: 2px;
    }

    /* Sidebar Batiscript Light Style */
    section[data-testid="stSidebar"] {
        background-color: #f4f5f8 !important;
        border-right: 1px solid #e2e8f0;
    }
    section[data-testid="stSidebar"] * {
        color: #2d3748 !important;
    }

    /* Custom Project Header Card in Sidebar */
    .sidebar-project-card {
        background-color: #ffffff;
        border-radius: 10px;
        padding: 12px;
        border: 1px solid #e2e8f0;
        margin-bottom: 15px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.03);
    }

    /* Green Bottom Action Button Style */
    .st-key-btn_validate > button {
        background-color: #27c383 !important;
        color: #ffffff !important;
        border: none !important;
        border-radius: 8px !important;
        font-weight: 700 !important;
        padding: 10px 20px !important;
        width: 100%;
        box-shadow: 0 2px 4px rgba(39, 195, 131, 0.2);
    }
    .st-key-btn_validate > button:hover {
        background-color: #21a971 !important;
    }

    /* Action Buttons Top Right (+ Titre, + Tâche, + Note) */
    .btn-action-purple {
        background-color: #6b1d99;
        color: white;
        padding: 6px 16px;
        border-radius: 6px;
        font-weight: 700;
        font-size: 13px;
        display: inline-block;
    }
    .btn-action-blue {
        background-color: #33b3ff;
        color: white;
        padding: 6px 16px;
        border-radius: 6px;
        font-weight: 700;
        font-size: 13px;
        display: inline-block;
    }
    .btn-action-green {
        background-color: #26a69a;
        color: white;
        padding: 6px 16px;
        border-radius: 6px;
        font-weight: 700;
        font-size: 13px;
        display: inline-block;
    }

    /* Tabs Layout Style */
    .stTabs [data-baseweb="tab-list"] {
        background-color: #edf2f7;
        padding: 5px;
        border-radius: 10px;
        gap: 6px;
    }
    .stTabs [data-baseweb="tab"] {
        background-color: transparent;
        border-radius: 8px;
        padding: 8px 18px;
        font-weight: 700;
        color: #4a5568;
        border: none !important;
    }
    .stTabs [aria-selected="true"] {
        background-color: #ffffff !important;
        color: #5b58e6 !important;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
    }

    /* Primary Submit Button */
    .stButton > button[kind="primary"] {
        background-color: #5b58e6 !important;
        color: #ffffff !important;
        border-radius: 8px !important;
        font-weight: 700 !important;
        border: none !important;
        padding: 10px 22px !important;
    }
</style>
""", unsafe_allow_html=True)

# ==========================================
# 1. FONCTIONS ET INITIALISATION
# ==========================================
def text_to_richtext(text):
    if not text or pd.isna(text):
        return ""
    rt = RichText()
    lines = str(text).split('\n')
    for i, line in enumerate(lines):
        rt.add(line)
        if i < len(lines) - 1:
            rt.add('\n')
    return rt

DOSSIER_CHANTIER = os.path.dirname(os.path.abspath(__file__)) if '__file__' in globals() else os.getcwd()

chemin_excel_defaut = os.path.join(DOSSIER_CHANTIER, "suivi .xlsx")
if not os.path.exists(chemin_excel_defaut):
    chemin_excel_defaut = os.path.join(DOSSIER_CHANTIER, "suivi.xlsx")

COL_PARTIE = "PARTIE D'OUVRAGE"

COLUMNS_TEMPLATE = [
    "DATE", "TITRE DE LA NATURE DES TRAVAUX", COL_PARTIE, 
    "SITUATION", "ACTIVITÉ RÉALISÉE", "ÉSSAI/ CONTRÔLE RÉALISÉE", 
    "RÉFÉRENCE DE PROCÉDURE", "PIÈCES JOINTES"
]

LIAISONS = {
    "ARASE DE PST": {"procedure": "TER-PEX-05-00", "pieces": "* Fiche de suivi de la PST\n* Fiche de réception topographique\n* PVs laboratoire"},
    "ARASE DE TERRASSEMENT": {"procedure": "TER-PEX-03-00", "pieces": "* Fiche de contrôle des déblais\n* Fiche de réception topographique\n* PVs laboratoire"},
    "ASSISE DE REMBLAIS PURGE": {"procedure": "TER-PEX-04-00", "pieces": "* Fiche de réception de l'assise des remblais\n* Fiche de réception topographique\n* Fiche d'identification de la purge\n* PVs laboratoire"},
    "ASSISE DE REMBLAIS": {"procedure": "TER-PEX-04-00", "pieces": "* Fiche de réception de l'assise des remblais\n* Fiche de réception topographique\n* PVs laboratoire"},
    "ASSISE DE REMBLAIS CDF": {"procedure": "TER-PEX-04-00", "pieces": "* Fiche de réception de l'assise des remblais\n* Fiche de réception topographique\n* PVs laboratoire"},
    "ASSISE DE REMBLAIS CONTIGUS": {"procedure": "OVA-PEX-16-00", "pieces": "* Fiche de suivi des remblais contigus\n* Fiche de contrôle des remblais contigus\n* PVs laboratoire\n* Fiche de réception topographique"},
    "ASSISE DE REMBLAI DE FOUILLE": {"procedure": "OVA-PEX-04-00", "pieces": "* Fiche de suivi et de contrôle des fouilles et remblaiement de fouilles\n* PVs laboratoire"},
    "ASSISE DE REMBLAIS RENFORCE": {"procedure": "TER-PEX-13-00", "pieces": "* PV Manifold\n* PVs laboratoire\n* Fiche de réception topographique\n* Fiche de réception assise remblai renforcé"},
    "ASSISE DRAINANTE": {"procedure": "TER-PEX-13-00", "pieces": "* Fiche de réception topographique\n* PVs laboratoire\n* Fiche de contrôle de l'assise drainante"},
    "COUCHE DE FORME": {"procedure": "TER-PEX-09-00", "pieces": "* Fiche de suivi et de contrôle de la CDF\n* Fiche de réception topographique\n* PVs laboratoire"},
    "DÉCAPAGE": {"procedure": "TER-PEX-02-00", "pieces": "* Fiche de suivi et de contrôle du décapage\n* Fiche des sections à décaper\n* Fiche de réception topographique"},
    "DEGAGEMENT D'EMPRISE": {"procedure": "TER-PEX-01-00", "pieces": "* Fiche de suivi et de contrôle du dégagement des emprises\n* Fiche de réception topographique\n* Constat dégagement d'emprise"},
    "REMBLAIS": {"procedure": "TER-PEX-04-00", "pieces": "* Fiche de suivi et de contrôle des remblais\n* PVs laboratoire"},
    "REMBLAIS CDF": {"procedure": "TER-PEX-04-00", "pieces": "* Fiche de suivi et de contrôle des remblais\n* PVs laboratoire"},
    "REMBLAIS CONTIGUS": {"procedure": "OVA-PEX-16-00", "pieces": "* Fiche de suivi des remblais contigus\n* Fiche de contrôle des remblais contigus\n* PVs laboratoire\n* Fiche de réception topographique"},
    "REMBLAIS DE FOUILLE": {"procedure": "OVA-PEX-04-00", "pieces": "* Fiche de suivi et de contrôle des fouilles et remblaiement de fouilles\n* PVs laboratoire"},
    "REMBLAIS DE FOUILLS CDF": {"procedure": "OVA-PEX-04-00", "pieces": "* Fiche de suivi et de contrôle des fouilles et remblaiement de fouilles\n* PVs laboratoire"},
    "REMBLAIS RENFORCE": {"procedure": "TER-PEX-13-00", "pieces": "* Fiche de suivi des remblais renforcé\n* Fiche de contrôle des armatures Geostrap\n* Fiche de réception de pose des ecailles\n* PVs laboratoire"},
    "REMBLAIS PST": {"procedure": "TER-PEX-05-00", "pieces": "* Fiche de suivi et de contrôle des remblais PST\n* PVs laboratoire"}
}

def clean_filename(text):
    if not text:
        return ""
    text = str(text)
    if text.lower().endswith('.docx'):
        text = text[:-5]
    text = unicodedata.normalize('NFD', text).encode('ascii', 'ignore').decode("utf-8")
    text = re.sub(r'\s+', ' ', text)
    return re.sub(r'[^a-zA-Z0-9]', '', text).lower()

def trouver_modele_word(nom_nature):
    target_clean = clean_filename(nom_nature)
    if os.path.exists(DOSSIER_CHANTIER):
        for file in os.listdir(DOSSIER_CHANTIER):
            if file.lower().endswith('.docx') and not file.startswith('~$'):
                if clean_filename(file) == target_clean:
                    return os.path.join(DOSSIER_CHANTIER, file)
    return None

def construire_nom_pdf(row):
    nature = str(row.get('TITRE DE LA NATURE DES TRAVAUX', '')).strip()
    partie = str(row.get(COL_PARTIE, row.get("PARTIE D meOUVRAGE", ''))).strip()
    situation = str(row.get('SITUATION', '')).strip()

    nom_brut = f"{nature} - {partie} - {situation}"
    nom_propre = re.sub(r'[\\/*?:"<>|]', "_", nom_brut)
    nom_propre = re.sub(r'\s+', ' ', nom_propre).strip()
    return f"{nom_propre}.pdf"

def trouver_executable_libreoffice():
    paths_windows = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ]
    for p in paths_windows:
        if os.path.exists(p):
            return p
    for cmd in ["libreoffice", "soffice"]:
        if shutil.which(cmd):
            return cmd
    return None

def generer_docx_et_pdf_bytes(chemin_modele, contexte):
    with tempfile.TemporaryDirectory() as temp_dir:
        doc = DocxTemplate(chemin_modele)
        doc.render(contexte)
        
        docx_temp_path = os.path.join(temp_dir, "temp.docx")
        doc.save(docx_temp_path)
        
        with open(docx_temp_path, "rb") as f:
            docx_bytes = f.read()

        pdf_temp_path = os.path.join(temp_dir, "temp.pdf")
        pdf_bytes = None

        try:
            from docx2pdf import convert
            convert(docx_temp_path, pdf_temp_path)
            if os.path.exists(pdf_temp_path):
                with open(pdf_temp_path, "rb") as f:
                    pdf_bytes = f.read()
        except Exception:
            pass

        if pdf_bytes is None:
            exe_libreoffice = trouver_executable_libreoffice()
            if exe_libreoffice:
                cmd = [exe_libreoffice, "--headless", "--convert-to", "pdf", docx_temp_path, "--outdir", temp_dir]
                subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                if os.path.exists(pdf_temp_path):
                    with open(pdf_temp_path, "rb") as f:
                        pdf_bytes = f.read()

        if pdf_bytes is None:
            raise RuntimeError("Convertisseur PDF introuvable.")

        return docx_bytes, pdf_bytes

def get_sheet_names(filepath):
    if os.path.exists(filepath):
        try:
            return pd.ExcelFile(filepath).sheet_names
        except Exception:
            return ["Chantier Principal"]
    return ["Chantier Principal"]

def save_to_excel_with_formatting(df_to_save, filepath, sheet_name="Chantier Principal"):
    try:
        if "Imprimer" in df_to_save.columns:
            df_to_save = df_to_save.drop(columns=["Imprimer"])

        mode = "a" if os.path.exists(filepath) else "w"
        kwargs = {"mode": mode, "engine": "openpyxl"}
        if mode == "a":
            kwargs["if_sheet_exists"] = "replace"

        with pd.ExcelWriter(filepath, **kwargs) as writer:
            df_to_save.to_excel(writer, sheet_name=sheet_name, index=False)
            worksheet = writer.sheets[sheet_name]
            
            for column in worksheet.columns:
                max_len = max(len(str(cell.value or "")) for cell in column)
                col_letter = column[0].column_letter
                worksheet.column_dimensions[col_letter].width = min(max_len + 4, 60)
            
            num_rows = max(len(df_to_save) + 1, 2)
            num_cols = len(df_to_save.columns)
            if num_cols > 0:
                end_col = get_column_letter(num_cols)
                clean_sheet_name = re.sub(r'\W+', '_', sheet_name)
                worksheet._tables.clear()
                tab = Table(displayName=f"Tableau_{clean_sheet_name}", ref=f"A1:{end_col}{num_rows}")
                tab.tableStyleInfo = TableStyleInfo(name="TableStyleMedium3", showRowStripes=True)
                worksheet.add_table(tab)
        return True, "✅ Mis à jour !"
    except Exception as e:
        return False, f"❌ Erreur : {e}"

# ==========================================
# 2. HEADER TOP BAR (BATISCRIPT STYLE)
# ==========================================
st.markdown("""
<div class="batiscript-topbar">
    <div class="topbar-brand">
        🏢 BATISCRIPT
    </div>
    <div class="topbar-nav">
        <span class="active">▲ Chantiers</span>
        <span>📑 Referentiel</span>
        <span>📋 Non-conformités</span>
        <span>📊 Statistiques</span>
    </div>
    <div style="font-size: 13px; font-weight: 600;">
        👤 Conducteur de Travaux ▼
    </div>
</div>
""", unsafe_allow_html=True)

# ==========================================
# 3. SIDEBAR (BATISCRIPT LIGHT NAVIGATION)
# ==========================================
chantiers_existants = get_sheet_names(chemin_excel_defaut)

st.sidebar.markdown('<div class="sidebar-project-card">', unsafe_allow_html=True)
st.sidebar.markdown("##### 🏢 **Sélection du Chantier**")
chantier_actif = st.sidebar.selectbox("", options=chantiers_existants, label_visibility="collapsed")
st.sidebar.markdown('</div>', unsafe_allow_html=True)

st.sidebar.markdown("---")
st.sidebar.markdown("📌 **NAVIGATION PAR NATURE DES TRAVAUX**")

# Liste des natures de travaux pour la navigation
liste_natures_nav = ["📌 Tous les travaux"] + list(LIAISONS.keys())
nature_selectionnee_sidebar = st.sidebar.radio(
    "Filtrer par nature :",
    options=liste_natures_nav,
    index=0
)

st.sidebar.markdown("---")
with st.sidebar.expander("➕ Ajouter un nouveau chantier/onglet", expanded=False):
    nouveau_projet_nom = st.text_input("Nom du projet :", placeholder="Ex: NGE ADM Lot 02...")
    if st.button("➕ Créer", use_container_width=True):
        if nouveau_projet_nom.strip():
            nom_clean = nouveau_projet_nom.strip()
            if nom_clean not in chantiers_existants:
                df_vide = pd.DataFrame(columns=COLUMNS_TEMPLATE)
                save_to_excel_with_formatting(df_vide, chemin_excel_defaut, sheet_name=nom_clean)
                st.success("Chantier créé !")
                st.rerun()

st.sidebar.markdown("---")
source_excel = st.sidebar.radio("Source de données :", ["Fichier Excel Local", "Téléverser Fichier"])

df = None
if source_excel == "Fichier Excel Local":
    if os.path.exists(chemin_excel_defaut):
        try:
            df = pd.read_excel(chemin_excel_defaut, sheet_name=chantier_actif).fillna("")
        except Exception as e:
            st.sidebar.error(f"Erreur : {e}")
else:
    fichier_upload = st.sidebar.file_uploader("Importer Excel", type=["xlsx", "xls"])
    if fichier_upload is not None:
        try:
            df = pd.read_excel(fichier_upload).fillna("")
            st.sidebar.success("✅ Importé !")
        except Exception as e:
            st.sidebar.error(f"Erreur : {e}")

st.sidebar.markdown("---")
if st.sidebar.button("✓ Valider le compte rendu", key="btn_validate", type="secondary"):
    st.toast("✅ Compte rendu validé !", icon="🎉")

# ==========================================
# 4. INTERFACE PRINCIPALE (WORKSPACE)
# ==========================================
col_title, col_actions = st.columns([3, 2])

with col_title:
    st.markdown(f"## **Chantier : {chantier_actif}**")
    st.markdown(f"<span style='color: #718096; font-weight: 500;'>Filtre actif : <b>{nature_selectionnee_sidebar}</b></span>", unsafe_allow_html=True)

with col_actions:
    st.markdown("""
    <div style="display: flex; gap: 8px; justify-content: flex-end; margin-top: 10px;">
        <span class="btn-action-purple">+ Titre</span>
        <span class="btn-action-blue">+ Tâche</span>
        <span class="btn-action-green">+ Note</span>
    </div>
    """, unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

if df is not None:
    tab1, tab2 = st.tabs(["📝 **Saisie d'Avancement**", "📊 **Suivi des Tâches & Exports (Word / PDF)**"])

    # TAB 1: FORMULAIRE DE SAISIE
    with tab1:
        st.markdown("#### ➕ **Nouvelle entrée d'avancement**")
        
        # Pre-select nature from sidebar if chosen
        default_idx = 0
        if nature_selectionnee_sidebar != "📌 Tous les travaux":
            try:
                default_idx = list(LIAISONS.keys()).index(nature_selectionnee_sidebar)
            except ValueError:
                default_idx = 0

        col1, col2 = st.columns(2)
        with col1:
            date_saisie = st.date_input("🗓️ Date d'intervention", value=datetime.today(), format="DD/MM/YYYY")
            nature_selectionnee = st.selectbox("🏗️ Nature des Travaux", options=list(LIAISONS.keys()), index=default_idx)
            info_liaison = LIAISONS.get(nature_selectionnee, {"procedure": "", "pieces": ""})
            partie_ouvrage = st.text_input("🧱 Partie d'ouvrage", placeholder="Ex: Hall RDC Bâtiment A...")
            situation = st.text_input("📍 Situation / PK", placeholder="Ex: Du PK 0+050 au PK 0+100")
        with col2:
            activite = st.text_area("🚜 Activité réalisée", height=80, placeholder="Ex: Peinture des plafonds...")
            essai = st.selectbox("🧪 Essai / Contrôle réalisé", options=[
                "Aucun", "ESSAI À LA PLAQUE", "DENSITÉ", "ESSAI À LA PLAQUE + DENSITÉ",
                "TENEUR EN EAU", "IDENTIFICATION DES MATERIAUX" , "PRELEVEMENT AVANT COMPACTAGE", "PRELEVEMENT APRES COMPACTAGE", "PRELEVEMENT"])
            procedure = st.text_input("📑 Référence de procédure", value=info_liaison["procedure"])
            pieces_jointes = st.text_area("📎 Pièces jointes", value=info_liaison["pieces"], height=100)

        if st.button("💾 Enregistrer la tâche", type="primary"):
            new_entry = {
                "DATE": date_saisie.strftime('%d/%m/%Y'),
                "TITRE DE LA NATURE DES TRAVAUX": nature_selectionnee,
                COL_PARTIE: partie_ouvrage,
                "SITUATION": situation,
                "ACTIVITÉ RÉALISÉE": activite,
                "ÉSSAI/ CONTRÔLE RÉALISÉE": None if essai == "Aucun" else essai,
                "RÉFÉRENCE DE PROCÉDURE": procedure,
                "PIÈCES JOINTES": pieces_jointes
            }
            df_updated = pd.concat([df, pd.DataFrame([new_entry])], ignore_index=True)
            if source_excel == "Fichier Excel Local":
                save_to_excel_with_formatting(df_updated, chemin_excel_defaut, sheet_name=chantier_actif)
            st.success("✅ Entrée enregistrée dans le projet !")
            st.rerun()

    # TAB 2: TABLEAU DES TÂCHES ET EXPORTS
    with tab2:
        st.markdown("#### 🔍 **Filtrer & Sélectionner les Tâches**")
        col_partie_name = COL_PARTIE if COL_PARTIE in df.columns else "PARTIE D meOUVRAGE"
        
        natures_uniques = sorted([str(x) for x in df['TITRE DE LA NATURE DES TRAVAUX'].unique() if str(x).strip()])
        parties_uniques = sorted([str(x) for x in df[col_partie_name].unique() if str(x).strip()])

        col_f1, col_f2 = st.columns(2)
        with col_f1:
            filter_nature = st.multiselect("📌 Nature des travaux (Filtre secondaire)", options=natures_uniques)
        with col_f2:
            filter_partie = st.multiselect("🧱 Partie d'Ouvrage", options=parties_uniques)

        search_text = st.text_input("⚡ Recherche rapide")

        df_filtered = df.copy()

        # Application du filtre principal de la Sidebar Navigation
        if nature_selectionnee_sidebar != "📌 Tous les travaux":
            df_filtered = df_filtered[df_filtered['TITRE DE LA NATURE DES TRAVAUX'] == nature_selectionnee_sidebar]

        # Application des filtres secondaires
        if filter_nature:
            df_filtered = df_filtered[df_filtered['TITRE DE LA NATURE DES TRAVAUX'].isin(filter_nature)]
        if filter_partie:
            df_filtered = df_filtered[df_filtered[col_partie_name].isin(filter_partie)]
        if search_text:
            mask = df_filtered.apply(lambda col: col.astype(str).str.contains(search_text, case=False, na=False)).any(axis=1)
            df_filtered = df_filtered[mask]

        df_editor = df_filtered.copy()
        if "Imprimer" not in df_editor.columns:
            df_editor.insert(0, "Imprimer", False)

        edited_df = st.data_editor(
            df_editor,
            num_rows="dynamic",
            height=380,
            use_container_width=True
        )

        lignes_selectionnees = edited_df[edited_df["Imprimer"] == True]
        nb_selections = len(lignes_selectionnees)

        st.markdown("---")
        col_act1, col_act2 = st.columns(2)

        with col_act1:
            if st.button("💾 Enregistrer les modifications dans Excel"):
                if source_excel == "Fichier Excel Local":
                    df.update(edited_df)
                    save_to_excel_with_formatting(df, chemin_excel_defaut, sheet_name=chantier_actif)
                    st.success("✅ Base Excel mise à jour !")
                    st.rerun()

        with col_act2:
            if nb_selections == 0:
                btn_title = "📄 Générer Word & PDF (Cochez au moins une ligne)"
            elif nb_selections == 1:
                btn_title = "📄 Générer Word & PDF (1 Fiche)"
            else:
                btn_title = f"📦 Générer Pack ZIP ({nb_selections} Fiches Word + PDF)"

            if st.button(btn_title, type="primary"):
                if nb_selections == 0:
                    st.warning("⚠️ Cochez la case 'Imprimer' pour au moins une ligne.")
                elif nb_selections == 1:
                    ligne_choisie = lignes_selectionnees.iloc[0]
                    nom_modele = str(ligne_choisie.get('TITRE DE LA NATURE DES TRAVAUX', '')).strip()
                    chemin_modele = trouver_modele_word(nom_modele)

                    if not chemin_modele:
                        st.error(f"❌ Modèle Word `{nom_modele}.docx` introuvable.")
                    else:
                        try:
                            with st.spinner("⏳ Génération des documents en cours..."):
                                val_partie = ligne_choisie.get(COL_PARTIE, ligne_choisie.get("PARTIE D meOUVRAGE", ''))
                                contexte = {
                                    'NATURE': str(ligne_choisie.get('TITRE DE LA NATURE DES TRAVAUX', '')),
                                    'REF': str(ligne_choisie.get('RÉFÉRENCE DE PROCÉDURE', '')),
                                    'PARTIE': str(val_partie),
                                    'SITUATION': str(ligne_choisie.get('SITUATION', '')),
                                    'PIECES': text_to_richtext(ligne_choisie.get('PIÈCES JOINTES', '')),
                                    'DATE': str(ligne_choisie.get('DATE', '')),
                                    'ACTIVITE': text_to_richtext(ligne_choisie.get('ACTIVITÉ RÉALISÉE', '')),
                                    'ESSAI': str(ligne_choisie.get('ÉSSAI/ CONTRÔLE RÉALISÉE', ''))
                                }
                                docx_bytes, pdf_bytes = generer_docx_et_pdf_bytes(chemin_modele, contexte)
                                nom_base = construire_nom_pdf(ligne_choisie).replace(".pdf", "")

                                st.success("✅ Fiches Word & PDF générées avec succès !")
                                c_down1, c_down2 = st.columns(2)
                                with c_down1:
                                    st.download_button("📝 Télécharger WORD (.docx)", data=docx_bytes, file_name=f"{nom_base}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
                                with c_down2:
                                    st.download_button("📕 Télécharger PDF (.pdf)", data=pdf_bytes, file_name=f"{nom_base}.pdf", mime="application/pdf", use_container_width=True)
                        except Exception as e:
                            st.error(f"❌ Erreur : {e}")
                else:
                    zip_buffer = io.BytesIO()
                    fichiers_crees = 0

                    with st.spinner(f"⏳ Génération du Pack ZIP ({nb_selections} fiches)..."):
                        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                            for idx, row in lignes_selectionnees.iterrows():
                                nom_modele = str(row.get('TITRE DE LA NATURE DES TRAVAUX', '')).strip()
                                chemin_modele = trouver_modele_word(nom_modele)
                                if not chemin_modele:
                                    continue
                                try:
                                    val_partie = row.get(COL_PARTIE, row.get("PARTIE D meOUVRAGE", ''))
                                    contexte = {
                                        'NATURE': str(row.get('TITRE DE LA NATURE DES TRAVAUX', '')),
                                        'REF': str(row.get('RÉFÉRENCE DE PROCÉDURE', '')),
                                        'PARTIE': str(val_partie),
                                        'SITUATION': str(row.get('SITUATION', '')),
                                        'PIECES': text_to_richtext(row.get('PIÈCES JOINTES', '')),
                                        'DATE': str(row.get('DATE', '')),
                                        'ACTIVITE': text_to_richtext(row.get('ACTIVITÉ RÉALISÉE', '')),
                                        'ESSAI': str(row.get('ÉSSAI/ CONTRÔLE RÉALISÉE', ''))
                                    }
                                    docx_bytes, pdf_bytes = generer_docx_et_pdf_bytes(chemin_modele, contexte)
                                    nom_base = construire_nom_pdf(row).replace(".pdf", "")

                                    zip_file.writestr(f"{nom_base}.docx", docx_bytes)
                                    zip_file.writestr(f"{nom_base}.pdf", pdf_bytes)
                                    fichiers_crees += 1
                                except Exception:
                                    pass

                    if fichiers_crees > 0:
                        zip_buffer.seek(0)
                        st.success(f"✅ {fichiers_crees} fiches (Word & PDF) prêtes !")
                        st.download_button(
                            label=f"📦 Télécharger le Pack ZIP",
                            data=zip_buffer,
                            file_name=f"Fiches_Batiscript_{datetime.now().strftime('%Y%m%d_%H%M')}.zip",
                            mime="application/zip",
                            use_container_width=True
                        )
else:
    st.info("💡 Veuillez charger un fichier Excel pour commencer.")
