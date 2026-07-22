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
# 0. CONFIGURATION DE LA PAGE ET STYLES CSS
# ==========================================
st.set_page_config(
    page_title="Suivi Chantier - Génie Civil & Routes",
    page_icon="🏗️",
    layout="wide"
)

# 🎨 Injecting Custom Civil Engineering / Highway CSS Theme
st.markdown("""
<style>
    /* Background General */
    .stApp {
        background-color: #f8fafc;
    }
    
    /* Main Banner Header */
    .gc-header {
        background: linear-gradient(135deg, #0f172a 0%, #1e293b 100%);
        color: #ffffff;
        padding: 22px 28px;
        border-radius: 12px;
        border-left: 8px solid #ff6b00; /* Safety Orange Accent */
        box-shadow: 0px 4px 12px rgba(0, 0, 0, 0.08);
        margin-bottom: 25px;
    }
    .gc-header h1 {
        color: #ffffff !important;
        font-size: 26px !important;
        font-weight: 800 !important;
        margin: 0 !important;
        padding: 0 !important;
        letter-spacing: 0.5px;
    }
    .gc-header p {
        color: #94a3b8;
        margin: 6px 0 0 0;
        font-size: 14px;
    }

    /* KPI Cards Styling */
    .kpi-card {
        background-color: #ffffff;
        border: 1px solid #e2e8f0;
        border-top: 4px solid #ff6b00;
        border-radius: 10px;
        padding: 16px;
        text-align: center;
        box-shadow: 0 2px 6px rgba(0,0,0,0.03);
    }
    .kpi-value {
        font-size: 24px;
        font-weight: 800;
        color: #0f172a;
    }
    .kpi-label {
        font-size: 11px;
        font-weight: 600;
        color: #64748b;
        text-transform: uppercase;
        letter-spacing: 0.5px;
        margin-top: 4px;
    }

    /* Customizing Buttons */
    .stButton > button[kind="primary"] {
        background-color: #ff6b00 !important;
        color: #ffffff !important;
        border: none !important;
        border-radius: 8px !important;
        font-weight: 700 !important;
        padding: 10px 20px !important;
        transition: all 0.2s ease-in-out !important;
    }
    .stButton > button[kind="primary"]:hover {
        background-color: #e05e00 !important;
        box-shadow: 0 4px 12px rgba(255, 107, 0, 0.25) !important;
    }

    /* Sidebar Customization */
    section[data-testid="stSidebar"] {
        background-color: #0f172a !important;
        color: #ffffff !important;
    }
    section[data-testid="stSidebar"] .stMarkdown h1, 
    section[data-testid="stSidebar"] .stMarkdown h2, 
    section[data-testid="stSidebar"] .stMarkdown h3,
    section[data-testid="stSidebar"] label {
        color: #f1f5f9 !important;
    }

    /* Tab Design */
    .stTabs [data-baseweb="tab-list"] {
        gap: 10px;
    }
    .stTabs [data-baseweb="tab"] {
        background-color: #ffffff;
        border-radius: 8px;
        padding: 10px 20px;
        border: 1px solid #e2e8f0;
        font-weight: 700;
        color: #475569;
    }
    .stTabs [aria-selected="true"] {
        background-color: #ff6b00 !important;
        color: #ffffff !important;
        border-color: #ff6b00 !important;
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
            raise RuntimeError("Impossible de convertir en PDF (MS Word ou LibreOffice requis).")

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
# 2. BARRE LATÉRALE (SIDEBAR)
# ==========================================
st.sidebar.markdown("### 🏗️ **Gestion de Chantier**")

chantiers_existants = get_sheet_names(chemin_excel_defaut)
chantier_actif = st.sidebar.selectbox("📌 **Projet / Tronçon Actif :**", options=chantiers_existants)

with st.sidebar.expander("➕ Nouveau Projet", expanded=False):
    nouveau_projet_nom = st.text_input("Nom du projet :", placeholder="Ex: Autoroute PK 12, Viaduc...")
    if st.button("➕ Créer le Projet", use_container_width=True):
        if nouveau_projet_nom.strip():
            nom_clean = nouveau_projet_nom.strip()
            if nom_clean not in chantiers_existants:
                df_vide = pd.DataFrame(columns=COLUMNS_TEMPLATE)
                save_to_excel_with_formatting(df_vide, chemin_excel_defaut, sheet_name=nom_clean)
                st.success(f"Projet '{nom_clean}' créé !")
                st.rerun()

st.sidebar.markdown("---")
st.sidebar.markdown("### 📂 **Source des Données**")
source_excel = st.sidebar.radio(
    "Source :",
    ["Fichier système (suivi.xlsx)", "Téléverser un autre fichier Excel"]
)

df = None
if source_excel == "Fichier système (suivi.xlsx)":
    if os.path.exists(chemin_excel_defaut):
        try:
            df = pd.read_excel(chemin_excel_defaut, sheet_name=chantier_actif).fillna("")
            st.sidebar.success(f"✅ Projet '{chantier_actif}' chargé !")
        except Exception as e:
            st.sidebar.error(f"❌ Erreur de lecture : {e}")
    else:
        st.sidebar.error("❌ Fichier Excel introuvable.")
else:
    fichier_upload = st.sidebar.file_uploader("Fichier Excel (.xlsx)", type=["xlsx", "xls"])
    if fichier_upload is not None:
        try:
            df = pd.read_excel(fichier_upload).fillna("")
            st.sidebar.success("✅ Importation réussie !")
        except Exception as e:
            st.sidebar.error(f"❌ Erreur : {e}")

# ==========================================
# 3. INTERFACE PRINCIPALE
# ==========================================

# Banner Title
st.markdown(f"""
<div class="gc-header">
    <h1>🛣️ Plateforme Génie Civil & Travaux Routiers</h1>
    <p>Gestion des suivi de travaux, fiches de contrôle & édition automatique des documents Word/PDF | Projet : <b>{chantier_actif}</b></p>
</div>
""", unsafe_allow_html=True)

if df is not None:
    # 📊 KPI Cards Section
    k1, k2, k3, k4 = st.columns(4)
    with k1:
        st.markdown(f'<div class="kpi-card"><div class="kpi-value">{len(df)}</div><div class="kpi-label">📝 Fiches Enregistrées</div></div>', unsafe_allow_html=True)
    with k2:
        nb_natures = df['TITRE DE LA NATURE DES TRAVAUX'].nunique() if 'TITRE DE LA NATURE DES TRAVAUX' in df.columns else 0
        st.markdown(f'<div class="kpi-card"><div class="kpi-value">{nb_natures}</div><div class="kpi-label">🚜 Natures de Travaux</div></div>', unsafe_allow_html=True)
    with k3:
        col_p = COL_PARTIE if COL_PARTIE in df.columns else "PARTIE D meOUVRAGE"
        nb_parties = df[col_p].nunique() if col_p in df.columns else 0
        st.markdown(f'<div class="kpi-card"><div class="kpi-value">{nb_parties}</div><div class="kpi-label">🧱 Parties d\'Ouvrage</div></div>', unsafe_allow_html=True)
    with k4:
        st.markdown(f'<div class="kpi-card"><div class="kpi-value" style="color:#ff6b00;">Active</div><div class="kpi-label">⚙️ État du Système</div></div>', unsafe_allow_html=True)

    st.write("")

    tab1, tab2 = st.tabs(["📝 **Nouvelle Saisie Chantier**", "📊 **Registre & Génération Word / PDF**"])

    # -------------------------------------------------------------
    # TAB 1 : SAISIE
    # -------------------------------------------------------------
    with tab1:
        st.markdown("##### 👷 **Ajouter une nouvelle fiche de contrôle / suivi**")
        col1, col2 = st.columns(2)
        with col1:
            date_saisie = st.date_input("🗓️ Date des Travaux", value=datetime.today(), format="DD/MM/YYYY")
            nature_selectionnee = st.selectbox("📌 Nature des travaux", options=list(LIAISONS.keys()))
            info_liaison = LIAISONS.get(nature_selectionnee, {"procedure": "", "pieces": ""})
            partie_ouvrage = st.text_input("🧱 Partie d'ouvrage", placeholder="Ex: Bretelle B, Tranchée 1...")
            situation = st.text_input("📍 Situation / PK", placeholder="Ex: PK 12+400 au PK 12+800")
        with col2:
            activite = st.text_area("🚜 Activité réalisée", height=80, placeholder="Ex: Réalisation de la couche de forme...")
            essai = st.selectbox("🧪 Essai / Contrôle réalisé", options=[
                "Aucun", "ESSAI À LA PLAQUE", "DENSITÉ", "ESSAI À LA PLAQUE + DENSITÉ",
                "TENEUR EN EAU", "IDENTIFICATION DES MATERIAUX" , "PRELEVEMENT AVANT COMPACTAGE", "PRELEVEMENT APRES COMPACTAGE", "PRELEVEMENT"])
            procedure = st.text_input("📑 Référence procédure", value=info_liaison["procedure"])
            pieces_jointes = st.text_area("📎 Pièces jointes", value=info_liaison["pieces"], height=100)

        if st.button("💾 Enregistrer la Fiche", type="primary"):
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
            if source_excel == "Fichier système (suivi.xlsx)":
                save_to_excel_with_formatting(df_updated, chemin_excel_defaut, sheet_name=chantier_actif)
            st.success("✅ Fiche enregistrée avec succès !")
            st.rerun()

    # -------------------------------------------------------------
    # TAB 2 : REGISTRE & GENERATION
    # -------------------------------------------------------------
    with tab2:
        st.markdown("##### 🔍 **Registre des Travaux & Exportation**")

        with st.expander("🔻 **Filtres de Recherche Avancés**", expanded=False):
            col_partie_name = COL_PARTIE if COL_PARTIE in df.columns else "PARTIE D meOUVRAGE"
            
            natures_uniques = sorted([str(x) for x in df['TITRE DE LA NATURE DES TRAVAUX'].unique() if str(x).strip()])
            parties_uniques = sorted([str(x) for x in df[col_partie_name].unique() if str(x).strip()])

            cf1, cf2 = st.columns(2)
            with cf1:
                filter_nature = st.multiselect("📌 Filtrer par Nature", options=natures_uniques)
            with cf2:
                filter_partie = st.multiselect("🧱 Filtrer par Partie d'Ouvrage", options=parties_uniques)

            search_text = st.text_input("⚡ Recherche rapide par mot-clé")

        # Application Filtres
        df_filtered = df.copy()
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
            height=360,
            use_container_width=True
        )

        lignes_selectionnees = edited_df[edited_df["Imprimer"] == True]
        nb_selections = len(lignes_selectionnees)

        st.markdown("---")
        col_act1, col_act2 = st.columns(2)

        with col_act1:
            if st.button("💾 Enregistrer les modifications Excel", type="secondary", use_container_width=True):
                if source_excel == "Fichier système (suivi.xlsx)":
                    df.update(edited_df)
                    save_to_excel_with_formatting(df, chemin_excel_defaut, sheet_name=chantier_actif)
                    st.success("✅ Base Excel mise à jour !")
                    st.rerun()

        with col_act2:
            if nb_selections == 0:
                btn_title = "📄 Générer Word & PDF (Cochez une ligne)"
            elif nb_selections == 1:
                btn_title = f"📄 Générer Word & PDF (1 Fiche)"
            else:
                btn_title = f"📦 Pack ZIP : {nb_selections} Fiches (Word + PDF)"

            if st.button(btn_title, type="primary", use_container_width=True):
                if nb_selections == 0:
                    st.warning("⚠️ Veuillez cocher la case 'Imprimer' d'au moins une ligne dans le tableau !")
                
                elif nb_selections == 1:
                    ligne_choisie = lignes_selectionnees.iloc[0]
                    nom_modele = str(ligne_choisie.get('TITRE DE LA NATURE DES TRAVAUX', '')).strip()
                    chemin_modele = trouver_modele_word(nom_modele)

                    if not chemin_modele:
                        st.error(f"❌ Le modèle Word `{nom_modele}.docx` est introuvable.")
                    else:
                        try:
                            with st.spinner("⏳ Génération Word & PDF..."):
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

                                st.success("✅ Fiches générées !")
                                c_down1, c_down2 = st.columns(2)
                                with c_down1:
                                    st.download_button("📝 Télécharger WORD", data=docx_bytes, file_name=f"{nom_base}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
                                with c_down2:
                                    st.download_button("📕 Télécharger PDF", data=pdf_bytes, file_name=f"{nom_base}.pdf", mime="application/pdf", use_container_width=True)
                        except Exception as e:
                            st.error(f"❌ Erreur : {e}")

                else:
                    zip_buffer = io.BytesIO()
                    fichiers_crees = 0

                    with st.spinner(f"⏳ Génération du Pack ({nb_selections} fiches)..."):
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
                        st.success(f"✅ Pack prêt ({fichiers_crees * 2} fichiers) !")
                        st.download_button(
                            label=f"📦 Télécharger le Pack ZIP",
                            data=zip_buffer,
                            file_name=f"Fiches_Chantier_{datetime.now().strftime('%Y%m%d_%H%M')}.zip",
                            mime="application/zip",
                            use_container_width=True
                        )
else:
    st.info("💡 Veuillez charger un fichier Excel pour commencer.")