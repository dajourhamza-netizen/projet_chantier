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
import hashlib
import gspread
from google.oauth2.service_account import Credentials

# Bibliothèques Word
from docx import Document
from docxtpl import DocxTemplate, RichText

# ==========================================
# 0. CONFIGURATION ET CONSTANTES GLOBALES
# ==========================================
st.set_page_config(
    page_title="Suivi Chantier - Génie Civil",
    page_icon="🏗️",
    layout="wide",
    initial_sidebar_state="collapsed"
)

DOSSIER_CHANTIER = os.path.dirname(os.path.abspath(__file__)) if '__file__' in globals() else os.getcwd()
COL_PARTIE = "PARTIE D'OUVRAGE"

COLUMNS_TEMPLATE = [
    "DATE", "TITRE DE LA NATURE DES TRAVAUX", COL_PARTIE, 
    "SITUATION", "ACTIVITÉ RÉALISÉE", "ÉSSAI/ CONTRÔLE RÉALISÉE", 
    "RÉFÉRENCE DE PROCÉDURE", "PIÈCES JOINTES", "CRÉÉ PAR"
]

USER_COLUMNS = ["username", "password", "role", "actif"]

LIAISONS = {
    "ARASE DE PST": {
        "procedure": "TER-PEX-05-00", 
        "pieces": "* Fiche de suivi de la PST\n* Fiche de réception topographique\n* PVs laboratoire"
    },
    "ARASE DE TERRASSEMENT": {
        "procedure": "TER-PEX-03-00", 
        "pieces": "* Fiche de contrôle des déblais\n* Fiche de réception topographique\n* PVs laboratoire"
    },
    "ASSISE DE REMBLAIS PURGE": {
        "procedure": "TER-PEX-04-00", 
        "pieces": "* Fiche de réception de l'assise des remblais\n* Fiche de réception topographique\n* Fiche d'identification de la purge\n* PVs laboratoire"
    },
    "ASSISE DE REMBLAIS": {
        "procedure": "TER-PEX-04-00", 
        "pieces": "* Fiche de réception de l'assise des remblais\n* Fiche de réception topographique\n* PVs laboratoire"
    },
    "COUCHE DE FORME": {
        "procedure": "TER-PEX-09-00", 
        "pieces": "* Fiche de suivi et de contrôle de la CDF\n* Fiche de réception topographique\n* PVs laboratoire"
    },
    "DÉCAPAGE": {
        "procedure": "TER-PEX-02-00", 
        "pieces": "* Fiche de suivi et de contrôle du décapage\n* Fiche des sections à décaper\n* Fiche de réception topographique"
    },
    "DEGAGEMENT D'EMPRISE": {
        "procedure": "TER-PEX-01-00", 
        "pieces": "* Fiche de suivi et de contrôle du dégagement des emprises\n* Fiche de réception topographique\n* Constat dégagement d'emprise"
    },
    "REMBLAIS": {
        "procedure": "TER-PEX-04-00", 
        "pieces": "* Fiche de suivi et de contrôle des remblais\n* PVs laboratoire"
    },
    "REMBLAIS PST": {
        "procedure": "TER-PEX-05-00", 
        "pieces": "* Fiche de suivi et de contrôle des remblais PST\n* PVs laboratoire"
    }
}

# ==========================================
# 1. STYLES CSS RESPONSIVES ET TOUCH-FRIENDLY
# ==========================================
st.markdown("""
<style>
    .stApp {
        background-image: linear-gradient(rgba(15, 23, 42, 0.88), rgba(15, 23, 42, 0.88)), 
                          url("https://i.pinimg.com/736x/3d/6b/f7/3d6bf78abc63f1c9b000d4bc5fbe7fa3.jpg");
        background-size: cover;
        background-position: center;
        background-repeat: no-repeat;
        background-attachment: fixed;
    }

    [data-testid="stHeader"] { background-color: rgba(0, 0, 0, 0); }

    .gc-header {
        background: linear-gradient(135deg, #0f172a 0%, #1e293b 100%);
        color: #ffffff;
        padding: 16px 20px;
        border-radius: 12px;
        border-left: 6px solid #ff6b00;
        box-shadow: 0px 4px 12px rgba(0, 0, 0, 0.25);
        margin-bottom: 15px;
    }
    .gc-header h1 { color: #ffffff !important; font-size: 22px !important; font-weight: 800 !important; margin: 0 !important; }
    .gc-header p { color: #94a3b8; margin: 4px 0 0 0; font-size: 13px; }
    
    .stButton > button {
        min-height: 48px !important;
        font-size: 15px !important;
        font-weight: 700 !important;
        border-radius: 10px !important;
        width: 100% !important;
    }

    .stButton > button[kind="primary"] {
        background-color: #ff6b00 !important; 
        color: #ffffff !important; 
        border: none !important; 
    }
    
    section[data-testid="stSidebar"] { 
        background-color: rgba(15, 23, 42, 0.98) !important; 
        color: #ffffff !important; 
    }
    section[data-testid="stSidebar"] label, section[data-testid="stSidebar"] .stMarkdown h1 { 
        color: #f1f5f9 !important; 
    }

    @media (max-width: 768px) {
        .block-container {
            padding-top: 1rem !important;
            padding-bottom: 2rem !important;
            padding-left: 0.8rem !important;
            padding-right: 0.8rem !important;
        }

        .gc-header { padding: 12px 14px !important; }
        .gc-header h1 { font-size: 18px !important; }
        .gc-header p { font-size: 11px !important; }

        button[data-baseweb="tab"] {
            font-size: 13px !important;
            padding: 8px 10px !important;
        }

        div[data-testid="stDataFrame"] {
            overflow-x: auto !important;
            -webkit-overflow-scrolling: touch;
        }
    }
</style>
""", unsafe_allow_html=True)

# ==========================================
# 2. SÉCURITÉ ET HACHAGE
# ==========================================
def hash_password(password):
    return hashlib.sha256(str(password).encode('utf-8')).hexdigest()

# ==========================================
# 3. INTERACTION GOOGLE SHEETS
# ==========================================
@st.cache_resource
def get_gsheets_client():
    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive"
    ]
    creds_dict = dict(st.secrets["gcp_service_account"])
    if "private_key" in creds_dict:
        creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
    creds = Credentials.from_service_account_info(creds_dict, scopes=scope)
    return gspread.authorize(creds)

def get_spreadsheet():
    client = get_gsheets_client()
    url = st.secrets["gsheets"]["spreadsheet_url"]
    return client.open_by_url(url)

def load_users():
    try:
        sh = get_spreadsheet()
        try:
            ws = sh.worksheet("Utilisateurs")
        except gspread.WorksheetNotFound:
            ws = sh.add_worksheet(title="Utilisateurs", rows=50, cols=10)
            admin_initial = pd.DataFrame([{
                "username": "admin",
                "password": hash_password("admin123"),
                "role": "Admin",
                "actif": "OUI"
            }])
            ws.update([admin_initial.columns.values.tolist()] + admin_initial.astype(str).values.tolist())
            return admin_initial

        records = ws.get_all_records()
        return pd.DataFrame(records)
    except Exception as e:
        st.error(f"Erreur lors du chargement des utilisateurs : {e}")
        return pd.DataFrame(columns=USER_COLUMNS)

def save_users(df_users):
    try:
        sh = get_spreadsheet()
        try:
            ws = sh.worksheet("Utilisateurs")
        except gspread.WorksheetNotFound:
            ws = sh.add_worksheet(title="Utilisateurs", rows=50, cols=10)
        
        ws.clear()
        values = [df_users.columns.values.tolist()] + df_users.astype(str).values.tolist()
        ws.update(values)
        return True, "✅ Utilisateurs mis à jour !"
    except Exception as e:
        return False, f"❌ Erreur : {e}"

def get_sheet_names_gsheets():
    try:
        sh = get_spreadsheet()
        return [ws.title for ws in sh.worksheets() if ws.title != "Utilisateurs"]
    except Exception as e:
        st.error(f"Erreur de connexion à Google Sheets : {e}")
        return ["Chantier Principal"]

def load_data_from_sheet(sheet_name):
    try:
        sh = get_spreadsheet()
        ws = sh.worksheet(sheet_name)
        records = ws.get_all_records()
        df_loaded = pd.DataFrame(records)
        if "CRÉÉ PAR" not in df_loaded.columns and not df_loaded.empty:
            df_loaded["CRÉÉ PAR"] = ""
        return df_loaded
    except Exception:
        return pd.DataFrame(columns=COLUMNS_TEMPLATE)

def save_data_to_sheet(df_to_save, sheet_name):
    try:
        sh = get_spreadsheet()
        try:
            ws = sh.worksheet(sheet_name)
        except gspread.WorksheetNotFound:
            ws = sh.add_worksheet(title=sheet_name, rows=500, cols=20)

        df_clean = df_to_save.copy()
        if "Imprimer" in df_clean.columns:
            df_clean = df_clean.drop(columns=["Imprimer"])

        ws.clear()
        values = [df_clean.columns.values.tolist()] + df_clean.astype(str).values.tolist()
        ws.update(values)
        return True, "✅ Données enregistrées dans Google Sheets !"
    except Exception as e:
        return False, f"❌ Erreur d'enregistrement : {e}"

# ==========================================
# 4. FONCTIONS DE GÉNÉRATION DOCX & PDF
# ==========================================
def text_to_richtext(text):
    if not text or pd.isna(text): return ""
    rt = RichText()
    lines = str(text).split('\n')
    for i, line in enumerate(lines):
        rt.add(line)
        if i < len(lines) - 1: rt.add('\n')
    return rt

def clean_filename(text):
    if not text: return ""
    text = str(text)
    if text.lower().endswith('.docx'): text = text[:-5]
    text = unicodedata.normalize('NFD', text).encode('ascii', 'ignore').decode("utf-8")
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
    nom_propre = re.sub(r'[\\/*?:"<>|]', "_", f"{nature} - {partie} - {situation}")
    return f"{re.sub(r'\s+', ' ', nom_propre).strip()}.pdf"

def get_col_val(row, *candidates):
    for c in candidates:
        for col in row.index:
            if str(col).strip().lower() == str(c).strip().lower():
                val = row[col]
                if isinstance(val, (datetime, pd.Timestamp)):
                    return val.strftime('%d/%m/%Y')
                val_str = str(val).strip()
                if val_str and val_str.lower() != "nan":
                    return val_str
    return ""

def convertir_docx_vers_pdf_bytes(docx_path, temp_dir):
    base_name = os.path.splitext(os.path.basename(docx_path))[0]
    expected_pdf = os.path.join(temp_dir, f"{base_name}.pdf")

    try:
        cmd = f'soffice --headless --convert-to pdf "{docx_path}" --outdir "{temp_dir}"'
        subprocess.run(cmd, shell=True, check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        if os.path.exists(expected_pdf):
            with open(expected_pdf, "rb") as f:
                return f.read()
    except Exception:
        pass

    try:
        from docx2pdf import convert
        convert(docx_path, expected_pdf)
        if os.path.exists(expected_pdf):
            with open(expected_pdf, "rb") as f:
                return f.read()
    except Exception:
        pass

    return None

def generer_docx_et_pdf_bytes(chemin_modele, contexte):
    with tempfile.TemporaryDirectory() as temp_dir:
        doc = DocxTemplate(chemin_modele)
        doc.render(contexte)
        docx_temp_path = os.path.join(temp_dir, "temp.docx")
        doc.save(docx_temp_path)
        
        with open(docx_temp_path, "rb") as f: docx_bytes = f.read()
        pdf_bytes = convertir_docx_vers_pdf_bytes(docx_temp_path, temp_dir)
        if pdf_bytes is None: pdf_bytes = docx_bytes

        return docx_bytes, pdf_bytes

def generer_di_style_vba(chemin_modele, df_jour):
    doc = Document(chemin_modele)
    word_table = None
    for tbl in doc.tables:
        if len(tbl.columns) >= 4:
            word_table = tbl
            break
            
    if not word_table: raise ValueError("Tableau introuvable dans le fichier Word.")
    
    for idx, (_, row) in enumerate(df_jour.iterrows()):
        date_val = get_col_val(row, "DATE")
        nature_val = get_col_val(row, "TITRE DE LA NATURE DES TRAVAUX", "NATURE")
        partie_val = get_col_val(row, "PARTIE D'OUVRAGE", "PARTIE D meOUVRAGE", "PARTIE")
        situation_val = get_col_val(row, "SITUATION", "PK")
        activite_val = get_col_val(row, "ACTIVITÉ RÉALISÉE", "ACTIVITE")
        essai_val = get_col_val(row, "ÉSSAI/ CONTRÔLE RÉALISÉE", "ESSAI")

        col2_text = f"{activite_val} - {nature_val}" if (activite_val and nature_val) else (activite_val or nature_val)
        col3_text = f"{partie_val} / {situation_val}" if (partie_val and situation_val) else (partie_val or situation_val)
        
        target_row_idx = idx + 1
        if target_row_idx < len(word_table.rows):
            row_cells = word_table.rows[target_row_idx].cells
        else:
            row_cells = word_table.add_row().cells

        row_cells[0].text = date_val
        row_cells[1].text = col2_text
        row_cells[2].text = col3_text
        row_cells[3].text = essai_val
        if len(row_cells) >= 5: row_cells[4].text = ""

    return doc

def generer_di_une_date(df_jour):
    modele_di = None
    if os.path.exists(DOSSIER_CHANTIER):
        for file in os.listdir(DOSSIER_CHANTIER):
            if file.lower().endswith('.docx') and not file.startswith('~$'):
                if 'di' in file.lower() or 'demande' in file.lower():
                    modele_di = os.path.join(DOSSIER_CHANTIER, file)
                    break

    with tempfile.TemporaryDirectory() as temp_dir:
        docx_temp_path = os.path.join(temp_dir, "di_single.docx")

        if modele_di and os.path.exists(modele_di):
            try:
                doc = generer_di_style_vba(modele_di, df_jour)
                doc.save(docx_temp_path)
            except Exception:
                doc = Document()
                doc.add_heading("Demande d'Intervention (DI)", 0)
                table = doc.add_table(rows=1, cols=4)
                for _, row in df_jour.iterrows():
                    row_cells = table.add_row().cells
                    row_cells[0].text = str(get_col_val(row, "DATE"))
                    row_cells[1].text = f"{get_col_val(row, 'ACTIVITÉ RÉALISÉE')} - {get_col_val(row, 'TITRE DE LA NATURE DES TRAVAUX')}"
                    row_cells[2].text = f"{get_col_val(row, COL_PARTIE)} / {get_col_val(row, 'SITUATION')}"
                    row_cells[3].text = str(get_col_val(row, "ÉSSAI/ CONTRÔLE RÉALISÉE"))
                doc.save(docx_temp_path)
        else:
            doc = Document()
            doc.add_heading("Demande d'Intervention (DI)", 0)
            table = doc.add_table(rows=1, cols=4)
            for _, row in df_jour.iterrows():
                row_cells = table.add_row().cells
                row_cells[0].text = str(get_col_val(row, "DATE"))
                row_cells[1].text = f"{get_col_val(row, 'ACTIVITÉ RÉALISÉE')} - {get_col_val(row, 'TITRE DE LA NATURE DES TRAVAUX')}"
                row_cells[2].text = f"{get_col_val(row, COL_PARTIE)} / {get_col_val(row, 'SITUATION')}"
                row_cells[3].text = str(get_col_val(row, "ÉSSAI/ CONTRÔLE RÉALISÉE"))
            doc.save(docx_temp_path)

        with open(docx_temp_path, "rb") as f: docx_bytes = f.read()
        pdf_bytes = convertir_docx_vers_pdf_bytes(docx_temp_path, temp_dir)
        if pdf_bytes is None: pdf_bytes = docx_bytes

        return docx_bytes, pdf_bytes

def generer_pack_di_zip(df_filtered):
    zip_buffer = io.BytesIO()
    dates_uniques = df_filtered["DATE"].unique()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        for date_val in dates_uniques:
            if not date_val or str(date_val).strip() == "" or str(date_val).lower() == "nan": continue
            df_jour = df_filtered[df_filtered["DATE"] == date_val]
            docx_b, pdf_b = generer_di_une_date(df_jour)
            date_clean = str(date_val).replace('/', '-').replace('\\', '-')
            zip_file.writestr(f"DI_{date_clean}.docx", docx_b)
            zip_file.writestr(f"DI_{date_clean}.pdf", pdf_b)
    zip_buffer.seek(0)
    return zip_buffer, len(dates_uniques)

# ==========================================
# 5. GESTION DE L'AUTHENTIFICATION
# ==========================================
if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False
    st.session_state["username"] = None
    st.session_state["role"] = None

def page_connexion():
    st.markdown("""
    <div style="max-width: 400px; margin: 30px auto; padding: 20px; background: rgba(15, 23, 42, 0.95); border-radius: 12px; border: 1px solid #334155; box-shadow: 0 4px 15px rgba(0,0,0,0.5);">
        <h3 style="text-align: center; color: #ffffff; margin-bottom: 15px;">🔒 Connexion Chantier</h3>
    """, unsafe_allow_html=True)
    
    with st.form("login_form"):
        username_input = st.text_input("Nom d'utilisateur").strip()
        password_input = st.text_input("Mot de passe", type="password").strip()
        btn_submit = st.form_submit_button("Se Connecter", type="primary", use_container_width=True)

        if btn_submit:
            if not username_input or not password_input:
                st.error("⚠️ Veuillez remplir tous les champs.")
            else:
                users_df = load_users()
                hashed_input = hash_password(password_input)
                
                user_row = users_df[
                    (users_df["username"].astype(str) == username_input) & 
                    (users_df["password"].astype(str) == hashed_input)
                ]

                if not user_row.empty:
                    info_user = user_row.iloc[0]
                    if str(info_user.get("actif", "OUI")).upper() in ["OUI", "TRUE", "1"]:
                        st.session_state["authenticated"] = True
                        st.session_state["username"] = info_user["username"]
                        st.session_state["role"] = info_user["role"]
                        st.success("Connexion réussie !")
                        st.rerun()
                    else:
                        st.error("🚫 Compte désactivé.")
                else:
                    st.error("❌ Identifiants invalides.")
    st.markdown("</div>", unsafe_allow_html=True)

if not st.session_state["authenticated"]:
    page_connexion()
    st.stop()

# ==========================================
# 6. BARRE LATÉRALE & GESTION DES PROJETS
# ==========================================
st.sidebar.markdown(f"👤 **{st.session_state['username']}** ({st.session_state['role']})")

if st.sidebar.button("🚪 Déconnexion", use_container_width=True):
    st.session_state["authenticated"] = False
    st.session_state["username"] = None
    st.session_state["role"] = None
    st.rerun()

st.sidebar.markdown("---")
st.sidebar.markdown("### 🌐 **Google Sheets**")

chantiers_existants = get_sheet_names_gsheets()
chantier_actif = st.sidebar.selectbox("📌 **Projet Actif :**", options=chantiers_existants)

if st.session_state["role"] == "Admin":
    with st.sidebar.expander("➕ **Nouveau Projet**", expanded=False):
        nouveau_projet_nom = st.text_input("Nom du projet :", key="new_proj_input")
        if st.button("✨ Créer Projet", type="primary", key="btn_create_proj", use_container_width=True):
            nom_clean = nouveau_projet_nom.strip()
            if nom_clean:
                if nom_clean in chantiers_existants:
                    st.sidebar.error("⚠️ Ce projet existe déjà !")
                else:
                    df_vide = pd.DataFrame(columns=COLUMNS_TEMPLATE)
                    success, msg = save_data_to_sheet(df_vide, sheet_name=nom_clean)
                    if success:
                        st.sidebar.success(f"✅ Projet '{nom_clean}' créé !")
                        st.rerun()
                    else:
                        st.sidebar.error(msg)

df = load_data_from_sheet(chantier_actif)

# ==========================================
# 7. INTERFACE PRINCIPALE (ONGLETS)
# ==========================================
st.markdown(f"""
<div class="gc-header">
    <h1>🛣️ Suivi Génie Civil</h1>
    <p>Projet : <b>{chantier_actif}</b></p>
</div>
""", unsafe_allow_html=True)

liste_onglets = [
    "📝 **Saisie**", 
    "📊 **Registre**", 
    "📅 **DI**"
]

if st.session_state["role"] == "Admin":
    liste_onglets.append("👥 **Accès**")

tabs = st.tabs(liste_onglets)
tab1, tab2, tab3 = tabs[0], tabs[1], tabs[2]

# -------------------------------------------------------------
# TAB 1 : SAISIE DES DONNÉES (OPTION DE NOUVEAUX ÉLÉMENTS DYNAMIQUE)
# -------------------------------------------------------------
with tab1:
    st.markdown("##### 👷 **Ajouter une fiche**")
    
    # 1. Préparation dynamique de la liste "Nature des travaux" (Base + Historique BDD)
    natures_bdd = sorted(list(set([str(n).strip() for n in df["TITRE DE LA NATURE DES TRAVAUX"].unique() if str(n).strip() and str(n).lower() != 'nan']))) if ("TITRE DE LA NATURE DES TRAVAUX" in df.columns and not df.empty) else []
    all_natures = sorted(list(set(list(LIAISONS.keys()) + natures_bdd)))
    options_nature = all_natures + ["➕ Autre / Nouvelle nature..."]

    # 2. Préparation dynamique de la liste "Partie d'ouvrage"
    parties_existantes = sorted(list(set([str(p).strip() for p in df[COL_PARTIE].unique() if str(p).strip() and str(p).lower() != 'nan']))) if (COL_PARTIE in df.columns and not df.empty) else []
    options_partie = parties_existantes + ["➕ Autre / Nouvelle partie..."]

    # 3. Préparation dynamique de la liste "Essai / Contrôle"
    essais_base = ["Aucun", "TENEUR EN EAU", "CAMPACITÉ", "ESSAI À LA PLAQUE", "ESSAI À LA PLAQUE + CAMPACITÉ", "PRELEVEMENT APRES COMPACTAGE", "PRELEVEMENT AVANT COMPACTAGE", "IDENTIFICATION DES MATERIAUX", "PRELEVEMENT"]
    essais_bdd = sorted(list(set([str(e).strip() for e in df["ÉSSAI/ CONTRÔLE RÉALISÉE"].unique() if str(e).strip() and str(e).lower() != 'nan']))) if ("ÉSSAI/ CONTRÔLE RÉALISÉE" in df.columns and not df.empty) else []
    all_essais = sorted(list(set(essais_base + essais_bdd)))
    if "Aucun" in all_essais:
        all_essais.remove("Aucun")
        all_essais = ["Aucun"] + all_essais
    options_essai = all_essais + ["➕ Autre / Nouvel essai..."]

    col1, col2 = st.columns([1, 1])
    
    with col1:
        date_saisie = st.date_input("🗓️ Date des Travaux", value=datetime.today(), format="DD/MM/YYYY")
        
        # Choix ou Saisie Nature des travaux
        nature_choisie = st.selectbox("📌 Nature des travaux", options=options_nature)
        if nature_choisie == "➕ Autre / Nouvelle nature...":
            nature_selectionnee = st.text_input("✍️ Saisir la nouvelle Nature :").strip()
            info_liaison = {"procedure": "", "pieces": ""}
        else:
            nature_selectionnee = nature_choisie
            info_liaison = LIAISONS.get(nature_selectionnee, {"procedure": "", "pieces": ""})
        
        # Choix ou Saisie Partie d'ouvrage
        partie_choisie = st.selectbox("🧱 Partie d'ouvrage", options=options_partie)
        if partie_choisie == "➕ Autre / Nouvelle partie...":
            partie_ouvrage = st.text_input("✍️ Saisir la nouvelle Partie d'ouvrage :").strip()
        else:
            partie_ouvrage = partie_choisie

        situation = st.text_input("📍 Situation / PK", placeholder="Ex: PK 1+120 AU PK 1+220")
        
    with col2:
        activite = st.text_area("🚜 Activité réalisée", height=70)
        
        # Choix ou Saisie Essai / Contrôle
        essai_choisi = st.selectbox("🧪 Essai / Contrôle", options=options_essai)
        if essai_choisi == "➕ Autre / Nouvel essai...":
            essai = st.text_input("✍️ Saisir le nouvel Essai / Contrôle :").strip()
        else:
            essai = "" if essai_choisi == "Aucun" else essai_choisi

        procedure = st.text_input("📑 Procédure", value=info_liaison["procedure"])
        pieces_jointes = st.text_area("📎 Pièces jointes", value=info_liaison["pieces"], height=80)

    if st.button("💾 Enregistrer dans Google Sheets", type="primary", use_container_width=True):
        if not nature_selectionnee or not partie_ouvrage:
            st.error("⚠️ Veuillez renseigner au moins la Nature et la Partie d'ouvrage.")
        else:
            new_entry = {
                "DATE": date_saisie.strftime('%d/%m/%Y'),
                "TITRE DE LA NATURE DES TRAVAUX": nature_selectionnee,
                COL_PARTIE: partie_ouvrage,
                "SITUATION": situation,
                "ACTIVITÉ RÉALISÉE": activite,
                "ÉSSAI/ CONTRÔLE RÉALISÉE": essai,
                "RÉFÉRENCE DE PROCÉDURE": procedure,
                "PIÈCES JOINTES": pieces_jointes,
                "CRÉÉ PAR": st.session_state["username"]
            }
            df_updated = pd.concat([df, pd.DataFrame([new_entry])], ignore_index=True)
            success, msg = save_data_to_sheet(df_updated, sheet_name=chantier_actif)
            if success:
                st.success(msg)
                st.rerun()
            else:
                st.error(msg)

# -------------------------------------------------------------
# TAB 2 : REGISTRE
# -------------------------------------------------------------
with tab2:
    st.markdown("##### 🔍 **Registre des Travaux**")

    try:
        with st.expander("🌪️ **Filtres de recherche avancés**", expanded=False):
            col_f1, col_f2, col_f3 = st.columns([1, 1, 1])
            with col_f1:
                auteurs_existants = sorted(list(set([str(a) for a in df["CRÉÉ PAR"].unique() if str(a).strip() and str(a).lower() != 'nan']))) if "CRÉÉ PAR" in df.columns else []
                filtre_auteur = st.multiselect("👤 Auteur :", options=auteurs_existants)
            
            with col_f2:
                natures_existantes = sorted(list(set([str(n) for n in df["TITRE DE LA NATURE DES TRAVAUX"].unique() if str(n).strip() and str(n).lower() != 'nan']))) if "TITRE DE LA NATURE DES TRAVAUX" in df.columns else []
                filtre_nature = st.multiselect("📌 Nature :", options=natures_existantes)

            with col_f3:
                parties_filtre = sorted(list(set([str(p) for p in df[COL_PARTIE].unique() if str(p).strip() and str(p).lower() != 'nan']))) if COL_PARTIE in df.columns else []
                filtre_partie = st.multiselect("🧱 Partie :", options=parties_filtre)

            recherche_mot = st.text_input("🔍 Recherche globale par mot-clé :")

        st.markdown("**🔀 Trier les données du tableau :**")
        col_t1, col_t2 = st.columns([1, 1])
        with col_t1:
            colonne_tri = st.selectbox(
                "Colonne à trier :", 
                options=[c for c in df.columns if c != "Imprimer"], 
                index=0,
                key="select_col_tri"
            )
        with col_t2:
            sens_tri = st.selectbox(
                "Sens du tri :", 
                options=["A ➔ Z (Croissant)", "Z ➔ A (Décroissant)"],
                key="select_sens_tri"
            )

        df_filtered = df.copy()

        if filtre_auteur:
            df_filtered = df_filtered[df_filtered["CRÉÉ PAR"].astype(str).isin(filtre_auteur)]

        if filtre_nature:
            df_filtered = df_filtered[df_filtered["TITRE DE LA NATURE DES TRAVAUX"].astype(str).isin(filtre_nature)]

        if filtre_partie:
            df_filtered = df_filtered[df_filtered[COL_PARTIE].astype(str).isin(filtre_partie)]

        if recherche_mot.strip():
            m_clean = recherche_mot.strip().lower()
            df_filtered = df_filtered[
                df_filtered.apply(lambda row: row.astype(str).str.lower().str.contains(m_clean).any(), axis=1)
            ]

        if colonne_tri in df_filtered.columns:
            est_croissant = (sens_tri == "A ➔ Z (Croissant)")
            if colonne_tri == "DATE":
                df_filtered["DATE_TEMP"] = pd.to_datetime(df_filtered["DATE"], dayfirst=True, errors='coerce')
                df_filtered = df_filtered.sort_values(by="DATE_TEMP", ascending=est_croissant, na_position='last')
                df_filtered = df_filtered.drop(columns=["DATE_TEMP"])
            else:
                df_filtered = df_filtered.sort_values(by=colonne_tri, ascending=est_croissant, na_position='last')

        df_editor = df_filtered.copy()
        if "Imprimer" not in df_editor.columns:
            df_editor.insert(0, "Imprimer", False)

        edited_df = st.data_editor(
            df_editor, 
            column_config={
                "Imprimer": st.column_config.CheckboxColumn("Sélection", default=False),
                "CRÉÉ PAR": st.column_config.TextColumn("CRÉÉ PAR", disabled=True)
            },
            num_rows="dynamic", 
            height=380, 
            use_container_width=True
        )

        col_act1, col_act2 = st.columns([1, 1])
        with col_act1:
            if st.button("💾 Enregistrer modifications", type="secondary", use_container_width=True):
                try:
                    edited_clean = edited_df.drop(columns=["Imprimer"], errors="ignore")

                    df_to_save = df.copy()
                    df_to_save.update(edited_clean)

                    nouveaux_indexes = edited_clean.index.difference(df_to_save.index)
                    if not nouveaux_indexes.empty:
                        df_to_save = pd.concat([df_to_save, edited_clean.loc[nouveaux_indexes]], ignore_index=True)

                    success, msg = save_data_to_sheet(df_to_save, sheet_name=chantier_actif)
                    if success:
                        st.success(msg)
                        st.rerun()
                    else:
                        st.error(msg)
                except Exception as e_save:
                    st.error(f"❌ Erreur lors de la sauvegarde : {e_save}")

        with col_act2:
            lignes_selectionnees = edited_df[edited_df["Imprimer"] == True].copy()
            nb_selections = len(lignes_selectionnees)
            
            if st.button(f"📦 Générer ({nb_selections}) Fiche(s)", type="primary", use_container_width=True):
                if nb_selections == 0:
                    st.warning("⚠️ Cochez au moins une case dans le tableau.")
                else:
                    try:
                        zip_buffer = io.BytesIO()
                        fichiers_crees = 0
                        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                            for idx, row in lignes_selectionnees.iterrows():
                                nom_modele = get_col_val(row, "TITRE DE LA NATURE DES TRAVAUX", "NATURE")
                                chemin_modele = trouver_modele_word(nom_modele)
                                if chemin_modele:
                                    contexte = {
                                        'NATURE': get_col_val(row, "TITRE DE LA NATURE DES TRAVAUX", "NATURE"),
                                        'REF': get_col_val(row, "RÉFÉRENCE DE PROCÉDURE", "REF"),
                                        'PARTIE': get_col_val(row, "PARTIE D'OUVRAGE", "PARTIE D meOUVRAGE", "PARTIE"),
                                        'SITUATION': get_col_val(row, "SITUATION", "PK"),
                                        'PIECES': text_to_richtext(get_col_val(row, "PIÈCES JOINTES", "PIECES")),
                                        'DATE': get_col_val(row, "DATE"),
                                        'ACTIVITE': text_to_richtext(get_col_val(row, "ACTIVITÉ RÉALISÉE", "ACTIVITE")),
                                        'ESSAI': get_col_val(row, "ÉSSAI/ CONTRÔLE RÉALISÉE", "ESSAI"),
                                        'AUTEUR': get_col_val(row, "CRÉÉ PAR", "AUTEUR")
                                    }
                                    docx_b, pdf_b = generer_docx_et_pdf_bytes(chemin_modele, contexte)
                                    nom_base = construire_nom_pdf(row).replace(".pdf", "")
                                    zip_file.writestr(f"{nom_base}.docx", docx_b)
                                    zip_file.writestr(f"{nom_base}.pdf", pdf_b)
                                    fichiers_crees += 1

                        if fichiers_crees > 0:
                            zip_buffer.seek(0)
                            st.download_button(
                                label="⬇️ Télécharger Pack ZIP",
                                data=zip_buffer,
                                file_name="Fiches_Chantier.zip",
                                mime="application/zip",
                                use_container_width=True
                            )
                        else:
                            st.error("❌ Aucun modèle Word trouvé correspondant.")
                    except Exception as e_gen:
                        st.error(f"❌ Erreur de génération : {e_gen}")

    except Exception as e_tab:
        st.error(f"❌ Erreur sur le tableau : {e_tab}")

# -------------------------------------------------------------
# TAB 3 : DEMANDES D'INTERVENTION (DI)
# -------------------------------------------------------------
with tab3:
    st.subheader("📅 Demandes d'Intervention (DI)")
    date_range = st.date_input("📅 Sélectionner date / période :", value=(), format="DD/MM/YYYY")
    if 'df' in locals() and df is not None and not df.empty:
        df_temp = df.copy()
        df_temp['DATE_DT'] = pd.to_datetime(df_temp['DATE'], dayfirst=True, errors='coerce').dt.date
        df_filtered_di = pd.DataFrame()
        
        if len(date_range) == 2:
            df_filtered_di = df_temp[(df_temp['DATE_DT'] >= date_range[0]) & (df_temp['DATE_DT'] <= date_range[1])]
        elif len(date_range) == 1:
            df_filtered_di = df_temp[df_temp['DATE_DT'] == date_range[0]]

        if not df_filtered_di.empty:
            st.dataframe(df_filtered_di.drop(columns=['DATE_DT'], errors='ignore'), use_container_width=True)
            if st.button("📦 Générer Pack DI", type="primary", use_container_width=True):
                zip_data, count_dates = generer_pack_di_zip(df_filtered_di)
                st.download_button(
                    label="⬇️ Télécharger Le Pack ZIP", 
                    data=zip_data, 
                    file_name="Pack_DI.zip", 
                    mime="application/zip", 
                    use_container_width=True
                )
        elif len(date_range) > 0:
            st.info("ℹ️ Aucune donnée trouvée pour la période sélectionnée.")

# -------------------------------------------------------------
# TAB 4 : ESPACE ADMINISTRATEUR
# -------------------------------------------------------------
if st.session_state["role"] == "Admin":
    tab_admin = tabs[3]
    with tab_admin:
        st.markdown("##### 👥 **Gestion des Utilisateurs**")
        df_users = load_users()

        col_u1, col_u2 = st.columns([1, 1])

        with col_u1:
            st.markdown("**➕ Ajouter un utilisateur**")
            with st.form("form_add_user"):
                new_username = st.text_input("Nom d'utilisateur").strip()
                new_password = st.text_input("Mot de passe", type="password").strip()
                new_role = st.selectbox("Rôle", options=["Utilisateur", "Admin"])
                new_status = st.selectbox("Compte Actif", options=["OUI", "NON"])
                btn_add_user = st.form_submit_button("Ajouter Utilisateur", type="primary", use_container_width=True)

                if btn_add_user:
                    if not new_username or not new_password:
                        st.error("⚠️ Champs obligatoires manquants.")
                    elif new_username in df_users["username"].astype(str).values:
                        st.error("⚠️ Cet utilisateur existe déjà.")
                    else:
                        new_row = {
                            "username": new_username,
                            "password": hash_password(new_password),
                            "role": new_role,
                            "actif": new_status
                        }
                        df_users_updated = pd.concat([df_users, pd.DataFrame([new_row])], ignore_index=True)
                        ok, msg = save_users(df_users_updated)
                        if ok:
                            st.success(f"Utilisateur {new_username} créé !")
                            st.rerun()
                        else:
                            st.error(msg)

        with col_u2:
            st.markdown("**📜 Utilisateurs inscrits**")
            users_edited = st.data_editor(
                df_users,
                column_config={
                    "password": st.column_config.TextColumn("Mot de passe (Hash)", disabled=True),
                    "role": st.column_config.SelectboxColumn("Rôle", options=["Admin", "Utilisateur"], required=True),
                    "actif": st.column_config.SelectboxColumn("Actif", options=["OUI", "NON"], required=True)
                },
                num_rows="dynamic",
                use_container_width=True
            )

            if st.button("💾 Sauvegarder les comptes", type="secondary", use_container_width=True):
                ok, msg = save_users(users_edited)
                if ok:
                    st.success("Modifications enregistrées !")
                    st.rerun()
                else:
                    st.error(msg)

            with st.expander("🔑 **Réinitialiser un mot de passe**"):
                user_to_reset = st.selectbox("Sélectionner un compte :", options=df_users["username"].tolist())
                reset_pass = st.text_input("Nouveau mot de passe :", type="password", key="reset_pass_val")
                if st.button("🔒 Valider le mot de passe", use_container_width=True):
                    if reset_pass.strip():
                        df_users.loc[df_users["username"] == user_to_reset, "password"] = hash_password(reset_pass.strip())
                        ok, msg = save_users(df_users)
                        if ok:
                            st.success(f"Mot de passe de {user_to_reset} mis à jour !")
                            st.rerun()
                        else:
                            st.error(msg)
                    else:
                        st.warning("⚠️ Entrez un mot de passe valide.")
