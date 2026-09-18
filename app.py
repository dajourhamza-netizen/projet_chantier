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

USER_COLUMNS = ["username", "password", "role", "actif", "chantiers"]

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
    "ASSISE DE REMBLAIS CDF": {
        "procedure": "TER-PEX-04-00",
        "pieces": "* Fiche de réception de l'assise des remblais\n* Fiche de réception topographique\n* PVs laboratoire"
    },
    "ASSISE DE REMBLAIS CONTIGUS": {
        "procedure": "OVA-PEX-16-00",
        "pieces": "* Fiche de suivi des remblais contigus\n* Fiche de contrôle des remblais contigus\n* PVs laboratoire\n* Fiche de réception topographique"
    },
    "ASSISE DE REMBLAI DE FOUILLE": {
        "procedure": "OVA-PEX-04-00",
        "pieces": "* Fiche de suivi et de contrôle des fouilles et remblaiement de fouilles\n* PVs laboratoire"
    },
    "ASSISE DE REMBLAIS RENFORCE": {
        "procedure": "TER-PEX-13-00",
        "pieces": "* PV Manifold\n* PVs laboratoire\n* Fiche de réception topographique\n* Fiche de réception assise remblai renforcé"
    },
    "ASSISE DRAINANTE": {
        "procedure": "TER-PEX-13-00",
        "pieces": "* Fiche de réception topographique\n* PVs laboratoire\n* Fiche de contrôle de l'assise drainante"
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
    "REMBLAIS CDF": {
        "procedure": "TER-PEX-04-00",
        "pieces": "* Fiche de suivi et de contrôle des remblais\n* PVs laboratoire"
    },
    "REMBLAIS CONTIGUS": {
        "procedure": "OVA-PEX-16-00",
        "pieces": "* Fiche de suivi des remblais contigus\n* Fiche de contrôle des remblais contigus\n* PVs laboratoire\n* Fiche de réception topographique"
    },
    "REMBLAIS DE FOUILLE": {
        "procedure": "OVA-PEX-04-00",
        "pieces": "* Fiche de suivi et de contrôle des fouilles et remblaiement de fouilles\n* PVs laboratoire"
    },
    "REMBLAIS DE FOUILLS CDF": {
        "procedure": "OVA-PEX-04-00",
        "pieces": "* Fiche de suivi et de contrôle des fouilles et remblaiement de fouilles\n* PVs laboratoire"
    },
    "REMBLAIS RENFORCE": {
        "procedure": "TER-PEX-13-00",
        "pieces": "* Fiche de suivi des remblais renforcé\n* Fiche de contrôle des armatures Geostrap\n* Fiche de réception de pose des ecailles\n* PVs laboratoire"
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

# إضافة التخزين المؤقت للاحتفاظ بالبيانات لمدة 60 ثانية
@st.cache_data(ttl=60) 
def load_users():
    try:
        sh = get_spreadsheet()
        try:
            ws = sh.worksheet("Utilisateurs")
        except gspread.WorksheetNotFound:
            return admin_initial

        records = ws.get_all_records()
        df_u = pd.DataFrame(records)
        if "chantiers" not in df_u.columns:
            df_u["chantiers"] = "TOUS"
        return df_u
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
        st.cache_data.clear() # Kymse7 l'cache bach l'application t9ra les utilisateurs jdad
        return True, "✅ Utilisateurs mis à jour !"
    except Exception as e:
        return False, f"❌ Erreur : {e}"
def log_user_login(username, role):
    try:
        sh = get_spreadsheet()
        try:
            ws = sh.worksheet("Connexions")
        except gspread.WorksheetNotFound:
            ws = sh.add_worksheet(title="Connexions", rows=200, cols=3)
            ws.append_row(["DATE ET HEURE", "UTILISATEUR", "RÔLE"])
        
        horodatage = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        ws.append_row([horodatage, username, role])
    except Exception:
        pass

def load_login_history():
    try:
        sh = get_spreadsheet()
        ws = sh.worksheet("Connexions")
        records = ws.get_all_records()
        return pd.DataFrame(records)
    except Exception:
        return pd.DataFrame(columns=["DATE ET HEURE", "UTILISATEUR", "RÔLE"])

@st.cache_data(ttl=60)
def get_sheet_names_gsheets():
    try:
        sh = get_spreadsheet()
        onglets_exclus = ["Utilisateurs", "Connexions"]
        return [ws.title for ws in sh.worksheets() if ws.title not in onglets_exclus]
    except Exception as e:
        st.error(f"Erreur de connexion à Google Sheets : {e}")
        return ["Chantier Principal"]

@st.cache_data(ttl=60)
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
        st.cache_data.clear() # مسح التخزين المؤقت ليتم جلب التحديثات الجديدة
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
    text = str(text).strip()
    if text.lower().endswith('.docx'): text = text[:-5]
    text = unicodedata.normalize('NFD', text).encode('ascii', 'ignore').decode("utf-8")
    return re.sub(r'[^a-zA-Z0-9]', '', text).lower()

def trouver_modele_word(nom_nature, nom_chantier=None):
    if not nom_nature:
        return None
        
    target_clean = clean_filename(nom_nature)

    # 1. Recherche dans le dossier spécifique du chantier (ex: NGE ou JET CONTRACTORS)
    if nom_chantier and os.path.exists(DOSSIER_CHANTIER):
        dossier_cible = None
        for item in os.listdir(DOSSIER_CHANTIER):
            chemin_item = os.path.join(DOSSIER_CHANTIER, item)
            if os.path.isdir(chemin_item) and item.strip().lower() == str(nom_chantier).strip().lower():
                dossier_cible = chemin_item
                break
        
        if dossier_cible and os.path.exists(dossier_cible):
            for root, _, files in os.walk(dossier_cible):
                for file in files:
                    if file.lower().endswith('.docx') and not file.startswith('~$'):
                        if clean_filename(file) == target_clean:
                            return os.path.join(root, file)

    # 2. Recherche globale dans tout le répertoire du projet
    if os.path.exists(DOSSIER_CHANTIER):
        for root, _, files in os.walk(DOSSIER_CHANTIER):
            for file in files:
                if file.lower().endswith('.docx') and not file.startswith('~$'):
                    if clean_filename(file) == target_clean:
                        return os.path.join(root, file)

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

def generer_di_une_date(df_jour, nom_chantier=None):
    modele_di = None
    if os.path.exists(DOSSIER_CHANTIER):
        # 1. Chercher dans le dossier du chantier
        if nom_chantier:
            dossier_cible = None
            for item in os.listdir(DOSSIER_CHANTIER):
                chemin_item = os.path.join(DOSSIER_CHANTIER, item)
                if os.path.isdir(chemin_item) and item.strip().lower() == str(nom_chantier).strip().lower():
                    dossier_cible = chemin_item
                    break
            if dossier_cible and os.path.exists(dossier_cible):
                for root, _, files in os.walk(dossier_cible):
                    for file in files:
                        if file.lower().endswith('.docx') and not file.startswith('~$'):
                            if 'di' in file.lower() or 'demande' in file.lower():
                                modele_di = os.path.join(root, file)
                                break
                    if modele_di: break

        # 2. Recherche globale si non trouvé
        if not modele_di:
            for root, _, files in os.walk(DOSSIER_CHANTIER):
                for file in files:
                    if file.lower().endswith('.docx') and not file.startswith('~$'):
                        if 'di' in file.lower() or 'demande' in file.lower():
                            modele_di = os.path.join(root, file)
                            break
                if modele_di: break

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

def generer_pack_di_zip(df_filtered, nom_chantier=None):
    zip_buffer = io.BytesIO()
    dates_uniques = df_filtered["DATE"].unique()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        for date_val in dates_uniques:
            if not date_val or str(date_val).strip() == "" or str(date_val).lower() == "nan": continue
            df_jour = df_filtered[df_filtered["DATE"] == date_val]
            docx_b, pdf_b = generer_di_une_date(df_jour, nom_chantier)
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
    st.session_state["chantiers"] = "TOUS"

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
                        st.session_state["chantiers"] = str(info_user.get("chantiers", "TOUS"))
                        
                        log_user_login(info_user["username"], info_user["role"])
                        
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
# 6. BARRE LATÉRALE & FILTRAGE DES CHANTIERS
# ==========================================
st.sidebar.markdown(f"👤 **{st.session_state['username']}** ({st.session_state['role']})")

if st.sidebar.button("🚪 Déconnexion", use_container_width=True):
    st.session_state["authenticated"] = False
    st.session_state["username"] = None
    st.session_state["role"] = None
    st.session_state["chantiers"] = "TOUS"
    st.rerun()

st.sidebar.markdown("---")
st.sidebar.markdown("### 🌐 **Google Sheets**")

chantiers_existants = get_sheet_names_gsheets()

user_chantiers_raw = st.session_state.get("chantiers", "TOUS")
if st.session_state["role"] != "Admin" and user_chantiers_raw != "TOUS":
    list_chantiers_user = [c.strip() for c in user_chantiers_raw.split(",") if c.strip()]
    chantiers_autorises = [c for c in chantiers_existants if c in list_chantiers_user]
    if not chantiers_autorises:
        chantiers_autorises = chantiers_existants
else:
    chantiers_autorises = chantiers_existants

chantier_actif = st.sidebar.selectbox("📌 **Projet Actif :**", options=chantiers_autorises)

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
# 7. INTERFACE PRINCIPALE SELON RÔLE
# ==========================================
st.markdown(f"""
<div class="gc-header">
    <h1>🛣️ Suivi Génie Civil</h1>
    <p>Projet : <b>{chantier_actif}</b></p>
</div>
""", unsafe_allow_html=True)

role_actuel = st.session_state["role"]

tab_saisie, tab_registre, tab_di, tab_admin = None, None, None, None

if role_actuel == "Lecteur":
    tabs = st.tabs(["📊 **Tableau de suivi**"])
    tab_registre = tabs[0]
elif role_actuel == "Admin":
    tabs = st.tabs(["📝 **Saisie**", "📊 **Registre**", "📅 **DI**", "👥 **Accès & Logs**"])
    tab_saisie, tab_registre, tab_di, tab_admin = tabs[0], tabs[1], tabs[2], tabs[3]
else:  # Utilisateur
    tabs = st.tabs(["📝 **Saisie**", "📊 **Registre**", "📅 **DI**"])
    tab_saisie, tab_registre, tab_di = tabs[0], tabs[1], tabs[2]

# -------------------------------------------------------------
# TAB 1 : SAISIE DES DONNÉES (ADMIN ET UTILISATEUR UNIQUEMENT)
# -------------------------------------------------------------
if tab_saisie:
    with tab_saisie:
        st.markdown("##### 👷 **Ajouter une fiche**")
        
        natures_bdd = sorted(list(set([str(n).strip() for n in df["TITRE DE LA NATURE DES TRAVAUX"].unique() if str(n).strip() and str(n).lower() != 'nan']))) if ("TITRE DE LA NATURE DES TRAVAUX" in df.columns and not df.empty) else []
        all_natures = sorted(list(set(list(LIAISONS.keys()) + natures_bdd)))
        options_nature = all_natures + ["➕ Autre / Nouvelle nature..."]

        parties_existantes = sorted(list(set([str(p).strip() for p in df[COL_PARTIE].unique() if str(p).strip() and str(p).lower() != 'nan']))) if (COL_PARTIE in df.columns and not df.empty) else []
        options_partie = parties_existantes + ["➕ Autre / Nouvelle partie..."]

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
            
            nature_choisie = st.selectbox("📌 Nature des travaux", options=options_nature)
            if nature_choisie == "➕ Autre / Nouvelle nature...":
                nature_selectionnee = st.text_input("✍️ Saisir la nouvelle Nature :").strip()
                info_liaison = {"procedure": "", "pieces": ""}
            else:
                nature_selectionnee = nature_choisie
                info_liaison = LIAISONS.get(nature_selectionnee, {"procedure": "", "pieces": ""})
            
            partie_choisie = st.selectbox("🧱 Partie d'ouvrage", options=options_partie)
            if partie_choisie == "➕ Autre / Nouvelle partie...":
                partie_ouvrage = st.text_input("✍️ Saisir la nouvelle Partie d'ouvrage :").strip()
            else:
                partie_ouvrage = partie_choisie

            situation = st.text_input("📍 Situation / PK", placeholder="Ex: PK 1+120 AU PK 1+220")
            
        with col2:
            activite = st.text_area("🚜 Activité réalisée", height=70)
            
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
# TAB 2 : REGISTRE / TABLEAU DE SUIVI
# -------------------------------------------------------------
if tab_registre:
    with tab_registre:
        st.markdown("##### 🔍 **Tableau de Suivi des Travaux**")

        try:
            with st.expander("🌪️ **Filtres de recherche avancés**", expanded=False):
                col_f1, col_f2, col_f3, col_f4 = st.columns([1, 1, 1, 1])
                with col_f1:
                    auteurs_existants = sorted(list(set([str(a) for a in df["CRÉÉ PAR"].unique() if str(a).strip() and str(a).lower() != 'nan']))) if "CRÉÉ PAR" in df.columns else []
                    filtre_auteur = st.multiselect("👤 Auteur :", options=auteurs_existants)
                
                with col_f2:
                    natures_existantes = sorted(list(set([str(n) for n in df["TITRE DE LA NATURE DES TRAVAUX"].unique() if str(n).strip() and str(n).lower() != 'nan']))) if "TITRE DE LA NATURE DES TRAVAUX" in df.columns else []
                    filtre_nature = st.multiselect("📌 Nature :", options=natures_existantes)

                with col_f3:
                    parties_filtre = sorted(list(set([str(p) for p in df[COL_PARTIE].unique() if str(p).strip() and str(p).lower() != 'nan']))) if COL_PARTIE in df.columns else []
                    filtre_partie = st.multiselect("🧱 Partie :", options=parties_filtre)

                with col_f4:
                    # 1. On crée une copie temporaire pour filtrer selon la Partie sélectionnée
                    df_pour_situation = df.copy()
                    if filtre_partie:
                        df_pour_situation = df_pour_situation[df_pour_situation[COL_PARTIE].astype(str).isin(filtre_partie)]
                    
                    # 2. On extrait les situations uniquement de cette liste filtrée
                    situations_existantes = sorted(list(set([str(s) for s in df_pour_situation["SITUATION"].unique() if str(s).strip() and str(s).lower() != 'nan']))) if "SITUATION" in df_pour_situation.columns else []
                    
                    # 3. On affiche le menu déroulant
                    filtre_situation = st.multiselect("📍 Situation :", options=situations_existantes)

                    # Passage à 5 colonnes
                col_f1, col_f2, col_f3, col_f4, col_f5 = st.columns(5)
                
                # ... (gardez votre code actuel pour col_f1, col_f2, col_f3 et col_f4) ...

                with col_f5:
                    essais_existants = sorted(list(set([str(e) for e in df["ÉSSAI/ CONTRÔLE RÉALISÉE"].unique() if str(e).strip() and str(e).lower() != 'nan']))) if "ÉSSAI/ CONTRÔLE RÉALISÉE" in df.columns else []
                    filtre_essai = st.multiselect("🔬 Essai / Contrôle :", options=essais_existants)


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

            if role_actuel == "Lecteur":
                st.dataframe(df_filtered, use_container_width=True, height=450)
            
            else:
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
                                        chemin_modele = trouver_modele_word(nom_modele, chantier_actif)
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
if tab_di:
    with tab_di:
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
                    zip_data, count_dates = generer_pack_di_zip(df_filtered_di, chantier_actif)
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
# TAB 4 : ESPACE ADMINISTRATEUR (COMPTES & HISTORIQUE D'ACCÈS)
# -------------------------------------------------------------
if tab_admin:
    with tab_admin:
        st.markdown("##### 👥 **Gestion des Utilisateurs & Droits par Chantier**")
        df_users = load_users()

        col_u1, col_u2 = st.columns([1, 1])

        with col_u1:
            st.markdown("**➕ Ajouter un utilisateur**")
            with st.form("form_add_user"):
                new_username = st.text_input("Nom d'utilisateur").strip()
                new_password = st.text_input("Mot de passe", type="password").strip()
                new_role = st.selectbox("Rôle", options=["Utilisateur", "Lecteur", "Admin"])
                new_status = st.selectbox("Compte Actif", options=["OUI", "NON"])
                
                chantiers_dispos = get_sheet_names_gsheets()
                chantiers_select = st.multiselect("🏗️ Chantiers autorisés :", options=chantiers_dispos, default=chantiers_dispos)
                
                btn_add_user = st.form_submit_button("Ajouter Utilisateur", type="primary", use_container_width=True)

                if btn_add_user:
                    if not new_username or not new_password:
                        st.error("⚠️ Champs obligatoires manquants.")
                    elif new_username in df_users["username"].astype(str).values:
                        st.error("⚠️ Cet utilisateur existe déjà.")
                    else:
                        str_chantiers = "TOUS" if new_role == "Admin" else ",".join(chantiers_select) if chantiers_select else "TOUS"
                        new_row = {
                            "username": new_username,
                            "password": hash_password(new_password),
                            "role": new_role,
                            "actif": new_status,
                            "chantiers": str_chantiers
                        }
                        df_users_updated = pd.concat([df_users, pd.DataFrame([new_row])], ignore_index=True)
                        ok, msg = save_users(df_users_updated)
                        if ok:
                            st.success(f"Utilisateur {new_username} ({new_role}) créé !")
                            st.rerun()
                        else:
                            st.error(msg)

        with col_u2:
            st.markdown("**📜 Utilisateurs inscrits & Accès**")
            users_edited = st.data_editor(
                df_users,
                column_config={
                    "password": st.column_config.TextColumn("Mot de passe (Hash)", disabled=True),
                    "role": st.column_config.SelectboxColumn("Rôle", options=["Admin", "Utilisateur", "Lecteur"], required=True),
                    "actif": st.column_config.SelectboxColumn("Actif", options=["OUI", "NON"], required=True),
                    "chantiers": st.column_config.TextColumn("Chantiers (séparés par virgule)")
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

        st.markdown("---")
        st.markdown("##### 🕒 **Historique des Connexions (Qui est entré et quand)**")
        df_logs = load_login_history()
        if not df_logs.empty:
            st.dataframe(df_logs.iloc[::-1], use_container_width=True, height=280)
        else:
            st.info("ℹ️ Aucune connexion enregistrée pour le moment.")