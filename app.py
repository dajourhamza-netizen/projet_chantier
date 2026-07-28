import io
import os
import tempfile
import zipfile
from datetime import datetime

import pandas as pd
import streamlit as st

from config import (
    CHEMIN_EXCEL_DEFAUT,
    COL_PARTIE,
    COLUMNS_TEMPLATE,
    ESSAI_OPTIONS,
    LIAISONS,
)
from data import DataStore, gsheets_available
from documents import (
    construire_nom_pdf,
    convert_docx_to_pdf,
    generer_di_style_vba,
    generer_docx_et_pdf_bytes,
    get_col_val,
    text_to_richtext,
    trouver_modele_word,
)
from ui import (
    apply_filters,
    inject_styles,
    render_dashboard_charts,
    render_header,
    render_kpi_cards,
    validate_saisie,
)

# ==========================================
# CONFIGURATION
# ==========================================
st.set_page_config(
    page_title="Suivi Chantier - Génie Civil & Routes",
    page_icon="🏗️",
    layout="wide",
    initial_sidebar_state="expanded",
)

inject_styles()

# ==========================================
# SIDEBAR — SOURCE DE DONNÉES & PROJETS
# ==========================================
st.sidebar.markdown("### 🏗️ **Gestion des Chantiers**")

source_options = ["Excel local"]
if gsheets_available():
    source_options.append("Google Sheets")

data_mode = st.sidebar.radio(
    "📂 **Source des données**",
    options=source_options,
    horizontal=True,
)

if data_mode == "Google Sheets":
    default_url = ""
    try:
        if "gsheets" in st.secrets:
            default_url = st.secrets["gsheets"].get("spreadsheet_url", "")
    except Exception:
        pass

    gs_url = st.sidebar.text_input(
        "🔗 URL Google Sheets",
        value=st.session_state.get("gsheets_url", default_url),
        placeholder="https://docs.google.com/spreadsheets/d/...",
    )
    st.session_state["gsheets_url"] = gs_url

store = DataStore(mode=data_mode)
chantiers_existants = store.list_projects()

chantier_actif = st.sidebar.selectbox(
    "📌 **Projet Actif :**",
    options=chantiers_existants,
)

st.sidebar.markdown("---")

with st.sidebar.expander("➕ **Créer un Nouveau Projet**", expanded=False):
    nouveau_projet_nom = st.text_input(
        "Nom du projet / chantier :",
        key="new_proj_input",
    )
    if st.button(
        "✨ Créer le Projet",
        type="primary",
        key="btn_create_proj",
        use_container_width=True,
    ):
        nom_clean = nouveau_projet_nom.strip()
        if not nom_clean:
            st.sidebar.warning("⚠️ Veuillez entrer un nom valide.")
        elif nom_clean in chantiers_existants:
            st.sidebar.error("⚠️ Ce projet existe déjà !")
        else:
            success, msg = store.create_project(nom_clean)
            if success:
                st.sidebar.success(f"✅ Projet '{nom_clean}' créé !")
                st.rerun()
            else:
                st.sidebar.error(msg)

if st.sidebar.button("🔄 Rafraîchir les données", use_container_width=True):
    st.cache_data.clear()
    st.rerun()

# Chargement des données
df = store.load(chantier_actif)
if df is not None and not df.empty:
    st.sidebar.success(
        f"✅ '{chantier_actif}' — {len(df)} fiche(s) · {store.source_label}"
    )
else:
    st.sidebar.info(f"📋 Projet vide · {store.source_label}")

# ==========================================
# INTERFACE PRINCIPALE
# ==========================================
render_header(chantier_actif, store.source_label)

if df is None:
    st.error("Impossible de charger les données.")
    st.stop()

tab_dash, tab1, tab2, tab3 = st.tabs([
    "📈 **Tableau de bord**",
    "📝 **Nouvelle Saisie**",
    "📊 **Registre & Export**",
    "📅 **Demandes d'Intervention (DI)**",
])

# -------------------------------------------------------------
# TAB DASHBOARD
# -------------------------------------------------------------
with tab_dash:
    render_kpi_cards(df)
    render_dashboard_charts(df)

    if not df.empty and COL_PARTIE in df.columns:
        st.markdown("**Répartition par partie d'ouvrage**")
        partie_counts = (
            df[COL_PARTIE].astype(str).value_counts().head(8).reset_index()
        )
        partie_counts.columns = ["Partie", "Fiches"]
        st.bar_chart(partie_counts, x="Partie", y="Fiches", height=240)

# -------------------------------------------------------------
# TAB 1 : SAISIE AVEC VALIDATION
# -------------------------------------------------------------
with tab1:
    st.markdown("##### 👷 **Ajouter une nouvelle fiche de contrôle**")

    with st.form("form_saisie", clear_on_submit=False):
        col1, col2 = st.columns(2)

        with col1:
            date_saisie = st.date_input(
                "🗓️ Date des travaux",
                value=datetime.today(),
                format="DD/MM/YYYY",
            )
            nature_selectionnee = st.selectbox(
                "📌 Nature des travaux",
                options=list(LIAISONS.keys()),
            )
            info_liaison = LIAISONS.get(
                nature_selectionnee, {"procedure": "", "pieces": ""}
            )

            parties_existantes = []
            if COL_PARTIE in df.columns:
                parties_existantes = sorted(
                    {
                        str(p).strip()
                        for p in df[COL_PARTIE].unique()
                        if str(p).strip() and str(p).lower() != "nan"
                    }
                )

            options_partie = parties_existantes + ["➕ Autre / Nouvelle partie..."]
            partie_choisie = st.selectbox(
                "🧱 Partie d'ouvrage",
                options=options_partie,
            )

            if partie_choisie == "➕ Autre / Nouvelle partie...":
                partie_ouvrage = st.text_input(
                    "✍️ Nouvelle partie d'ouvrage",
                    placeholder="Ex: CULEE C0...",
                )
            else:
                partie_ouvrage = partie_choisie

            situation = st.text_input(
                "📍 Situation / PK",
                placeholder="Ex: PK 1+120 AU PK 1+220",
            )

        with col2:
            activite = st.text_area(
                "🚜 Activité réalisée",
                height=80,
                placeholder="Ex: 1ère couche, 2ème couche",
            )
            essai = st.selectbox(
                "🧪 Essai / Contrôle réalisé",
                options=ESSAI_OPTIONS,
            )
            procedure = st.text_input(
                "📑 Référence procédure",
                value=info_liaison["procedure"],
            )
            pieces_jointes = st.text_area(
                "📎 Pièces jointes",
                value=info_liaison["pieces"],
                height=100,
            )

        submitted = st.form_submit_button(
            "💾 Enregistrer la fiche",
            type="primary",
            use_container_width=True,
        )

    if submitted:
        errors = validate_saisie(partie_ouvrage, situation, activite)
        if errors:
            for err in errors:
                st.error(err)
        else:
            new_entry = {
                "DATE": date_saisie.strftime("%d/%m/%Y"),
                "TITRE DE LA NATURE DES TRAVAUX": nature_selectionnee,
                COL_PARTIE: partie_ouvrage.strip(),
                "SITUATION": situation.strip(),
                "ACTIVITÉ RÉALISÉE": activite.strip(),
                "ÉSSAI/ CONTRÔLE RÉALISÉE": "" if essai == "Aucun" else essai,
                "RÉFÉRENCE DE PROCÉDURE": procedure,
                "PIÈCES JOINTES": pieces_jointes,
            }
            df_updated = pd.concat(
                [df, pd.DataFrame([new_entry])],
                ignore_index=True,
            )
            success, msg = store.save(df_updated, chantier_actif)
            if success:
                st.success(msg)
                st.rerun()
            else:
                st.error(msg)

# -------------------------------------------------------------
# TAB 2 : REGISTRE (FILTRES AVANCÉS + FRAGMENT)
# -------------------------------------------------------------
with tab2:
    st.markdown("##### 🔍 **Registre, filtrage & exportation**")

    @st.fragment
    def registre_filtres():
        with st.expander("🎯 **Filtres avancés**", expanded=True):
            col_f1, col_f2, col_f3 = st.columns(3)

            natures_all = sorted(
                {
                    str(n).strip()
                    for n in df["TITRE DE LA NATURE DES TRAVAUX"].unique()
                    if str(n).strip() and str(n).lower() != "nan"
                }
            ) if "TITRE DE LA NATURE DES TRAVAUX" in df.columns else []

            parties_all = sorted(
                {
                    str(p).strip()
                    for p in df[COL_PARTIE].unique()
                    if str(p).strip() and str(p).lower() != "nan"
                }
            ) if COL_PARTIE in df.columns else []

            with col_f1:
                filtre_natures = st.multiselect(
                    "📌 Nature(s)",
                    options=natures_all,
                )
            with col_f2:
                filtre_parties = st.multiselect(
                    "🧱 Partie(s) d'ouvrage",
                    options=parties_all,
                )
            with col_f3:
                recherche_texte = st.text_input(
                    "🔍 Recherche globale",
                    placeholder="PK, activité, nature...",
                )

            use_period = st.checkbox("Filtrer par période", value=False)
            if use_period:
                col_d1, col_d2 = st.columns(2)
                with col_d1:
                    date_debut = st.date_input(
                        "📅 Date début",
                        format="DD/MM/YYYY",
                    )
                with col_d2:
                    date_fin = st.date_input(
                        "📅 Date fin",
                        format="DD/MM/YYYY",
                    )
            else:
                date_debut = None
                date_fin = None

        return (
            filtre_natures,
            filtre_parties,
            date_debut,
            date_fin,
            recherche_texte,
        )

    filtre_natures, filtre_parties, date_debut, date_fin, recherche_texte = (
        registre_filtres()
    )

    df_filtre = apply_filters(
        df,
        filtre_natures,
        filtre_parties,
        date_debut,
        date_fin,
        recherche_texte,
    )

    st.caption(
        f"📊 Résultats : **{len(df_filtre)}** / {len(df)} fiche(s)"
    )

    df_editor = df_filtre.copy()
    if "Imprimer" not in df_editor.columns:
        df_editor.insert(0, "Imprimer", False)

    edited_df = st.data_editor(
        df_editor,
        num_rows="dynamic",
        height=400,
        use_container_width=True,
        key="registre_editor",
    )

    st.markdown("---")
    col_act1, col_act2, col_act3 = st.columns(3)

    with col_act1:
        if st.button(
            "💾 Enregistrer les modifications",
            type="secondary",
            use_container_width=True,
        ):
            df_sauvegarde = df.copy()
            edited_clean = edited_df.drop(columns=["Imprimer"], errors="ignore")
            df_sauvegarde.loc[edited_clean.index] = edited_clean
            success, msg = store.save(df_sauvegarde, chantier_actif)
            if success:
                st.success(msg)
                st.rerun()
            else:
                st.error(msg)

    with col_act2:
        export_buffer = io.BytesIO()
        edited_clean = edited_df.drop(columns=["Imprimer"], errors="ignore")
        with pd.ExcelWriter(export_buffer, engine="openpyxl") as writer:
            edited_clean.to_excel(writer, index=False, sheet_name="Export")
        export_buffer.seek(0)
        st.download_button(
            label="📥 Exporter la sélection (Excel)",
            data=export_buffer,
            file_name=f"Export_{chantier_actif}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    with col_act3:
        lignes_selectionnees = edited_df[edited_df["Imprimer"] == True]
        nb_selections = len(lignes_selectionnees)

        if st.button(
            f"📦 Générer fiches ({nb_selections})",
            type="primary",
            use_container_width=True,
        ):
            if nb_selections == 0:
                st.warning("⚠️ Cochez au moins une case 'Imprimer'.")
            else:
                zip_buffer = io.BytesIO()
                fichiers_crees = 0
                with zipfile.ZipFile(
                    zip_buffer, "w", zipfile.ZIP_DEFLATED
                ) as zip_file:
                    for _, row in lignes_selectionnees.iterrows():
                        nom_modele = get_col_val(
                            row, "TITRE DE LA NATURE DES TRAVAUX", "NATURE"
                        )
                        chemin_modele = trouver_modele_word(nom_modele)
                        if chemin_modele:
                            contexte = {
                                "NATURE": get_col_val(
                                    row,
                                    "TITRE DE LA NATURE DES TRAVAUX",
                                    "NATURE",
                                ),
                                "REF": get_col_val(
                                    row, "RÉFÉRENCE DE PROCÉDURE", "REF"
                                ),
                                "PARTIE": get_col_val(
                                    row,
                                    "PARTIE D'OUVRAGE",
                                    "PARTIE D meOUVRAGE",
                                    "PARTIE",
                                ),
                                "SITUATION": get_col_val(row, "SITUATION", "PK"),
                                "PIECES": text_to_richtext(
                                    get_col_val(row, "PIÈCES JOINTES", "PIECES")
                                ),
                                "DATE": get_col_val(row, "DATE"),
                                "ACTIVITE": text_to_richtext(
                                    get_col_val(
                                        row, "ACTIVITÉ RÉALISÉE", "ACTIVITE"
                                    )
                                ),
                                "ESSAI": get_col_val(
                                    row, "ÉSSAI/ CONTRÔLE RÉALISÉE", "ESSAI"
                                ),
                            }
                            docx_b, pdf_b = generer_docx_et_pdf_bytes(
                                chemin_modele, contexte
                            )
                            nom_base = construire_nom_pdf(row).replace(".pdf", "")
                            zip_file.writestr(f"{nom_base}.docx", docx_b)
                            zip_file.writestr(f"{nom_base}.pdf", pdf_b)
                            fichiers_crees += 1

                if fichiers_crees > 0:
                    zip_buffer.seek(0)
                    st.download_button(
                        label="📦 Télécharger Pack ZIP",
                        data=zip_buffer,
                        file_name="Fiches_Chantier.zip",
                        mime="application/zip",
                        use_container_width=True,
                    )
                else:
                    st.error(
                        "❌ Aucun modèle Word trouvé pour les lignes sélectionnées."
                    )

# -------------------------------------------------------------
# TAB 3 : DI MULTI-DATES
# -------------------------------------------------------------
with tab3:
    st.markdown(
        "##### 📅 **Génération des Demandes d'Intervention (DI) en PDF**"
    )

    dates_disponibles = sorted(
        {
            str(d).strip()
            for d in df["DATE"].unique()
            if str(d).strip() and str(d).lower() != "nan"
        }
    ) if "DATE" in df.columns else []

    dates_choisies = st.multiselect(
        "🗓️ Sélectionner une ou plusieurs dates :",
        options=dates_disponibles,
    )

    if dates_choisies and st.button(
        "📑 Générer DI Globale en PDF",
        type="primary",
    ):
        modele_di = os.path.join(
            os.path.dirname(CHEMIN_EXCEL_DEFAUT), "Demande d'intervention.docx"
        )
        if not os.path.exists(modele_di):
            modele_di = os.path.join(
                os.path.dirname(CHEMIN_EXCEL_DEFAUT),
                "Demande_intervention.docx",
            )

        if os.path.exists(modele_di):
            zip_buffer = io.BytesIO()
            has_pdf = False

            with zipfile.ZipFile(
                zip_buffer, "w", zipfile.ZIP_DEFLATED
            ) as zip_file:
                for d_single in dates_choisies:
                    df_sub = df[df["DATE"].astype(str).str.strip() == d_single]
                    if df_sub.empty:
                        continue

                    doc_rempli = generer_di_style_vba(modele_di, df_sub)

                    with tempfile.TemporaryDirectory() as temp_dir:
                        docx_path = os.path.join(temp_dir, "temp_di.docx")
                        pdf_path = os.path.join(temp_dir, "temp_di.pdf")
                        doc_rempli.save(docx_path)

                        pdf_bytes = convert_docx_to_pdf(docx_path, pdf_path)
                        nom_fichier = f"DI_Globale_{d_single.replace('/', '-')}"

                        if pdf_bytes:
                            zip_file.writestr(f"{nom_fichier}.pdf", pdf_bytes)
                            has_pdf = True
                        else:
                            with open(docx_path, "rb") as f_docx:
                                zip_file.writestr(
                                    f"{nom_fichier}.docx", f_docx.read()
                                )

            zip_buffer.seek(0)
            st.download_button(
                label="📦 Télécharger les DI (ZIP)",
                data=zip_buffer,
                file_name="DI_Globales_PDF.zip",
                mime="application/zip",
                use_container_width=True,
            )

            if not has_pdf:
                st.warning(
                    "⚠️ Conversion PDF impossible (Word/LibreOffice requis)."
                    " Fichiers Word (.docx) générés en secours."
                )
        else:
            st.error(
                "❌ Modèle `Demande d'intervention.docx` introuvable dans le"
                " dossier du projet."
            )
