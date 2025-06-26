import streamlit as st
import tempfile
import os
from main2 import traiter_fichiers_MP1
from Test import traiter_fichiers_JLN
import openpyxl

# --- Personnalisation du thème via CSS ---
st.markdown(
    """
    <style>
    .stButton>button {
        background-color: #0A6EBD;
        color: white;
        font-weight: bold;
        border-radius: 8px;
        padding: 0.5em 2em;
        margin-top: 1em;
    }
    .stDownloadButton>button {
        background-color: #198754;
        color: white;
        font-weight: bold;
        border-radius: 8px;
        padding: 0.5em 2em;
    }
    .stRadio>div>label {
        font-size: 1.1em;
        font-weight: 500;
    }
    </style>
    """,
    unsafe_allow_html=True
)

# --- Logo (optionnel) ---
# Place un fichier 'logo.png' dans le dossier du projet si tu veux l'afficher
if os.path.exists("logo.png"):
    st.image("logo.png", width=120)

# --- Titres et sous-titres ---
st.title("Prétraitement de données brutes")
st.markdown("<h4 style='color:#0A6EBD;'>Automatisez le prétraitement de vos fichiers Excel MP1 & JLN</h4>", unsafe_allow_html=True)
st.markdown("---")

st.write("""
Bienvenue sur l'interface de prétraitement de données brutes. Sélectionnez le module à traiter, chargez vos fichiers, puis lancez le traitement. Le fichier prétraité sera généré automatiquement.
""")

# --- Choix du module ---
module = st.radio("Choisissez le module à traiter :", ("MP1", "JLN"))

if module == "MP1":
    st.markdown("## 🔵 Prétraitement MP1")
    st.info("Veuillez charger le fichier source et le fichier cible MP1.")
    fichier_source = st.file_uploader("Fichier source MP1", type=["xlsx"])
    fichier_cible = st.file_uploader("Fichier cible MP1", type=["xlsx"])
    if fichier_source and fichier_cible:
        col1, col2, col3 = st.columns([1,2,1])
        with col2:
            if st.button("Lancer le traitement MP1"):
                with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_source, \
                     tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_cible:
                    tmp_source.write(fichier_source.read())
                    tmp_cible.write(fichier_cible.read())
                    tmp_source.flush()
                    tmp_cible.flush()
                    with st.spinner("Traitement en cours... Merci de patienter."):
                        try:
                            fichier_sortie = traiter_fichiers_MP1(tmp_source.name, tmp_cible.name)
                            if fichier_sortie:
                                with open(fichier_sortie, "rb") as f:
                                    st.success("Traitement terminé ! Téléchargez le fichier prétraité ci-dessous.")
                                    st.download_button(
                                        label="Télécharger le fichier prétraité",
                                        data=f,
                                        file_name="Fichier Prétraité MP1.xlsx",
                                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                    )
                            else:
                                st.warning("Le traitement n'a pas pu être effectué. Vérifiez vos fichiers.")
                        except KeyError as e:
                            if "xl/drawings/NULL" in str(e):
                                st.error("Erreur : Le fichier Excel contient une référence à un dessin/image corrompu ou manquant. Ouvrez le fichier dans Excel, supprimez tous les objets graphiques (images, graphiques, etc.), puis réenregistrez-le.")
                            else:
                                st.error(f"Erreur inattendue : {e}")
                        except Exception as e:
                            st.error(f"Erreur inattendue : {e}")

elif module == "JLN":
    st.markdown("## 🟢 Prétraitement JLN")
    st.info("Veuillez charger les deux fichiers source (Production & Arrêts) et le fichier cible JLN.")
    fichier_source1 = st.file_uploader("Fichier source 1 (Production)", type=["xlsx"], key="src1")
    fichier_source2 = st.file_uploader("Fichier source 2 (Arrêts)", type=["xlsx"], key="src2")
    fichier_cible = st.file_uploader("Fichier cible JLN", type=["xlsx"], key="ciblejln")
    if fichier_source1 and fichier_source2 and fichier_cible:
        col1, col2, col3 = st.columns([1,2,1])
        with col2:
            if st.button("Lancer le traitement JLN"):
                with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_s1, \
                     tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_s2, \
                     tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_cible:
                    tmp_s1.write(fichier_source1.read())
                    tmp_s2.write(fichier_source2.read())
                    tmp_cible.write(fichier_cible.read())
                    tmp_s1.flush()
                    tmp_s2.flush()
                    tmp_cible.flush()
                    with st.spinner("Traitement en cours... Merci de patienter."):
                        fichier_sortie = traiter_fichiers_JLN(tmp_s1.name, tmp_s2.name, tmp_cible.name)
                        if fichier_sortie:
                            with open(fichier_sortie, "rb") as f:
                                st.success("Traitement terminé ! Téléchargez le fichier prétraité ci-dessous.")
                                st.download_button(
                                    label="Télécharger le fichier prétraité",
                                    data=f,
                                    file_name="Fichier Prétraité JLN.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                )
                        else:
                            st.warning("Le traitement n'a pas pu être effectué. Vérifiez vos fichiers.")



