from datetime import datetime
from io import BytesIO
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Plateforme de Contrôle GFS", layout="wide"
)

st.title("🍽️ Plateforme de Contrôle - Global Food Safety Audit (GFS)")
st.sidebar.header("Navigation")

# Initialisation de l'historique des contrôles dans la session
if "evaluations" not in st.session_state:
    st.session_state.evaluations = []

# Sélection du module / section
section = st.sidebar.selectbox(
    "Choisir une section",
    [
        "Documentation",
        "Locaux (Facilities)",
        "Manipulation des aliments",
        "Hygiène personnelle",
    ],
)


# Fonction pour afficher une consigne avec choix (Oui/Non/N/A) et champ de commentaire
def render_question(q_id, question_text):
    st.markdown(f"**{question_text}**")
    col1, col2 = st.columns([1, 2])
    with col1:
        statut = st.radio(
            "Conformité",
            ["Oui", "Non", "N/A"],
            horizontal=True,
            key=f"status_{q_id}",
        )
    with col2:
        commentaire = st.text_input(
            "Observations / Action corrective",
            key=f"comm_{q_id}",
            placeholder="Précisez l'anomalie ou l'action...",
        )
    st.markdown("---")
    return {"id": q_id, "question": question_text, "statut": statut, "commentaire": commentaire}


# Formulaire selon la section
form_data = []

if section == "Documentation":
    st.header("📂 Sécurité sanitaire des aliments : Documentation")
    form_data.append(
        render_question(
            "FSDOC001",
            "1. Les procédures de notification des intoxications alimentaires sont documentées et accessibles.",
        )
    )
    form_data.append(
        render_question(
            "FSDOC002",
            "2. Le document de formation des employés (manipulateurs) est en cours de validité.",
        )
    )
    form_data.append(
        render_question(
            "FSDOC003",
            "3. Les responsables ont reçu une certification après avoir suivi un programme approuvé.",
        )
    )
    form_data.append(
        render_question(
            "FSDOC004", "4. Un programme de lutte contre les nuisibles est en place."
        )
    )

elif section == "Locaux (Facilities)":
    st.header("🏗️ Sécurité sanitaire des aliments : Locaux")
    form_data.append(
        render_question(
            "FSFAC001",
            "10. Les systèmes d'aération et hottes sont nettoyés par des professionnels.",
        )
    )
    form_data.append(
        render_question(
            "FSFAC002",
            "12. Les sols, murs et plafonds sont propres (pas de poussière, débris, eau stagnante).",
        )
    )
    form_data.append(
        render_question(
            "FSFAC003",
            "22. Les aliments et emballages sont stockés à au moins 15 cm du sol.",
        )
    )

elif section == "Manipulation des aliments":
    st.header("🥩 Sécurité sanitaire des aliments : Manipulation des aliments")
    form_data.append(
        render_question(
            "FSFH001",
            "31. Les aliments et surfaces sont protégés des dangers microbiologiques, physiques et chimiques.",
        )
    )
    form_data.append(
        render_question(
            "FSFH002", "32. Les aliments à risque sont décongelés selon les règles."
        )
    )

elif section == "Hygiène personnelle":
    st.header("🧼 Sécurité sanitaire des aliments : Hygiène personnelle")
    form_data.append(
        render_question(
            "FSPH001",
            "39. Les employés se lavent fréquemment et correctement les mains.",
        )
    )
    form_data.append(
        render_question(
            "FSPH002",
            "40. Aucun contact à mains nues n'a été constaté avec les produits prêts à consommer.",
        )
    )

# Bouton pour enregistrer les données de la section
if st.button("Enregistrer les contrôles de cette section"):
    date_jour = datetime.now().strftime("%Y-%m-%d %H:%M")
    for item in form_data:
        st.session_state.evaluations.append(
            {
                "Date": date_jour,
                "Section": section,
                "Référence": item["id"],
                "Consigne": item["question"],
                "Statut": item["statut"],
                "Commentaire": item["commentaire"],
            }
        )
    st.success("Données enregistrées avec succès pour cette section !")

# Section Exportation des résultats
st.sidebar.markdown("---")
st.sidebar.subheader("📥 Exportation du Rapport")

if len(st.session_state.evaluations) > 0:
    df_resultats = pd.DataFrame(st.session_state.evaluations)

    # Affichage d'un aperçu dans l'application
    with st.expander("Voir le tableau récapitulatif des contrôles"):
        st.dataframe(df_resultats)

    # Export Excel avec openpyxl
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_resultats.to_excel(writer, index=False, sheet_name="Rapport_GFS")
    excel_data = output.getvalue()

    st.sidebar.download_button(
        label="Télécharger le rapport en Excel (.xlsx)",
        data=excel_data,
        file_name=f"Rapport_GFS_{datetime.now().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    # Export CSV
    csv_data = df_resultats.to_csv(index=False).encode("utf-8")
    st.sidebar.download_button(
        label="Télécharger le rapport en CSV",
        data=csv_data,
        file_name=f"Rapport_GFS_{datetime.now().strftime('%Y%m%d')}.csv",
        mime="text/csv",
    )
else:
    st.sidebar.info(
        "Effectuez et enregistrez des contrôles pour activer l'exportation."
    )
