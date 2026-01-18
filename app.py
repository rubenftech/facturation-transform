import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl.styles import PatternFill
import csv
import io

# ======================
# CONFIG PAGE
# ======================
st.set_page_config(
    page_title="Transformation de facturation",
    layout="wide"
)

st.title("🧾 Outil de transformation de facturation")
st.caption(
    "Importez vos fichiers de facturation. "
    "L’outil génère automatiquement un Excel final propre et exploitable."
)

# ======================
# UPLOAD FICHIERS
# ======================
doc1 = st.file_uploader(
    "📄 Fichier de facturation (Doc 1)",
    type=["csv", "xlsx"]
)

doc2 = st.file_uploader(
    "📄 Liste des raisons sociales (Doc 2)",
    type=["csv", "xlsx"]
)

# ======================
# LECTURE ROBUSTE CSV / EXCEL
# ======================
def read_file(file):
    try:
        if file.name.lower().endswith(".xlsx"):
            return pd.read_excel(file)

        raw = file.read()
        file.seek(0)

        try:
            text = raw.decode("utf-8")
        except UnicodeDecodeError:
            text = raw.decode("latin1")

        sniffer = csv.Sniffer()
        dialect = sniffer.sniff(text[:5000], delimiters=";,|\t")
        sep = dialect.delimiter

        return pd.read_csv(
            io.StringIO(text),
            sep=sep,
            engine="python",
            on_bad_lines="skip"
        )

    except Exception:
        raise ValueError(
            "Impossible de lire le fichier CSV ou Excel. "
            "Merci de vérifier le format du fichier."
        )

# ======================
# TRAITEMENT PRINCIPAL
# ======================
if doc1 and doc2:
    try:
        df = read_file(doc1)
        rs_df = read_file(doc2)
    except Exception as e:
        st.error(f"❌ Erreur lors de la lecture des fichiers : {e}")
        st.stop()

    # ======================
    # VÉRIFICATIONS
    # ======================
    if df.shape[1] < 10:
        st.error("❌ Le fichier de facturation ne contient pas assez de colonnes.")
        st.stop()

    if rs_df.shape[1] < 1:
        st.error("❌ Le fichier des raisons sociales est invalide.")
        st.stop()

    st.success(f"📊 Fichier chargé : {df.shape[0]} lignes – {df.shape[1]} colonnes")

    # ======================
    # NORMALISATION
    # ======================
    df.iloc[:, 1] = df.iloc[:, 1].astype(str).str.strip()
    rs_df.iloc[:, 0] = rs_df.iloc[:, 0].astype(str).str.strip()

    df.iloc[:, 9] = pd.to_numeric(df.iloc[:, 9], errors="coerce").fillna(0)

    # ======================
    # RÈGLE MÉTIER FIXE
    # ======================
    service_col = df.columns[3]  # Colonne D

    # ======================
    # FILTRAGE DE BASE
    # ======================
    base_df = df[
        (df.iloc[:, 4] != "NOT INJECTED") &
        (df.iloc[:, 9] > 0)
    ].copy()

    in_doc2 = base_df.iloc[:, 1].isin(rs_df.iloc[:, 0])

    is_sms = base_df[service_col] == "SMS"
    is_vocal = base_df[service_col] == "VOCAL"

    # ======================
    # SYNTHÈSE GLOBALE
    # ======================
    summary = pd.DataFrame({
        "Catégorie": [
            "SMS – Raisons sociales du doc 2",
            "SMS – Autres raisons sociales",
            "Vocal – Raisons sociales du doc 2",
            "Vocal – Autres raisons sociales"
        ],
        "Montant total": [
            base_df[is_sms & in_doc2].iloc[:, 9].sum(),
            base_df[is_sms & ~in_doc2].iloc[:, 9].sum(),
            base_df[is_vocal & in_doc2].iloc[:, 9].sum(),
            base_df[is_vocal & ~in_doc2].iloc[:, 9].sum(),
        ]
    })

    # ======================
    # FACTURATION DÉTAILLÉE (DOC 2)
    # ======================
    df_filtered = base_df[in_doc2]

    if df_filtered.empty:
        st.warning("⚠️ Aucune ligne retenue après filtrage.")
        st.stop()

    group_cols = [df.columns[1], df.columns[2]]
    agg_rules = {col: "first" for col in df.columns}
    agg_rules[df.columns[9]] = "sum"

    df_final = (
        df_filtered
        .groupby(group_cols, as_index=False)
        .agg(agg_rules)
        .sort_values(by=df.columns[1])
    )

    # ======================
    # APERÇU
    # ======================
    st.subheader("🔎 Aperçu de la facturation finale")
    st.dataframe(df_final, use_container_width=True)

    st.info(
        f"✔ {df_final.shape[0]} lignes générées\n"
        f"✔ {df_final.iloc[:, 1].nunique()} raisons sociales"
    )

    # ======================
    # EXPORT EXCEL
    # ======================
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_final.to_excel(writer, index=False, sheet_name="Facturation détaillée")
        summary.to_excel(writer, index=False, sheet_name="Synthèse globale")

        ws = writer.sheets["Facturation détaillée"]

        fill_a = PatternFill("solid", fgColor="EEEEEE")
        fill_b = PatternFill("solid", fgColor="FFFFFF")

        last_rs = None
        toggle = False

        for row in range(2, ws.max_row + 1):
            rs = ws.cell(row=row, column=2).value
            if rs != last_rs:
                toggle = not toggle
                last_rs = rs

            fill = fill_a if toggle else fill_b
            for col in range(1, ws.max_column + 1):
                ws.cell(row=row, column=col).fill = fill


    # ======================
    # APERÇU SYNTHÈSE GLOBALE
    # ======================
    st.subheader("📊 Synthèse globale")
    st.dataframe(summary, use_container_width=True)
    
    st.info(
        "Cette synthèse est incluse dans la deuxième feuille de l’Excel téléchargé."
    )



    st.download_button(
        "⬇️ Télécharger l’Excel final",
        data=output.getvalue(),
        file_name="facturation_finale.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
