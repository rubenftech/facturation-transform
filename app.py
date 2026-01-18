import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl.styles import PatternFill

st.set_page_config(
    page_title="Transformation de facturation",
    layout="wide"
)

st.title("🧾 Outil de transformation de facturation")
st.caption("Importez vos fichiers et téléchargez l'Excel final prêt à l'emploi.")

# ======================
# Upload fichiers
# ======================
doc1 = st.file_uploader("📄 Fichier de facturation (Doc 1)", type=["csv", "xlsx"])
doc2 = st.file_uploader("📄 Liste des raisons sociales (Doc 2)", type=["csv", "xlsx"])

def read_file(file):
    if file.name.endswith(".csv"):
        return pd.read_csv(file, engine="python")
    return pd.read_excel(file)

if doc1 and doc2:
    try:
        df = read_file(doc1)
        rs_df = read_file(doc2)
    except Exception as e:
        st.error(f"Erreur de lecture des fichiers : {e}")
        st.stop()

    # ======================
    # Sécurité colonnes
    # ======================
    if df.shape[1] < 10:
        st.error("Le fichier de facturation n'a pas assez de colonnes.")
        st.stop()

    # ======================
    # Normalisation
    # ======================
    df.iloc[:, 1] = df.iloc[:, 1].astype(str).str.strip()
    rs_df.iloc[:, 0] = rs_df.iloc[:, 0].astype(str).str.strip()

    df.iloc[:, 9] = pd.to_numeric(df.iloc[:, 9], errors="coerce").fillna(0)

    # ======================
    # Filtrage métier
    # ======================
    df_filtered = df[
        df.iloc[:, 1].isin(rs_df.iloc[:, 0]) &
        (df.iloc[:, 4] != "NOT INJECTED") &
        (df.iloc[:, 9] > 0)
    ]

    if df_filtered.empty:
        st.warning("Aucune ligne après filtrage.")
        st.stop()

    # ======================
    # Agrégation
    # ======================
    group_cols = [df.columns[1], df.columns[2]]

    agg_rules = {col: "first" for col in df.columns}
    agg_rules[df.columns[9]] = "sum"

    df_final = (
        df_filtered
        .groupby(group_cols, as_index=False)
        .agg(agg_rules)
        .sort_values(by=df.columns[1])
    )

    st.subheader("🔎 Aperçu du résultat")
    st.dataframe(df_final, use_container_width=True)

    # ======================
    # Export Excel avec couleurs
    # ======================
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_final.to_excel(writer, index=False, sheet_name="Facturation")
        ws = writer.sheets["Facturation"]

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

    st.download_button(
        "⬇️ Télécharger l'Excel final",
        data=output.getvalue(),
        file_name="facturation_finale.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
