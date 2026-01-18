import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl.styles import PatternFill

st.set_page_config(page_title="Transformation Facturation", layout="wide")
st.title("🧾 Outil de transformation de facturation")

doc1 = st.file_uploader("📄 Fichier de facturation (Doc 1)", type=["csv", "xlsx"])
doc2 = st.file_uploader("📄 Liste des raisons sociales (Doc 2)", type=["csv", "xlsx"])

def read_file(file):
    if file.name.endswith(".csv"):
        return pd.read_csv(file, engine="python")
    return pd.read_excel(file)

if doc1 and doc2:
    df = read_file(doc1)
    rs_df = read_file(doc2)

    # Normalisation
    df.iloc[:, 1] = df.iloc[:, 1].astype(str).str.strip()
    rs_df.iloc[:, 0] = rs_df.iloc[:, 0].astype(str).str.strip()
    df.iloc[:, 9] = pd.to_numeric(df.iloc[:, 9], errors="coerce").fillna(0)

    # Filtrage de base (commun)
    base_df = df[
        (df.iloc[:, 4] != "NOT INJECTED") &
        (df.iloc[:, 9] > 0)
    ].copy()

    in_doc2 = base_df.iloc[:, 1].isin(rs_df.iloc[:, 0])
    service_col = base_df.columns[3]  # colonne type SMS / Vocal (à adapter si besoin)

    is_sms = base_df[service_col].str.contains("SMS", case=False, na=False)
    is_vocal = base_df[service_col].str.contains("VOCAL|VOICE", case=False, na=False)

    # ======================
    # Synthèse globale
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
    # Facturation détaillée (doc 2 uniquement)
    # ======================
    df_filtered = base_df[in_doc2]

    group_cols = [df.columns[1], df.columns[2]]
    agg_rules = {col: "first" for col in df.columns}
    agg_rules[df.columns[9]] = "sum"

    df_final = (
        df_filtered
        .groupby(group_cols, as_index=False)
        .agg(agg_rules)
        .sort_values(by=df.columns[1])
    )

    st.subheader("Aperçu facturation détaillée")
    st.dataframe(df_final, use_container_width=True)

    # ======================
    # Export Excel
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

    st.download_button(
        "⬇️ Télécharger l'Excel final",
        data=output.getvalue(),
        file_name="facturation_finale.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
