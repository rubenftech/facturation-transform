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
st.caption("Importez vos fichiers de facturation puis cliquez sur Transformer.")

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

# ======================
# TRANSFORMATION
# ======================
if doc1 and doc2 and st.button("🚀 Transformer les fichiers"):
    with st.spinner("⏳ Transformation en cours…"):
        df = read_file(doc1)
        rs_df = read_file(doc2)

        # ======================
        # NORMALISATION
        # ======================
        df.iloc[:, 1] = df.iloc[:, 1].astype(str).str.strip()   # raison sociale
        rs_df.iloc[:, 0] = rs_df.iloc[:, 0].astype(str).str.strip()

        df.iloc[:, 9] = pd.to_numeric(df.iloc[:, 9], errors="coerce")

        # ======================
        # FILTRAGE DES LIGNES INVALIDES
        # ======================
        df = df[
            df.iloc[:, 6].notna() &     # Date d'opération non nulle
            df.iloc[:, 4].notna()       # Status non nul
        ].copy()

        base_df = df[
            (df.iloc[:, 4] != "NOT INJECTED") &
            (df.iloc[:, 9] > 0)
        ].copy()

        # ======================
        # FILTRE DOC 2
        # ======================
        base_df = base_df[
            base_df.iloc[:, 1].isin(rs_df.iloc[:, 0])
        ]

        # ======================
        # AGRÉGATION
        # ======================
        group_cols = [df.columns[1], df.columns[2]]  # raison sociale + numéro op

        agg = {
            df.columns[0]: "first",  # plateforme
            df.columns[1]: "first",  # raison sociale
            df.columns[2]: "first",  # numéro opération
            df.columns[3]: "first",  # type
            df.columns[4]: "first",  # status
            df.columns[5]: "first",  # nom opération
            df.columns[6]: "first",  # date
            df.columns[7]: "first",  # validation
            df.columns[8]: "first",  # pays
            df.columns[9]: "sum",    # nombre messages
        }

        df_final = (
            base_df
            .groupby(group_cols, as_index=False)
            .agg(agg)
        )

        # ======================
        # RENOMMAGE DES COLONNES (OUTPUT FINAL)
        # ======================
        df_final.columns = [
            "Plateforme",
            "Raison sociale",
            "Numéro d’opération",
            "Type",
            "Status",
            "Nom de l’opération",
            "Date d’opération",
            "Validation",
            "Pays",
            "Nombre de messages envoyés"
        ]

    # ======================
    # FEEDBACK UTILISATEUR
    # ======================
    st.success("✅ Transformation terminée avec succès")

    st.info(
        f"✔ {len(df_final):,}".replace(",", " ") + " lignes générées\n"
        f"✔ {df_final['Raison sociale'].nunique():,}".replace(",", " ") + " raisons sociales\n"
        f"✔ {df_final.shape[1]} colonnes en sortie"
    )

    # ======================
    # AFFICHAGE
    # ======================
    st.subheader("🔎 Facturation détaillée")
    st.dataframe(df_final, width="stretch")

    # ======================
    # EXPORT EXCEL
    # ======================
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_final.to_excel(writer, index=False, sheet_name="Facturation détaillée")

        ws = writer.sheets["Facturation détaillée"]
        fill_a = PatternFill("solid", fgColor="EEEEEE")
        fill_b = PatternFill("solid", fgColor="FFFFFF")

        last_rs, toggle = None, False
        for row in range(2, ws.max_row + 1):
            rs = ws.cell(row=row, column=2).value
            if rs != last_rs:
                toggle = not toggle
                last_rs = rs
            fill = fill_a if toggle else fill_b
            for col in range(1, ws.max_column + 1):
                ws.cell(row=row, column=col).fill = fill

    st.download_button(
        "⬇️ Télécharger l’Excel final",
        data=output.getvalue(),
        file_name="facturation_finale.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
