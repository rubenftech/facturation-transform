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
    "Importez vos fichiers de facturation puis cliquez sur Transformer."
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
# BOUTON TRANSFORMER
# ======================
if doc1 and doc2:
    if st.button("🚀 Transformer les fichiers"):
        with st.spinner("⏳ Transformation en cours…"):
            try:
                df = read_file(doc1)
                rs_df = read_file(doc2)
            except Exception as e:
                st.error(f"❌ Erreur lecture fichiers : {e}")
                st.stop()

            # ======================
            # NORMALISATION
            # ======================
            df.iloc[:, 1] = df.iloc[:, 1].astype(str).str.strip()
            rs_df.iloc[:, 0] = rs_df.iloc[:, 0].astype(str).str.strip()
            df.iloc[:, 9] = pd.to_numeric(df.iloc[:, 9], errors="coerce").fillna(0)

            service_col = df.columns[3]

            base_df = df[
                (df.iloc[:, 4] != "NOT INJECTED") &
                (df.iloc[:, 9] > 0)
            ].copy()

            in_doc2 = base_df.iloc[:, 1].isin(rs_df.iloc[:, 0])
            is_sms = base_df[service_col] == "SMS"
            is_vocal = base_df[service_col] == "VOCAL"

            # ======================
            # SYNTHÈSE GLOBALE (NUMÉRIQUE)
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

            # Version affichage (format lisible)
            summary_display = summary.copy()
            summary_display["Montant total"] = summary_display["Montant total"].apply(
                lambda x: f"{int(x):,}".replace(",", " ")
            )

            # ======================
            # FACTURATION DÉTAILLÉE
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

        # ======================
        # FEEDBACK UTILISATEUR
        # ======================
        st.success("✅ Transformation terminée avec succès")

        st.info(
            f"✔ {len(df_final):,}".replace(",", " ") + " lignes générées\n"
            f"✔ {df_final.iloc[:, 1].nunique():,}".replace(",", " ") + " raisons sociales\n"
            f"✔ {df_final.shape[1]} colonnes en sortie"
        )

        # ======================
        # AFFICHAGE
        # ======================
        st.subheader("🔎 Facturation détaillée")
        st.dataframe(df_final, width="stretch")

        st.subheader("📊 Synthèse globale")
        st.dataframe(summary_display, width="stretch")

        st.info(
            "Les montants sont affichés avec un format lisible. "
            "La version numérique est conservée dans l’Excel."
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
