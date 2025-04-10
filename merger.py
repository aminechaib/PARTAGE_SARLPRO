import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Excel Data Merger", layout="wide")
st.title("📊 Excel Data Merger: Match & Fill Quantities Based on Reference")

# === File Upload ===
file1 = st.file_uploader("📁 Upload File 1 (Main Table)", type=["xlsx"])
file2 = st.file_uploader("📁 Upload File 2 (Source Data)", type=["xlsx"])

if file1 and file2:
    try:
        df1 = pd.read_excel(file1)
        df2 = pd.read_excel(file2)
        st.success("✅ Files loaded successfully!")

        # Column Selection
        st.subheader("Step 1️⃣: Match Columns Between Files")

        col1, col2 = st.columns(2)

        with col1:
            file1_ref_col = st.selectbox("🔑 Select Reference Column from File 1", df1.columns)

        with col2:
            file2_ref_col = st.selectbox("🔑 Select Reference Column from File 2", df2.columns)

        file2_source_col = st.selectbox("📂 Select 'Source File' Column from File 2", df2.columns)
        file2_quantity_col = st.selectbox("🔢 Select 'Quantity' Column from File 2", df2.columns)

        # Source Columns Input
        st.subheader("Step 2️⃣: Define Source Headers (These will be new columns)")

        source_cols_input = st.text_input(
            "📝 Enter source headers (comma-separated, e.g. 117,226,306):",
            placeholder="117, 226, 306"
        )

        # Process
        if source_cols_input and st.button("🔄 Process and Merge"):
            try:
                file1_source_cols = [col.strip() for col in source_cols_input.split(',')]

                for col in file1_source_cols:
                    df1[col] = df1.apply(
                        lambda row: df2.loc[
                            (df2[file2_ref_col] == row[file1_ref_col]) &
                            (df2[file2_source_col] == int(col)),
                            file2_quantity_col
                        ].values[0] if not df2.loc[
                            (df2[file2_ref_col] == row[file1_ref_col]) &
                            (df2[file2_source_col] == int(col))
                        ].empty else '',
                        axis=1
                    )

                st.success("✅ Data processed successfully!")

                st.subheader("📥 Download Result")
                output = BytesIO()
                df1.to_excel(output, index=False, engine='openpyxl')
                output.seek(0)

                st.download_button(
                    label="⬇️ Download Filled Excel",
                    data=output,
                    file_name="filled_table.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

                st.subheader("📋 Preview Result")
                st.dataframe(df1)

            except Exception as e:
                st.error(f"⚠️ Error during processing: {str(e)}")

    except Exception as e:
        st.error(f"❌ Failed to load files: {str(e)}")