import streamlit as st
import pandas as pd
import os
from io import BytesIO
import zipfile
import tempfile

# === Setup ===
st.set_page_config(
    page_title="Excel Data Merger",
    page_icon="fav.png",  # Make sure fav.png is in the same folder
    layout="wide"
)

st.image("prg.png", width=200)
st.title("📊 Excel Data Merger")
st.markdown("Match references and fill quantities with ease. Upload files or process multiple `.xls` files at once.")

st.markdown("---")

# === DRAG & DROP ===
tab1, tab2 = st.tabs(["🔁 Match & Merge Two Files", "📂 Merge Multiple `.xls` Files"])

# ========== TAB 1 ==========
with tab1:
    st.subheader("Step 1️⃣: Upload Files to Match and Merge")

    col1, col2 = st.columns(2)
    with col1:
        file1 = st.file_uploader("📁 Drag & Drop or Browse File 1 (Main Table)", type=["xlsx"])
    with col2:
        file2 = st.file_uploader("📁 Drag & Drop or Browse File 2 (Source Data)", type=["xlsx"])

    if file1 and file2:
        try:
            df1 = pd.read_excel(file1)
            df2 = pd.read_excel(file2)
            st.success("✅ Files loaded successfully!")

            st.subheader("Step 2️⃣: Select Matching Columns")

            c1, c2 = st.columns(2)
            with c1:
                file1_ref_col = st.selectbox("🔑 Reference Column in File 1", df1.columns)
            with c2:
                file2_ref_col = st.selectbox("🔑 Reference Column in File 2", df2.columns)

            file2_source_col = st.selectbox("📂 Source Column in File 2", df2.columns)
            file2_quantity_col = st.selectbox("🔢 Quantity Column in File 2", df2.columns)

            st.subheader("Step 3️⃣: Enter Source Headers (New Columns to Fill)")
            source_cols_input = st.text_input("📝 Enter values like `117, 226, 306`")

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

                    # Download button
                    st.subheader("📥 Download Filled Excel File")
                    output = BytesIO()
                    df1.to_excel(output, index=False, engine='openpyxl')
                    output.seek(0)
                    st.download_button(
                        label="⬇️ Download",
                        data=output,
                        file_name="filled_table.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

                    st.subheader("📋 Preview Merged Table")
                    st.dataframe(df1)

                except Exception as e:
                    st.error(f"⚠️ Error: {str(e)}")

        except Exception as e:
            st.error(f"❌ Could not load files: {str(e)}")

# ========== TAB 2 ==========
with tab2:
    st.subheader("Merge Multiple `.xls` Files (Batch Mode)")

    zip_file = st.file_uploader("📦 Upload ZIP file containing `.xls` files", type=["zip"])
    
    if zip_file:
        try:
            with tempfile.TemporaryDirectory() as tmpdirname:
                with zipfile.ZipFile(zip_file, "r") as zip_ref:
                    zip_ref.extractall(tmpdirname)

                # Read all .xls files
                files = [f for f in os.listdir(tmpdirname) if f.endswith('.xls')]
                df_list = []

                for file in files:
                    path = os.path.join(tmpdirname, file)
                    try:
                        df = pd.read_excel(path, engine='xlrd')
                        df['file name'] = file
                        df_list.append(df)
                    except Exception as e:
                        st.warning(f"⚠️ Failed to read {file}: {e}")

                if df_list:
                    merged_df = pd.concat(df_list, ignore_index=True)

                    st.success("✅ Files merged successfully!")
                    output = BytesIO()
                    merged_df.to_excel(output, index=False)
                    output.seek(0)

                    st.download_button(
                        label="⬇️ Download Merged File",
                        data=output,
                        file_name="merged_data.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

                    st.subheader("📋 Preview Merged Data")
                    st.dataframe(merged_df)
                else:
                    st.error("❌ No valid .xls files found in ZIP.")

        except Exception as e:
            st.error(f"❌ Failed to process ZIP file: {str(e)}")

# === Footer ===
st.markdown("---")
st.caption("🔧 Created by Amine Auto | 📌 Need help? Contact support.")
