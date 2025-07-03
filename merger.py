import streamlit as st
import pandas as pd
import os
from io import BytesIO
import zipfile

# --- Page Setup ---
st.set_page_config(
    page_title="Excel Tools",
    page_icon="fav.png",
    layout="wide"
)

st.image("prg.png", width=200)
st.title("📊 Excel Tools")

# === Tabs ===
tab3, tab1, tab2, tab4 = st.tabs([
    "🛠 Convert XLS ➜ XLSX",
    "📦 Merge Multiple Excel Files",
    "🔁 Match & Merge Two Files by Reference",
    "📊 Pivot-style Merger (Group & Aggregate)"
])


# === Tab 3: XLS to XLSX Converter ===
with tab3:
    st.header("🛠 Convert `.xls` ➜ `.xlsx`")
    xls_files = st.file_uploader("📁 Upload `.xls` files (older Excel format)", type=["xls"], accept_multiple_files=True)

    if xls_files:
        converted_files = []
        for xls_file in xls_files:
            try:
                df = pd.read_excel(xls_file, engine="xlrd")
                output = BytesIO()
                df.to_excel(output, index=False, engine='openpyxl')
                output.seek(0)
                converted_files.append((xls_file.name.replace(".xls", ".xlsx"), output))
                st.success(f"✅ Converted: {xls_file.name}")
            except Exception as e:
                st.error(f"❌ Failed to convert {xls_file.name}: {e}")

        if converted_files:
            zip_buffer = BytesIO()
            with zipfile.ZipFile(zip_buffer, "w") as zipf:
                for filename, file_data in converted_files:
                    zipf.writestr(filename, file_data.read())
            zip_buffer.seek(0)

            st.download_button(
                label="⬇️ Download All Converted Files (ZIP)",
                data=zip_buffer,
                file_name="converted_xlsx_files.zip",
                mime="application/zip"
            )
    else:
        st.info("📌 Upload one or more `.xls` files to convert them to `.xlsx`.")

# === Tab 1: Merge ===
with tab1:
    st.header("📦 Merge Multiple .xlsx Files")
    uploaded_files = st.file_uploader("📁 Upload one or more Excel files (.xlsx)", type=["xlsx"], accept_multiple_files=True)

    if uploaded_files:
        df_list = []
        for uploaded_file in uploaded_files:
            try:
                df = pd.read_excel(uploaded_file)
                df['file name'] = uploaded_file.name
                df_list.append(df)
            except Exception as e:
                st.error(f"❌ Failed to read {uploaded_file.name}: {e}")

        if df_list:
            merged_df = pd.concat(df_list, ignore_index=True)
            st.success("✅ Files merged successfully!")

            st.subheader("📋 Preview Merged Data")
            st.dataframe(merged_df, use_container_width=True)

            output = BytesIO()
            merged_df.to_excel(output, index=False, engine='openpyxl')
            output.seek(0)

            st.download_button(
                label="⬇️ Download Merged Excel",
                data=output,
                file_name="merged_data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("⚠️ No valid data found in uploaded files.")
    else:
        st.info("📝 Please drag and drop `.xlsx` files to begin.")

# === Tab 2: Match & Merge ===
with tab2:
    st.header("🔁 Match & Merge Two Files Based on Reference")
    file1 = st.file_uploader("📁 Upload File 1 (Main Table) 'reference dispo'", type=["xlsx"], key="file1")
    file2 = st.file_uploader("📁 Upload File 2 (Source Data)", type=["xlsx"], key="file2")

    if file1 and file2:
        try:
            df1 = pd.read_excel(file1)
            df2 = pd.read_excel(file2)
            st.success("✅ Files loaded successfully!")

            st.subheader("Step 1️⃣: Match Columns Between Files")
            col1, col2 = st.columns(2)

            with col1:
                file1_ref_col = st.selectbox("🔑 Reference Column from File 1", df1.columns)

            with col2:
                file2_ref_col = st.selectbox("🔑 Reference Column from File 2", df2.columns)

            file2_source_col = st.selectbox("📂 Select 'Source File' Column from File 2", df2.columns)
            file2_quantity_col = st.selectbox("🔢 Select 'Quantity' Column from File 2", df2.columns)

            st.subheader("Step 2️⃣: Define Target Headers for File 1")
            source_cols_input = st.text_input(
                "📝 Enter source headers (comma-separated, e.g. 117,226,306):",
                placeholder="117, 226, 306"
            )

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

                    st.success("✅ Data matched and merged successfully!")

                    output = BytesIO()
                    df1.to_excel(output, index=False, engine='openpyxl')
                    output.seek(0)

                    st.subheader("📥 Download Result")
                    st.download_button(
                        label="⬇️ Download Filled Excel",
                        data=output,
                        file_name="matched_result.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

                    st.subheader("📋 Preview Result")
                    st.dataframe(df1, use_container_width=True)

                except Exception as e:
                    st.error(f"⚠️ Error during processing: {str(e)}")
        except Exception as e:
            st.error(f"❌ Failed to read files: {str(e)}")
with tab4:
    st.header("📊 Pivot-style Merger (Group & Aggregate)")

    pivot_files = st.file_uploader("📁 Upload Excel files to group and aggregate", type=["xlsx"], accept_multiple_files=True, key="pivot")

    if pivot_files:
        df_list = []
        for f in pivot_files:
            try:
                df = pd.read_excel(f)
                df['source_file'] = f.name
                df_list.append(df)
            except Exception as e:
                st.error(f"❌ Error reading {f.name}: {e}")

        if df_list:
            merged = pd.concat(df_list, ignore_index=True)
            st.success("✅ Files loaded and merged successfully.")
            st.subheader("🔍 Preview of Combined Data")
            st.dataframe(merged.head(10), use_container_width=True)

            all_columns = merged.columns.tolist()
            numeric_cols = merged.select_dtypes(include=['number']).columns.tolist()

            group_cols = st.multiselect("🔗 Group By Columns", all_columns)
            agg_options = ["sum", "mean", "count", "max", "min"]

            agg_config = {}
            st.subheader("🔣 Choose Aggregations")

            for col in numeric_cols:
                agg_choice = st.selectbox(f"📌 Aggregate `{col}` by:", ["(skip)"] + agg_options, key=col)
                if agg_choice != "(skip)":
                    agg_config[col] = agg_choice

            if group_cols and agg_config and st.button("🔄 Run Aggregation"):
                try:
                    grouped = merged.groupby(group_cols).agg(agg_config).reset_index()
                    st.success("✅ Aggregation completed!")

                    st.subheader("📋 Aggregated Result")
                    st.dataframe(grouped, use_container_width=True)

                    output = BytesIO()
                    grouped.to_excel(output, index=False, engine='openpyxl')
                    output.seek(0)

                    st.download_button(
                        label="⬇️ Download Aggregated Excel",
                        data=output,
                        file_name="pivot_aggregated.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                except Exception as e:
                    st.error(f"⚠️ Aggregation error: {str(e)}")
            else:
                st.info("ℹ️ Please select group and aggregation columns.")
    else:
        st.info("📂 Upload `.xlsx` files to begin.")