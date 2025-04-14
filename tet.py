import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO

st.set_page_config(page_title="Excel Matcher", layout="wide")

st.title("🔍 Excel Matcher App")

# Upload files
database_file = st.file_uploader("Upload the database Excel file", type=["xlsx"])
search_terms_file = st.file_uploader("Upload the search terms Excel file", type=["xlsx"])

if database_file and search_terms_file:
    database_df = pd.read_excel(database_file)
    search_terms_df = pd.read_excel(search_terms_file)

    st.success("Files uploaded successfully.")

    search_terms_columns = st.multiselect("Select columns for search terms", search_terms_df.columns.tolist())
    database_columns = st.multiselect("Select columns to search in", database_df.columns.tolist())
    output_columns = st.multiselect("Select columns to include in the output", database_df.columns.tolist())

    if st.button("Start Matching") and search_terms_columns and database_columns and output_columns:
        search_terms = {
            col: search_terms_df[col].fillna('').astype(str).tolist()
            for col in search_terms_columns
        }

        unique_search_terms = {
            col: list(dict.fromkeys(terms)) for col, terms in search_terms.items()
        }

        database_df = database_df.astype(str).fillna('')
        matched_results = []

        total_iterations = len(database_df) * sum(len(terms) for terms in search_terms.values())
        progress_bar = st.progress(0)
        progress_counter = 0

        for index, row in database_df.iterrows():
            for term_sets in zip(*search_terms.values()):
                if any(term in row[col] for col in database_columns for term in term_sets if term):
                    match_dict = {f'searched_ref_{i+1}': term for i, term in enumerate(term_sets)}
                    match_dict.update({col: row[col] for col in output_columns})
                    matched_results.append(match_dict)
                progress_counter += 1
                progress_bar.progress(min(progress_counter / total_iterations, 1.0))

        matched_df = pd.DataFrame(matched_results)

        for i, col in enumerate(search_terms_columns):
            matched_df[f'searched_ref_{i+1}'] = pd.Categorical(
                matched_df[f'searched_ref_{i+1}'],
                categories=unique_search_terms[col],
                ordered=True
            )

        sort_columns = [f'searched_ref_{i+1}' for i in range(len(search_terms_columns))]
        matched_df.sort_values(by=sort_columns, inplace=True)

        st.subheader("🎯 Matching Results")
        st.dataframe(matched_df.head(100))

        output = BytesIO()
        writer = pd.ExcelWriter(output, engine='xlsxwriter')

        def write_chunks(writer, df, sheet_name_base, max_rows=1048576):
            num_chunks = (len(df) // max_rows) + 1
            for i in range(num_chunks):
                start_row = i * max_rows
                end_row = (i + 1) * max_rows
                chunk_df = df.iloc[start_row:end_row]
                sheet_name = f"{sheet_name_base}_{i + 1}"
                chunk_df.to_excel(writer, sheet_name=sheet_name, index=False, startrow=1, header=False)

                workbook = writer.book
                worksheet = writer.sheets[sheet_name]

                header = sort_columns + output_columns
                for col_num, value in enumerate(header):
                    worksheet.write(0, col_num, value)

                searched_ref_format = workbook.add_format({'bold': True, 'font_color': 'blue'})
                normal_format = workbook.add_format({'font_color': 'black'})
                prev_searched_refs = [None] * len(search_terms_columns)

                for row_num, row_data in enumerate(chunk_df.values, start=1):
                    for i in range(len(search_terms_columns)):
                        if row_data[i] != prev_searched_refs[i]:
                            worksheet.write(row_num, i, row_data[i], searched_ref_format)
                            prev_searched_refs[i] = row_data[i]
                        else:
                            worksheet.write_blank(row_num, i, None, normal_format)
                    for col_num, value in enumerate(row_data[len(search_terms_columns):], start=len(search_terms_columns)):
                        worksheet.write(row_num, col_num, value, normal_format)

        write_chunks(writer, matched_df, "Results")
        writer.close()
        output.seek(0)

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"matched_results_{timestamp}.xlsx"
        st.download_button(
            label="📥 Download Results",
            data=output,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

else:
    st.info("Please upload both Excel files to get started.")
