# app.py
# ============================================================
# CLIENT ACTIVITY ANALYTICS DASHBOARD
# Streamlit + Pandas + Plotly
# ============================================================

import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO

# ------------------------------------------------------------
# PAGE CONFIG & STATE
# ------------------------------------------------------------

st.set_page_config(
    page_title="Client Activity Dashboard",
    page_icon="📊",
    layout="wide"
)

# This function forces the app to reset whenever you upload or remove a file
def reset_app():
    st.session_state.dashboard_ready = False

if "dashboard_ready" not in st.session_state:
    st.session_state.dashboard_ready = False

# ------------------------------------------------------------
# STYLE
# ------------------------------------------------------------

st.markdown("""
<style>
.main {
    background-color: #f5f7fa;
}
.block-container {
    padding-top: 1rem;
}
.metric-card {
    background: white;
    padding: 15px;
    border-radius: 12px;
    box-shadow: 0px 2px 10px rgba(0,0,0,0.08);
}
.mapping-label {
    font-size: 16px;
    font-weight: bold;
    margin-top: 10px;
}
</style>
""", unsafe_allow_html=True)

# ------------------------------------------------------------
# TITLE
# ------------------------------------------------------------

st.title("📊 Client Activity Analytics Dashboard")
st.markdown("Analyze client invoices, yearly activity, monthly movement and chiffre d'affaires.")

# ------------------------------------------------------------
# 1. FILE UPLOAD
# ------------------------------------------------------------

# The on_change=reset_app ensures the form ALWAYS shows up when you upload a file
uploaded_file = st.file_uploader(
    "Upload Excel File",
    type=["xlsx", "xls"],
    on_change=reset_app
)

if uploaded_file:

    # --------------------------------------------------------
    # 2. READ ALL COLUMN NAMES
    # --------------------------------------------------------

    try:
        raw_df = pd.read_excel(uploaded_file)
    except Exception as e:
        st.error(f"Error reading file: {e}")
        st.stop()

    excel_columns = raw_df.columns.tolist()

    # --------------------------------------------------------
    # 3. SHOW MAPPING INTERFACE
    # --------------------------------------------------------

    if not st.session_state.dashboard_ready:
        st.subheader("⚙️ Map Excel Columns")
        st.write("Please select which column from your Excel file corresponds to each dashboard field.")

        with st.form("mapping_form"):
            col1, col2 = st.columns([1, 2])
            
            with col1:
                st.markdown("<div class='mapping-label'>Dashboard Field</div>", unsafe_allow_html=True)
                st.markdown("<div style='margin-top: 25px;'>Client Code</div>", unsafe_allow_html=True)
                st.markdown("<div style='margin-top: 35px;'>Client Name</div>", unsafe_allow_html=True)
                st.markdown("<div style='margin-top: 35px;'>Invoice ID</div>", unsafe_allow_html=True)
                st.markdown("<div style='margin-top: 35px;'>Invoice Date</div>", unsafe_allow_html=True)
                st.markdown("<div style='margin-top: 35px;'>Amount (Montant)</div>", unsafe_allow_html=True)
                
            with col2:
                st.markdown("<div class='mapping-label'>Select Excel Column</div>", unsafe_allow_html=True)
                map_code = st.selectbox("Client Code", excel_columns, label_visibility="collapsed")
                map_name = st.selectbox("Client Name", excel_columns, label_visibility="collapsed")
                map_inv_id = st.selectbox("Invoice ID", excel_columns, label_visibility="collapsed")
                map_date = st.selectbox("Invoice Date", excel_columns, label_visibility="collapsed")
                map_amount = st.selectbox("Amount", excel_columns, label_visibility="collapsed")

            st.write("") 
            submitted = st.form_submit_button("🚀 Generate Dashboard", type="primary")

            if submitted:
                # Validate that the user didn't select the same column twice
                selected_cols = [map_code, map_name, map_inv_id, map_date, map_amount]
                if len(set(selected_cols)) < 5:
                    st.error("⚠️ Error: You selected the same Excel column multiple times. Please map each field to a different column.")
                else:
                    st.session_state.mapping = {
                        "Code Client": map_code,
                        "Nom Client": map_name,
                        "Facture id": map_inv_id,
                        "Date Fact.": map_date,
                        "Montant TTC": map_amount
                    }
                    st.session_state.dashboard_ready = True
                    st.rerun() 

        # This st.stop() physically prevents the dashboard from running until the form is successfully submitted
        st.stop()

    # --------------------------------------------------------
    # 4. LOAD DASHBOARD
    # --------------------------------------------------------

    if st.session_state.dashboard_ready:
        
        if st.button("🔄 Edit Column Mapping"):
            st.session_state.dashboard_ready = False
            st.rerun()
            
        st.divider()

        mapping = st.session_state.mapping
        df = raw_df.rename(columns={
            mapping["Code Client"]: "Code Client",
            mapping["Nom Client"]: "Nom Client",
            mapping["Facture id"]: "Facture id",
            mapping["Date Fact."]: "Date Fact.",
            mapping["Montant TTC"]: "Montant TTC"
        })

        df = df[["Code Client", "Nom Client", "Facture id", "Date Fact.", "Montant TTC"]]

        # --------------------------------------------------------
        # CLEANING
        # --------------------------------------------------------

        df["Date Fact."] = pd.to_datetime(df["Date Fact."], errors="coerce")
        df = df.dropna(subset=["Date Fact."])

        df["Montant TTC"] = df["Montant TTC"].astype(str).str.replace(r'[^\d.,]', '', regex=True)
        df["Montant TTC"] = df["Montant TTC"].str.replace(',', '.')
        df["Montant TTC"] = pd.to_numeric(df["Montant TTC"], errors="coerce").fillna(0)

        df["Year"] = df["Date Fact."].dt.year
        df["Month"] = df["Date Fact."].dt.strftime("%B")

        # --------------------------------------------------------
        # SIDEBAR FILTERS
        # --------------------------------------------------------

        st.sidebar.header("🔎 Filters")

        years = sorted(df["Year"].dropna().unique())

        selected_years = st.sidebar.multiselect(
            "Select Year(s)",
            years,
            default=years
        )

        filtered_df = df[df["Year"].isin(selected_years)]

        clients = sorted(filtered_df["Nom Client"].dropna().unique())

        selected_clients = st.sidebar.multiselect(
            "Select Client(s)",
            clients,
            default=clients
        )

        filtered_df = filtered_df[
            filtered_df["Nom Client"].isin(selected_clients)
        ]

        if filtered_df.empty:
            st.warning("No data available for the selected filters.")
            st.stop()

        # --------------------------------------------------------
        # KPI METRICS
        # --------------------------------------------------------

        total_clients = filtered_df["Nom Client"].nunique()
        total_factures = filtered_df["Facture id"].nunique()
        total_ca = filtered_df["Montant TTC"].sum()

        c1, c2, c3 = st.columns(3)

        c1.metric("Clients", f"{total_clients:,}")
        c2.metric("Factures", f"{total_factures:,}")
        c3.metric("Chiffre d'Affaires", f"{total_ca:,.2f} DA")

        st.divider()

        # --------------------------------------------------------
        # CLIENT SUMMARY TABLE
        # --------------------------------------------------------

        st.subheader("📋 Client Activity Summary")

        summary = (
            filtered_df
            .groupby(["Code Client", "Nom Client", "Year"])
            .agg(
                Nombre_Factures=("Facture id", "nunique"),
                Chiffre_Affaire=("Montant TTC", "sum"),
                Mois_Activite=("Month",
                               lambda x: ", ".join(sorted(set(x))))
            )
            .reset_index()
        )

        summary = summary.sort_values(
            by="Chiffre_Affaire",
            ascending=False
        )

        st.dataframe(
            summary,
            use_container_width=True,
            height=500
        )

        # --------------------------------------------------------
        # YEARLY FACTURES CHART
        # --------------------------------------------------------

        st.subheader("📈 Factures by Year")

        year_chart = (
            filtered_df
            .groupby("Year")["Facture id"]
            .nunique()
            .reset_index()
        )

        fig_year = px.bar(
            year_chart,
            x="Year",
            y="Facture id",
            text_auto=True,
            title="Total Factures per Year"
        )

        st.plotly_chart(fig_year, use_container_width=True)

        # --------------------------------------------------------
        # MONTHLY ACTIVITY
        # --------------------------------------------------------

        st.subheader("📅 Monthly Client Activity")

        monthly = (
            filtered_df
            .groupby(["Year", "Month"])
            .agg(
                Factures=("Facture id", "nunique"),
                CA=("Montant TTC", "sum")
            )
            .reset_index()
        )

        month_order = [
            "January", "February", "March", "April",
            "May", "June", "July", "August",
            "September", "October", "November", "December"
        ]

        monthly["Month"] = pd.Categorical(
            monthly["Month"],
            categories=month_order,
            ordered=True
        )

        monthly = monthly.sort_values(["Year", "Month"])

        fig_month = px.line(
            monthly,
            x="Month",
            y="Factures",
            color="Year",
            markers=True,
            title="Monthly Facture Activity"
        )

        st.plotly_chart(fig_month, use_container_width=True)

        # --------------------------------------------------------
        # TOP CLIENTS
        # --------------------------------------------------------

        st.subheader("🏆 Top Clients by Chiffre d'Affaires")

        top_clients = (
            filtered_df
            .groupby(["Code Client", "Nom Client"])["Montant TTC"]
            .sum()
            .reset_index()
            .sort_values(by="Montant TTC", ascending=False)
            .head(15)
        )

        fig_top = px.bar(
            top_clients,
            x="Nom Client",
            y="Montant TTC",
            text_auto=".2s",
            title="Top Clients"
        )

        st.plotly_chart(fig_top, use_container_width=True)

        # --------------------------------------------------------
        # EXPORT EXCEL
        # --------------------------------------------------------

        st.subheader("⬇ Export Summary")

        output = BytesIO()

        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            summary.to_excel(writer, sheet_name="Client Summary", index=False)
            monthly.to_excel(writer, sheet_name="Monthly Activity", index=False)
            top_clients.to_excel(writer, sheet_name="Top Clients", index=False)

        st.download_button(
            label="📥 Download Excel Report",
            data=output.getvalue(),
            file_name="client_activity_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

else:
    st.info("Upload an Excel file to start.")
