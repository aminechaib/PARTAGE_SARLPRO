import streamlit as st

# ✅ MUST be first Streamlit command
st.set_page_config(
    page_title="Client Dispatch Assistant",
    layout="wide",
    page_icon="fav.png"
)

# Imports
from streamlit_option_menu import option_menu
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from io import BytesIO
from matplotlib.backends.backend_pdf import PdfPages

# Show logo
st.image("prg.png", width=250)

# Title
st.title("📦 Client Dispatch and Satisfaction Dashboard")

# Upload
st.sidebar.header("📁 Upload Files")
orders_file = st.sidebar.file_uploader("Upload Orders File", type=["xlsx"])
stock_file = st.sidebar.file_uploader("Upload Stock File", type=["xlsx"])

if orders_file and stock_file:
    try:
        orders_df = pd.read_excel(orders_file)
        stock_df = pd.read_excel(stock_file)
        st.success("✅ Files loaded successfully!")

        # Column mapping
        st.sidebar.subheader("🔧 Column Mapping")
        orders_columns = orders_df.columns.tolist()
        stock_columns = stock_df.columns.tolist()

        product_col = st.sidebar.selectbox("Product Column (Orders)", orders_columns)
        client_col = st.sidebar.selectbox("Client Column", orders_columns)
        qty_ordered_col = st.sidebar.selectbox("Ordered Quantity Column", orders_columns)
        vip_col = st.sidebar.selectbox("VIP Flag Column", orders_columns)  # VIP Flag
        stock_product_col = st.sidebar.selectbox("Product Column (Stock)", stock_columns)
        stock_qty_col = st.sidebar.selectbox("Stock Quantity Column", stock_columns)

        # Rename columns
        orders_df = orders_df.rename(columns={
            product_col: "Product",
            client_col: "Client",
            qty_ordered_col: "Ordered_Qty",
            vip_col: "VIP"
        })
        stock_df = stock_df.rename(columns={
            stock_product_col: "Product",
            stock_qty_col: "Available_Qty"
        })

        # Merge Data
        merged_df = orders_df.merge(stock_df, on="Product", how="left")
        merged_df["Available_Qty"] = pd.to_numeric(merged_df["Available_Qty"], errors="coerce").fillna(0)
        merged_df["Ordered_Qty"] = pd.to_numeric(merged_df["Ordered_Qty"], errors="coerce").fillna(0)
        merged_df["VIP"] = pd.to_numeric(merged_df["VIP"], errors="coerce").fillna(0)

        # Dispatch Calculation with VIP priority
        merged_df["Auto_Dispatch_Qty"] = 0

        for product, group in merged_df.groupby("Product"):
            total_stock = stock_df[stock_df["Product"] == product]["Available_Qty"].sum()

            vip_group = group[group["VIP"] == 1].copy()
            normal_group = group[group["VIP"] == 0].copy()

            for subset in [vip_group, normal_group]:
                for i in subset.index:
                    if total_stock <= 0:
                        break
                    requested = merged_df.at[i, "Ordered_Qty"]
                    allocated = min(requested, total_stock)
                    merged_df.at[i, "Auto_Dispatch_Qty"] = allocated
                    total_stock -= allocated

        merged_df["To_Give"] = merged_df["Auto_Dispatch_Qty"]

        # Client selector
        st.subheader("✍️ Adjust Quantities for a Client")
        selected_client = st.selectbox("Choose Client", merged_df["Client"].unique())
        client_data = merged_df[merged_df["Client"] == selected_client].copy()

        st.markdown("### You can adjust 'To_Give'. Cannot exceed ordered quantity.")

        editor = st.data_editor(
            client_data[["Product", "Ordered_Qty", "Available_Qty", "To_Give"]],
            column_config={"To_Give": st.column_config.NumberColumn("To Give", min_value=0)},
            key="editor"
        )

        for i, row in editor.iterrows():
            max_allowed = client_data.iloc[i]["Ordered_Qty"]
            updated_val = min(row["To_Give"], max_allowed)
            merged_df.at[client_data.index[i], "To_Give"] = updated_val

        # Satisfaction Calculation
        merged_df["Satisfaction (%)"] = round((merged_df["To_Give"] / merged_df["Ordered_Qty"]) * 100, 2).fillna(0)

        # Display Dispatch Summary
        st.subheader("📋 Dispatch Summary")
        st.dataframe(merged_df)

        # Satisfaction Chart
        st.subheader("📊 Client Satisfaction Overview")
        satisfaction_by_client = merged_df.groupby("Client")["Satisfaction (%)"].mean().reset_index()

        fig, ax = plt.subplots(figsize=(10, 5))
        sns.barplot(data=satisfaction_by_client, x="Client", y="Satisfaction (%)", palette="viridis", ax=ax)
        ax.set_ylim(0, 110)
        ax.set_title("Client Satisfaction (%)")
        ax.set_xlabel("Client")
        ax.set_ylabel("Satisfaction (%)")
        for bar in ax.patches:
            ax.annotate(f'{bar.get_height():.1f}%', (bar.get_x() + bar.get_width() / 2, bar.get_height() + 1),
                        ha='center')
        plt.xticks(rotation=45)
        st.pyplot(fig)

        # Fulfillment Pie Chart
        st.subheader("🥧 Overall Fulfillment")
        total_ordered = merged_df["Ordered_Qty"].sum()
        total_given = merged_df["To_Give"].sum()
        fig2, ax2 = plt.subplots()
        ax2.pie(
            [total_given, total_ordered - total_given],
            labels=["Fulfilled", "Unfulfilled"],
            colors=["#2ecc71", "#e74c3c"],
            autopct="%1.1f%%",
            startangle=90,
            wedgeprops={'edgecolor': 'white'}
        )
        ax2.axis("equal")
        st.pyplot(fig2)

        # Stock Audit Table
        st.subheader("🧮 Stock vs Demand Audit")
        audit = merged_df.groupby("Product").agg({
            "Ordered_Qty": "sum",
            "To_Give": "sum",
            "Available_Qty": "first"
        }).reset_index()
        audit["Remaining_Stock"] = audit["Available_Qty"] - audit["To_Give"]
        audit["Unmet_Demand"] = audit["Ordered_Qty"] - audit["To_Give"]
        st.dataframe(audit)

        # Download All Tables as Excel
        tables_output = BytesIO()
        with pd.ExcelWriter(tables_output, engine="xlsxwriter") as writer:
            merged_df.to_excel(writer, sheet_name="Dispatch", index=False)
            audit.to_excel(writer, sheet_name="Audit", index=False)

        st.download_button(
            label="📥 Download All Tables (Excel)",
            data=tables_output.getvalue(),
            file_name="All_Tables_Dispatch_Audit.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # Download Charts as PDF
        chart_pdf = BytesIO()
        with PdfPages(chart_pdf) as pdf_pages:
            pdf_pages.savefig(fig)
            pdf_pages.savefig(fig2)

        st.download_button(
            label="📥 Download Charts (PDF)",
            data=chart_pdf.getvalue(),
            file_name="Dispatch_Charts.pdf",
            mime="application/pdf"
        )

    except Exception as e:
        st.error(f"❌ Error loading files: {e}")
else:
    st.warning("📂 Please upload both Orders and Stock files to continue.")
