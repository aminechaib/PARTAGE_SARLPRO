import streamlit as st
from streamlit_option_menu import option_menu
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from io import BytesIO
from matplotlib.backends.backend_pdf import PdfPages

# ✅ Must be the first Streamlit command
st.set_page_config(
    page_title="Client Dispatch Assistant",
    layout="wide",
    page_icon="fav.png"
)

# Show logo and title
st.image("prg.png", width=250)
st.title("📦 Client Dispatch and Satisfaction Dashboard")

# Sidebar: File uploads
st.sidebar.header("📁 Upload Files")
orders_file = st.sidebar.file_uploader("Upload Orders File", type=["xlsx"])
stock_file = st.sidebar.file_uploader("Upload Stock File", type=["xlsx"])

def dispatch_allocation(df, stock_df):
    df["Auto_Dispatch_Qty"] = 0
    for product, group in df.groupby("Product"):
        total_stock = stock_df.loc[stock_df["Product"] == product, "Available_Qty"].sum()

        for priority in [1, 0]:  # VIP first
            sub = group[group["VIP"] == priority]
            for i in sub.index:
                if total_stock <= 0:
                    break
                requested = df.at[i, "Ordered_Qty"]
                allocated = min(requested, total_stock)
                df.at[i, "Auto_Dispatch_Qty"] = allocated
                total_stock -= allocated

    df["To_Give"] = df["Auto_Dispatch_Qty"]
    return df

def generate_charts(satisfaction_by_client, total_ordered, total_given):
    fig, ax = plt.subplots(figsize=(10, 5))
    sns.barplot(data=satisfaction_by_client, x="Client", y="Satisfaction (%)", palette="viridis", ax=ax)
    ax.set_ylim(0, 110)
    ax.set_title("Client Satisfaction (%)")
    ax.set_xlabel("Client")
    ax.set_ylabel("Satisfaction (%)")
    for bar in ax.patches:
        ax.annotate(f'{bar.get_height():.1f}%', (bar.get_x() + bar.get_width() / 2, bar.get_height() + 1), ha='center')
    plt.xticks(rotation=45)

    fig2, ax2 = plt.subplots()
    ax2.pie([total_given, total_ordered - total_given],
            labels=["Fulfilled", "Unfulfilled"],
            colors=["#2ecc71", "#e74c3c"],
            autopct="%1.1f%%",
            startangle=90,
            wedgeprops={'edgecolor': 'white'})
    ax2.axis("equal")

    return fig, fig2

if orders_file and stock_file:
    try:
        # Load files
        orders_df = pd.read_excel(orders_file)
        stock_df = pd.read_excel(stock_file)

        # Sidebar column mapping
        st.sidebar.subheader("🔧 Column Mapping")
        orders_cols = orders_df.columns.tolist()
        stock_cols = stock_df.columns.tolist()

        product_col = st.sidebar.selectbox("Product Column (Orders)", orders_cols)
        client_col = st.sidebar.selectbox("Client Column", orders_cols)
        qty_ordered_col = st.sidebar.selectbox("Ordered Quantity Column", orders_cols)
        vip_col = st.sidebar.selectbox("VIP Flag Column", orders_cols)

        stock_product_col = st.sidebar.selectbox("Product Column (Stock)", stock_cols)
        stock_qty_col = st.sidebar.selectbox("Stock Quantity Column", stock_cols)

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

        if "Available_Qty" not in stock_df.columns:
            raise KeyError(f"'Available_Qty' column not found in stock file. Selected column was: '{stock_qty_col}'")

        # Numeric conversion
        for col in ["Ordered_Qty", "VIP"]:
            orders_df[col] = pd.to_numeric(orders_df[col], errors="coerce").fillna(0)
        stock_df["Available_Qty"] = pd.to_numeric(stock_df["Available_Qty"], errors="coerce").fillna(0)

        # Merge and dispatch
        merged_df = orders_df.merge(stock_df, on="Product", how="left")
        merged_df = dispatch_allocation(merged_df, stock_df)

        # Client Quantity Adjustment
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
            idx = row.name
            max_allowed = merged_df.loc[idx, "Ordered_Qty"]
            merged_df.loc[idx, "To_Give"] = min(row["To_Give"], max_allowed)

        # Satisfaction and Summary
        merged_df["Satisfaction (%)"] = round((merged_df["To_Give"] / merged_df["Ordered_Qty"]) * 100, 2).fillna(0)

        st.subheader("📋 Dispatch Summary")
        st.dataframe(merged_df)

        # Charts
        st.subheader("📊 Client Satisfaction Overview")
        satisfaction_by_client = merged_df.groupby("Client")["Satisfaction (%)"].mean().reset_index()
        total_ordered = merged_df["Ordered_Qty"].sum()
        total_given = merged_df["To_Give"].sum()
        fig, fig2 = generate_charts(satisfaction_by_client, total_ordered, total_given)
        st.pyplot(fig)

        st.subheader("🥧 Overall Fulfillment")
        st.pyplot(fig2)

        # Audit Table
        st.subheader("🧮 Stock vs Demand Audit")
        audit = merged_df.groupby("Product").agg({
            "Ordered_Qty": "sum",
            "To_Give": "sum",
            "Available_Qty": "first"
        }).reset_index()
        audit["Remaining_Stock"] = audit["Available_Qty"] - audit["To_Give"]
        audit["Unmet_Demand"] = audit["Ordered_Qty"] - audit["To_Give"]
        st.dataframe(audit)

        # Download Excel
        excel_output = BytesIO()
        with pd.ExcelWriter(excel_output, engine="xlsxwriter") as writer:
            merged_df.to_excel(writer, sheet_name="Dispatch", index=False)
            audit.to_excel(writer, sheet_name="Audit", index=False)

        st.download_button(
            label="📥 Download All Tables (Excel)",
            data=excel_output.getvalue(),
            file_name="All_Tables_Dispatch_Audit.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # Download Charts PDF
        pdf_output = BytesIO()
        with PdfPages(pdf_output) as pdf:
            pdf.savefig(fig)
            pdf.savefig(fig2)

        st.download_button(
            label="📥 Download Charts (PDF)",
            data=pdf_output.getvalue(),
            file_name="Dispatch_Charts.pdf",
            mime="application/pdf"
        )

    except Exception as e:
        st.error(f"❌ Error loading files: {e}")
else:
    st.warning("📂 Please upload both Orders and Stock files to continue.")