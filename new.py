import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from io import BytesIO

st.set_page_config(page_title="Client Dispatch Assistant", layout="wide")
st.title("📦 Client Dispatch and Satisfaction Dashboard")

# Upload Excel files
st.sidebar.header("📁 Upload Files")
orders_file = st.sidebar.file_uploader("Upload Orders File", type=["xlsx"])
stock_file = st.sidebar.file_uploader("Upload Stock File", type=["xlsx"])

if orders_file and stock_file:
    try:
        orders_df = pd.read_excel(orders_file)
        stock_df = pd.read_excel(stock_file)

        st.success("Files loaded successfully!")

        # Show columns to let user pick relevant ones
        st.sidebar.subheader("🔧 Column Mapping")
        st.sidebar.write("Select matching columns from your files")

        orders_columns = orders_df.columns.tolist()
        stock_columns = stock_df.columns.tolist()

        product_col = st.sidebar.selectbox("Select Product Column", orders_columns)
        client_col = st.sidebar.selectbox("Select Client Column", orders_columns)
        qty_ordered_col = st.sidebar.selectbox("Select Ordered Qty Column", orders_columns)
        stock_product_col = st.sidebar.selectbox("Select Product Column in Stock", stock_columns)
        stock_qty_col = st.sidebar.selectbox("Select Stock Quantity Column", stock_columns)

        # Rename for merge
        orders_df = orders_df.rename(columns={
            product_col: "Product",
            client_col: "Client",
            qty_ordered_col: "Ordered_Qty"
        })

        stock_df = stock_df.rename(columns={
            stock_product_col: "Product",
            stock_qty_col: "Available_Qty"
        })

        # Merge and clean
        merged_df = orders_df.merge(stock_df, on="Product", how="left")
        merged_df["Available_Qty"] = pd.to_numeric(merged_df["Available_Qty"], errors="coerce").fillna(0)
        merged_df["Ordered_Qty"] = pd.to_numeric(merged_df["Ordered_Qty"], errors="coerce").fillna(0)

        # Initialize column
        merged_df["Auto_Dispatch_Qty"] = 0

        # Proportional dispatch logic (group-wise)
        for product, group in merged_df.groupby("Product"):
            total_available = group["Available_Qty"].iloc[0]
            total_ordered = group["Ordered_Qty"].sum()

            if total_ordered == 0 or total_available == 0:
                merged_df.loc[group.index, "Auto_Dispatch_Qty"] = 0
                continue

            # Proportional allocation capped at order qty
            allocated = []
            for idx, row in group.iterrows():
                proportional_qty = (row["Ordered_Qty"] / total_ordered) * total_available
                to_dispatch = min(row["Ordered_Qty"], int(round(proportional_qty)))
                allocated.append(to_dispatch)

            # Fix any over-allocation due to rounding
            total_allocated = sum(allocated)
            while total_allocated > total_available:
                for i in range(len(allocated)):
                    if allocated[i] > 0:
                        allocated[i] -= 1
                        total_allocated -= 1
                        if total_allocated <= total_available:
                            break

            merged_df.loc[group.index, "Auto_Dispatch_Qty"] = allocated

        # Final To_Give column
        merged_df["To_Give"] = merged_df["Auto_Dispatch_Qty"]

        # Interface: Client Selection
        st.subheader("🚚 Assign Quantities for Selected Client")
        selected_client = st.selectbox("Select Client to Assign Quantities", merged_df["Client"].unique())
        client_data = merged_df[merged_df["Client"] == selected_client].copy()

        # Allow manual override of quantities
        st.write("### ✍️ Adjust Quantities")
        new_quantities = []
        for idx, row in client_data.iterrows():
            new_qty = st.number_input(
                f"{row['Product']} - Ordered: {row['Ordered_Qty']}, Available: {row['Available_Qty']}",
                value=int(row["To_Give"]),
                min_value=0,
                max_value=int(row["Ordered_Qty"])
            )
            new_quantities.append(new_qty)

        merged_df.loc[client_data.index, "To_Give"] = new_quantities

        # Recalculate satisfaction
        merged_df["Satisfaction (%)"] = round((merged_df["To_Give"] / merged_df["Ordered_Qty"]) * 100, 2)
        merged_df["Satisfaction (%)"] = merged_df["Satisfaction (%)"].fillna(0)

        st.subheader("📋 Full Merged View")
        st.dataframe(merged_df)

        # 📊 Client Satisfaction Overview - Improved Visual
        st.subheader("📊 Client Satisfaction Breakdown")

        fig, ax = plt.subplots(figsize=(12, 6))
        satisfaction = merged_df.groupby("Client")["Satisfaction (%)"].mean().reset_index()
        bars = sns.barplot(data=satisfaction, x="Client", y="Satisfaction (%)", palette="coolwarm", ax=ax)
        ax.set_ylim(0, 110)
        ax.set_title("Average Satisfaction per Client", fontsize=16)
        ax.set_ylabel("Satisfaction (%)")
        ax.set_xlabel("Client")
        plt.xticks(rotation=45)

        # Annotate bars
        for p in bars.patches:
            height = p.get_height()
            bars.annotate(f'{height:.1f}%', (p.get_x() + p.get_width() / 2., height),
                          ha='center', va='bottom', fontsize=10, color='black')

        st.pyplot(fig)

        # Download result
        to_download = merged_df.copy()
        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
            to_download.to_excel(writer, index=False, sheet_name="Dispatch")
        st.download_button(
            "📥 Download Dispatch Report",
            data=buffer.getvalue(),
            file_name="dispatch_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"❌ Error loading files: {e}")

else:
    st.warning("📂 Please upload both the Orders and Stock files to proceed.")
