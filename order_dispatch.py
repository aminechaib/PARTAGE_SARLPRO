import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from io import BytesIO
import numpy as np

# 🧷 Page Configuration
st.set_page_config(
    page_title="Client Dispatch Assistant",
    layout="wide",
    page_icon="fav.png"
)

# 🖼️ Logo
st.image("prg.png", width=250)

# 📁 File Upload
st.sidebar.header("📁 Upload Files")
orders_file = st.sidebar.file_uploader("Upload Orders File", type=["xlsx"])
stock_file = st.sidebar.file_uploader("Upload Stock File", type=["xlsx"])

if orders_file and stock_file:
    try:
        orders_df = pd.read_excel(orders_file)
        stock_df = pd.read_excel(stock_file)

        st.success("✅ Files loaded successfully!")

        # Column Mapping
        st.sidebar.subheader("🔧 Column Mapping")
        orders_columns = orders_df.columns.tolist()
        stock_columns = stock_df.columns.tolist()

        product_col = st.sidebar.selectbox("Product Column (Orders)", orders_columns)
        client_col = st.sidebar.selectbox("Client Column", orders_columns)
        qty_ordered_col = st.sidebar.selectbox("Ordered Quantity Column", orders_columns)
        vip_col = st.sidebar.selectbox("VIP Column (1=VIP, 0=Not)", orders_columns)
        stock_product_col = st.sidebar.selectbox("Product Column (Stock)", stock_columns)
        stock_qty_col = st.sidebar.selectbox("Stock Quantity Column", stock_columns)

        # Rename Columns
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

        # Ensure numeric
        orders_df["Ordered_Qty"] = pd.to_numeric(orders_df["Ordered_Qty"], errors="coerce").fillna(0)
        orders_df["VIP"] = pd.to_numeric(orders_df["VIP"], errors="coerce").fillna(0).astype(int)
        stock_df["Available_Qty"] = pd.to_numeric(stock_df["Available_Qty"], errors="coerce").fillna(0)

        # Merge Orders + Stock
        merged_df = orders_df.merge(stock_df, on="Product", how="left")
        merged_df["Available_Qty"] = merged_df["Available_Qty"].fillna(0)

        # Initialize dispatch column
        merged_df["Auto_Dispatch_Qty"] = 0

        # 🚚 Dispatch Calculation with VIP priority
        for product, group in merged_df.groupby("Product"):
            stock_qty = stock_df.loc[stock_df["Product"] == product, "Available_Qty"].sum()
            if stock_qty == 0:
                continue

            vip_group = group[group["VIP"] == 1].copy()
            regular_group = group[group["VIP"] == 0].copy()

            def allocate(group, stock_left, vip_boost=0):
                total_ordered = group["Ordered_Qty"].sum()
                if total_ordered == 0 or stock_left == 0:
                    return [0] * len(group), 0

                alloc = []
                for _, row in group.iterrows():
                    proportional = (row["Ordered_Qty"] / total_ordered) * stock_left
                    to_give = min(row["Ordered_Qty"], int(round(proportional))) + vip_boost
                    alloc.append(to_give)

                # Adjust if overallocated
                while sum(alloc) > stock_left:
                    for i in range(len(alloc)):
                        if alloc[i] > 0:
                            alloc[i] -= 1
                            if sum(alloc) <= stock_left:
                                break
                return alloc, sum(alloc)

            # Boost for VIP clients to increase their To_Give priority
            vip_alloc, vip_sum = allocate(vip_group, stock_qty, vip_boost=5)
            stock_left = stock_qty - vip_sum
            reg_alloc, _ = allocate(regular_group, stock_left)

            merged_df.loc[vip_group.index, "Auto_Dispatch_Qty"] = vip_alloc
            merged_df.loc[regular_group.index, "Auto_Dispatch_Qty"] = reg_alloc

        # Create To_Give for manual adjustment
        merged_df["To_Give"] = merged_df["Auto_Dispatch_Qty"]

        # ✍️ Client Adjustment UI
        st.subheader("✍️ Adjust Quantities for a Client")
        selected_client = st.selectbox("Choose Client", merged_df["Client"].unique())
        client_data = merged_df[merged_df["Client"] == selected_client].copy()

        st.markdown("### You can edit ‘To_Give’. Cannot exceed Ordered Quantity.")

        edited = st.data_editor(
            client_data[["Product", "Ordered_Qty", "Available_Qty", "VIP", "To_Give"]],
            column_config={
                "To_Give": st.column_config.NumberColumn("To Give", min_value=0),
            },
            use_container_width=True,
            key="editor"
        )

        for i, row in edited.iterrows():
            max_val = int(client_data.iloc[i]["Ordered_Qty"])
            if row["To_Give"] > max_val:
                edited.at[i, "To_Give"] = max_val

        for i, row in edited.iterrows():
            idx = client_data.index[i]
            merged_df.at[idx, "To_Give"] = row["To_Give"]

        # 💯 Satisfaction (Boosted for VIPs)
        merged_df["Satisfaction (%)"] = round((merged_df["To_Give"] / merged_df["Ordered_Qty"]) * 100, 2).fillna(0)

        # Ensure To_Give does not exceed Ordered_Qty
        merged_df["To_Give"] = merged_df.apply(lambda row: min(row["To_Give"], row["Ordered_Qty"]), axis=1)

        # Increase Satisfaction for VIPs
        merged_df["Satisfaction (%)"] = np.where(merged_df["VIP"] == 1, merged_df["Satisfaction (%)"] + 10, merged_df["Satisfaction (%)"])

        st.subheader("📋 Dispatch Summary")
        st.dataframe(merged_df)

        # 📊 Client Satisfaction
        st.subheader("📊 Client Satisfaction Overview")
        bar_data = merged_df.groupby(["Client", "VIP"])["Satisfaction (%)"].mean().reset_index()

        fig, ax = plt.subplots(figsize=(12, 6))
        sns.barplot(data=bar_data, x="Client", y="Satisfaction (%)", hue="VIP", ax=ax)
        ax.set_title("Client Satisfaction by VIP Status")
        ax.set_ylim(0, 110)
        plt.xticks(rotation=45)
        st.pyplot(fig)

        # 🥧 Fulfillment Pie
        st.subheader("🥧 Overall Fulfillment")
        total_ordered = merged_df["Ordered_Qty"].sum()
        total_given = merged_df["To_Give"].sum()

        fig2, ax2 = plt.subplots()
        ax2.pie(
            [total_given, total_ordered - total_given],
            labels=["Fulfilled", "Unfulfilled"],
            autopct='%1.1f%%',
            startangle=90,
            colors=["#2ecc71", "#e74c3c"],
            wedgeprops={"edgecolor": "white"}
        )
        ax2.axis("equal")
        ax2.set_title("Fulfillment Status")
        st.pyplot(fig2)

        # 📦 Stock Audit
        st.subheader("🧮 Stock vs Demand Audit")
        audit_df = merged_df.groupby("Product").agg({
            "Ordered_Qty": "sum",
            "To_Give": "sum",
            "Available_Qty": "first"
        }).reset_index()
        audit_df["Unallocated_Stock"] = audit_df["Available_Qty"] - audit_df["To_Give"]
        audit_df["Unmet_Demand"] = audit_df["Ordered_Qty"] - audit_df["To_Give"]
        st.dataframe(audit_df)

        # 📥 Download Button
        st.subheader("📥 Download Report")
        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
            merged_df.to_excel(writer, index=False, sheet_name="Dispatch")
        st.download_button(
            "Download Dispatch Report",
            data=buffer.getvalue(),
            file_name="dispatch_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"❌ Error loading files: {e}")
else:
    st.warning("📂 Please upload both Orders and Stock files to continue.")
