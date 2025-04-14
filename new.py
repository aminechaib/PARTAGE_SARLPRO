from streamlit_option_menu import option_menu

# Language dictionary
translations = {
    "en": {
        "title": "📦 Client Dispatch and Satisfaction Dashboard",
        "upload_orders": "Upload Orders File",
        "upload_stock": "Upload Stock File",
        "choose_client": "Choose Client",
        "edit_quantities": "✍️ Adjust Quantities for a Client",
        "dispatch_summary": "📋 Dispatch Summary",
        "satisfaction_chart": "📊 Client Satisfaction Overview",
        "fulfillment_pie": "🥧 Overall Fulfillment",
        "audit": "🧮 Stock vs Demand Audit",
        "download_report": "📥 Download Report",
        "success": "✅ Files loaded successfully!",
        "warning": "📂 Please upload both Orders and Stock files to continue.",
        "error": "❌ Error loading files"
    },
    "fr": {
        "title": "📦 Tableau de Répartition et Satisfaction Client",
        "upload_orders": "Télécharger le fichier de commandes",
        "upload_stock": "Télécharger le fichier de stock",
        "choose_client": "Choisir le client",
        "edit_quantities": "✍️ Ajuster les quantités pour un client",
        "dispatch_summary": "📋 Résumé de la répartition",
        "satisfaction_chart": "📊 Vue de satisfaction client",
        "fulfillment_pie": "🥧 Taux de satisfaction global",
        "audit": "🧮 Audit de stock vs demande",
        "download_report": "📥 Télécharger le rapport",
        "success": "✅ Fichiers chargés avec succès !",
        "warning": "📂 Veuillez télécharger les fichiers de commandes et de stock.",
        "error": "❌ Erreur lors du chargement des fichiers"
    }
}

# Select language
lang = option_menu(None, ["🇬🇧 English", "🇫🇷 Français"], orientation="horizontal")
lang_code = "en" if "English" in lang else "fr"

T = translations[lang_code]


import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from io import BytesIO

st.set_page_config(page_title="Client Dispatch Assistant", layout="wide")
st.set_page_config(
    page_title="Client Dispatch Assistant",
    layout="wide",
    page_icon="fav.png"  # 🧷 Add favicon here
)

# Optional: Show logo at the top
st.image("prg.png", width=250)


# Upload Excel files
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

        product_col = st.sidebar.selectbox("Select Product Column (Orders)", orders_columns)
        client_col = st.sidebar.selectbox("Select Client Column", orders_columns)
        qty_ordered_col = st.sidebar.selectbox("Select Ordered Quantity Column", orders_columns)
        stock_product_col = st.sidebar.selectbox("Select Product Column (Stock)", stock_columns)
        stock_qty_col = st.sidebar.selectbox("Select Stock Quantity Column", stock_columns)

        # Rename columns
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

        # Dispatch Calculation
        merged_df["Auto_Dispatch_Qty"] = 0
        for product, group in merged_df.groupby("Product"):
            total_ordered = group["Ordered_Qty"].sum()
            stock_match = stock_df[stock_df["Product"] == product]
            total_available = stock_match["Available_Qty"].sum() if not stock_match.empty else 0

            if total_ordered == 0 or total_available == 0:
                merged_df.loc[group.index, "Auto_Dispatch_Qty"] = 0
                continue

            allocated = []
            for _, row in group.iterrows():
                proportional_qty = (row["Ordered_Qty"] / total_ordered) * total_available
                to_dispatch = min(row["Ordered_Qty"], int(round(proportional_qty)))
                allocated.append(to_dispatch)

            total_allocated = sum(allocated)
            while total_allocated > total_available:
                for i in range(len(allocated)):
                    if allocated[i] > 0:
                        allocated[i] -= 1
                        total_allocated -= 1
                        if total_allocated <= total_available:
                            break

            merged_df.loc[group.index, "Auto_Dispatch_Qty"] = allocated

        # Initial value for editing
        merged_df["To_Give"] = merged_df["Auto_Dispatch_Qty"]

        # Client-wise adjustments
        st.subheader("✍️ Adjust Quantities for a Client")
        selected_client = st.selectbox("Choose Client", merged_df["Client"].unique())
        client_data = merged_df[merged_df["Client"] == selected_client].copy()

        st.markdown("### Edit ‘To_Give’ column only. You cannot exceed Ordered Quantity.")

        edited = st.data_editor(
            client_data[["Product", "Ordered_Qty", "Available_Qty", "To_Give"]],
            column_config={
                "To_Give": st.column_config.NumberColumn("To Give", min_value=0),
            },
            use_container_width=True,
            key="editor"
        )

        # Limit to Ordered_Qty
        for i, row in edited.iterrows():
            max_allowed = int(client_data.iloc[i]["Ordered_Qty"])
            if row["To_Give"] > max_allowed:
                edited.at[i, "To_Give"] = max_allowed

        # Update merged_df
        for i, row in edited.iterrows():
            idx = client_data.index[i]
            merged_df.at[idx, "To_Give"] = row["To_Give"]

        # Calculate Satisfaction
        merged_df["Satisfaction (%)"] = round((merged_df["To_Give"] / merged_df["Ordered_Qty"]) * 100, 2)
        merged_df["Satisfaction (%)"] = merged_df["Satisfaction (%)"].fillna(0)

        st.subheader("📋 Dispatch Summary")
        st.dataframe(merged_df)

        # 📊 Satisfaction Visualization
        st.subheader("📊 Client Satisfaction Overview")
        satisfaction_by_client = merged_df.groupby("Client")["Satisfaction (%)"].mean().reset_index()

        fig, ax = plt.subplots(figsize=(12, 6))
        bars = sns.barplot(data=satisfaction_by_client, x="Client", y="Satisfaction (%)", palette="coolwarm", ax=ax)
        ax.set_ylim(0, 110)
        ax.set_title("Client Satisfaction (%)")
        ax.set_ylabel("Satisfaction (%)")
        ax.set_xlabel("Client")
        plt.xticks(rotation=45)

        for p in bars.patches:
            height = p.get_height()
            bars.annotate(f'{height:.1f}%', (p.get_x() + p.get_width() / 2., height + 2),
                          ha='center', fontsize=9, color='black')

        st.pyplot(fig)
        plt.close(fig)

        # 🥧 Fulfillment Pie Chart
        st.subheader("🥧 Overall Fulfillment")
        total_ordered = merged_df["Ordered_Qty"].sum()
        total_given = merged_df["To_Give"].sum()
        fulfilled = total_given
        unfulfilled = max(0, total_ordered - total_given)

        pie_fig, pie_ax = plt.subplots()
        pie_ax.pie(
            [fulfilled, unfulfilled],
            labels=["Fulfilled", "Unfulfilled"],
            autopct='%1.1f%%',
            startangle=90,
            colors=["#2ecc71", "#e74c3c"],
            wedgeprops={'edgecolor': 'white'}
        )
        pie_ax.axis("equal")
        pie_ax.set_title("Total Order Fulfillment")
        st.pyplot(pie_fig)
        plt.close(pie_fig)

        # 🧾 Audit Stock & Demand
        st.subheader("🧮 Stock vs Demand Audit")
        stock_check = merged_df.groupby("Product").agg({
            "Ordered_Qty": "sum",
            "To_Give": "sum",
            "Available_Qty": "first"
        }).reset_index()

        stock_check["Unallocated_Stock"] = stock_check["Available_Qty"] - stock_check["To_Give"]
        stock_check["Unmet_Demand"] = stock_check["Ordered_Qty"] - stock_check["To_Give"]

        st.dataframe(stock_check)

        # Download
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
