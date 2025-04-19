import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from io import BytesIO
from streamlit_option_menu import option_menu

# ------------------- CONFIGURATION -------------------
st.set_page_config(
    page_title="Client Dispatch Assistant",
    layout="wide",
    page_icon="ac.png"  # Favicon
)
st.image("logo.png", width=250)

# ------------------- LANGUAGE SUPPORT -------------------
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
lang = option_menu(None, ["🇬🇧 English", "🇫🇷 Français"], orientation="horizontal")
lang_code = "en" if "English" in lang else "fr"
T = translations[lang_code]

st.title(T["title"])

# ------------------- FILE UPLOAD -------------------
st.sidebar.header("📁 Upload Files")
orders_file = st.sidebar.file_uploader(T["upload_orders"], type=["xlsx"])
stock_file = st.sidebar.file_uploader(T["upload_stock"], type=["xlsx"])

if orders_file and stock_file:
    try:
        orders_df = pd.read_excel(orders_file)
        stock_df = pd.read_excel(stock_file)

        st.success(T["success"])

        # Column mapping
        st.sidebar.subheader("🔧 Column Mapping")
        orders_columns = orders_df.columns.tolist()
        stock_columns = stock_df.columns.tolist()

        product_col = st.sidebar.selectbox("Select Product Column (Orders)", orders_columns)
        client_col = st.sidebar.selectbox("Select Client Column", orders_columns)
        qty_ordered_col = st.sidebar.selectbox("Select Ordered Quantity Column", orders_columns)
        vip_col = st.sidebar.selectbox("Select VIP Column (Yes/No)", orders_columns)
        stock_product_col = st.sidebar.selectbox("Select Product Column (Stock)", stock_columns)
        stock_qty_col = st.sidebar.selectbox("Select Stock Quantity Column", stock_columns)

        # Rename for processing
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

        # Clean & convert
        orders_df["VIP"] = orders_df["VIP"].fillna("No")
        orders_df["Ordered_Qty"] = pd.to_numeric(orders_df["Ordered_Qty"], errors="coerce").fillna(0)
        stock_df["Available_Qty"] = pd.to_numeric(stock_df["Available_Qty"], errors="coerce").fillna(0)

        # Merge
        merged_df = orders_df.merge(stock_df, on="Product", how="left").fillna(0)
        merged_df["Auto_Dispatch_Qty"] = 0

        # Allocation logic
        for product, group in merged_df.groupby("Product"):
            available = stock_df[stock_df["Product"] == product]["Available_Qty"].sum()
            group = group.sort_values(by="VIP", ascending=False)  # VIP first
            total_ordered = group["Ordered_Qty"].sum()

            if total_ordered == 0 or available == 0:
                continue

            allocated = []
            for _, row in group.iterrows():
                proportional_qty = (row["Ordered_Qty"] / total_ordered) * available
                to_give = min(int(round(proportional_qty)), row["Ordered_Qty"])
                allocated.append(to_give)

            while sum(allocated) > available:
                for i in range(len(allocated)):
                    if allocated[i] > 0:
                        allocated[i] -= 1
                        if sum(allocated) <= available:
                            break

            merged_df.loc[group.index, "Auto_Dispatch_Qty"] = allocated

        merged_df["To_Give"] = merged_df["Auto_Dispatch_Qty"]

        # Client editor
        st.subheader(T["edit_quantities"])
        selected_client = st.selectbox(T["choose_client"], merged_df["Client"].unique())
        client_data = merged_df[merged_df["Client"] == selected_client].copy()

        edited = st.data_editor(
            client_data[["Product", "Ordered_Qty", "Available_Qty", "To_Give"]],
            column_config={"To_Give": st.column_config.NumberColumn("To Give", min_value=0)},
            use_container_width=True,
            key="editor"
        )

        for i, row in edited.iterrows():
            idx = client_data.index[i]
            max_allowed = int(client_data.iloc[i]["Ordered_Qty"])
            merged_df.at[idx, "To_Give"] = min(row["To_Give"], max_allowed)

        # Satisfaction
        merged_df["Satisfaction (%)"] = round((merged_df["To_Give"] / merged_df["Ordered_Qty"]) * 100, 2).fillna(0)

        st.subheader(T["dispatch_summary"])
        st.dataframe(merged_df)

        # Visualization
        st.subheader(T["satisfaction_chart"])
        sat = merged_df.groupby("Client")["Satisfaction (%)"].mean().reset_index()
        fig, ax = plt.subplots(figsize=(12, 6))
        sns.barplot(data=sat, x="Client", y="Satisfaction (%)", ax=ax, palette="coolwarm")
        plt.xticks(rotation=45)
        plt.title("Client Satisfaction (%)")
        st.pyplot(fig)

        st.subheader(T["fulfillment_pie"])
        fulfilled = merged_df["To_Give"].sum()
        total = merged_df["Ordered_Qty"].sum()
        pie_fig, pie_ax = plt.subplots()
        pie_ax.pie([fulfilled, total - fulfilled], labels=["Fulfilled", "Unfulfilled"], autopct='%1.1f%%',
                   startangle=90, colors=["#2ecc71", "#e74c3c"])
        pie_ax.axis("equal")
        st.pyplot(pie_fig)

        st.subheader(T["audit"])
        audit = merged_df.groupby("Product").agg({
            "Ordered_Qty": "sum",
            "To_Give": "sum",
            "Available_Qty": "first"
        }).reset_index()
        audit["Unallocated_Stock"] = audit["Available_Qty"] - audit["To_Give"]
        audit["Unmet_Demand"] = audit["Ordered_Qty"] - audit["To_Give"]
        st.dataframe(audit)

        st.subheader(T["download_report"])
        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
            merged_df.to_excel(writer, index=False, sheet_name="Dispatch")
        st.download_button(T["download_report"], buffer.getvalue(), "dispatch_report.xlsx")

    except Exception as e:
        st.error(f"{T['error']}: {e}")

else:
    st.warning(T["warning"])
