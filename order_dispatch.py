import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from io import BytesIO
from streamlit_option_menu import option_menu

# Set page config
st.set_page_config(page_title="Client Dispatch Assistant", layout="wide", page_icon="fav.png")

# Logo
st.image("logo.png", width=250)

# Language support
def get_translations():
    return {
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

# Language selection
translations = get_translations()
lang = option_menu(None, ["\U0001F1EC\U0001F1E7 English", "\U0001F1EB\U0001F1F7 Français"], orientation="horizontal")
lang_code = "en" if "English" in lang else "fr"
T = translations[lang_code]

# Sidebar file upload
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
        vip_col = st.sidebar.selectbox("Select VIP Column", orders_columns)
        stock_product_col = st.sidebar.selectbox("Select Product Column (Stock)", stock_columns)
        stock_qty_col = st.sidebar.selectbox("Select Stock Quantity Column", stock_columns)

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

        # Clean data
        orders_df["VIP"] = orders_df["VIP"].fillna("No").str.strip().str.lower()
        orders_df["VIP"] = orders_df["VIP"].apply(lambda x: "Yes" if x == "yes" else "No")

        orders_df["Ordered_Qty"] = pd.to_numeric(orders_df["Ordered_Qty"], errors="coerce").fillna(0)
        stock_df["Available_Qty"] = pd.to_numeric(stock_df["Available_Qty"], errors="coerce").fillna(0)

        merged_df = orders_df.merge(stock_df, on="Product", how="left")

        # Allocation with VIP priority
        merged_df["Auto_Dispatch_Qty"] = 0
        for product, group in merged_df.groupby("Product"):
            available = stock_df[stock_df["Product"] == product]["Available_Qty"].sum()
            if available == 0:
                continue

            group = group.sort_values(by="VIP", ascending=False)
            allocated = []
            for _, row in group.iterrows():
                need = int(row["Ordered_Qty"])
                give = min(need, available)
                allocated.append(give)
                available -= give

            merged_df.loc[group.index, "Auto_Dispatch_Qty"] = allocated

        merged_df["To_Give"] = merged_df["Auto_Dispatch_Qty"]

        # Manual adjustment per client
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
            max_allowed = int(client_data.iloc[i]["Ordered_Qty"])
            if row["To_Give"] > max_allowed:
                edited.at[i, "To_Give"] = max_allowed

        for i, row in edited.iterrows():
            idx = client_data.index[i]
            merged_df.at[idx, "To_Give"] = row["To_Give"]

        # Satisfaction
        merged_df["Satisfaction (%)"] = round((merged_df["To_Give"] / merged_df["Ordered_Qty"]) * 100, 2).fillna(0)

        st.subheader(T["dispatch_summary"])
        st.dataframe(merged_df)

        # 📊 Client Satisfaction
        st.subheader(T["satisfaction_chart"])
        satisfaction_by_client = merged_df.groupby("Client")["Satisfaction (%)"].mean().reset_index()
        fig, ax = plt.subplots(figsize=(12, 6))
        bars = sns.barplot(data=satisfaction_by_client, x="Client", y="Satisfaction (%)", palette="coolwarm", ax=ax)
        ax.set_ylim(0, 110)
        plt.xticks(rotation=45)
        for p in bars.patches:
            height = p.get_height()
            bars.annotate(f'{height:.1f}%', (p.get_x() + p.get_width() / 2., height + 2), ha='center', fontsize=9)
        st.pyplot(fig)

        # 🥧 Fulfillment Pie Chart
        st.subheader(T["fulfillment_pie"])
        total_ordered = merged_df["Ordered_Qty"].sum()
        total_given = merged_df["To_Give"].sum()
        fulfilled = total_given
        unfulfilled = max(0, total_ordered - total_given)
        pie_fig, pie_ax = plt.subplots()
        pie_ax.pie([fulfilled, unfulfilled], labels=["Fulfilled", "Unfulfilled"], autopct='%1.1f%%', startangle=90,
                   colors=["#2ecc71", "#e74c3c"], wedgeprops={'edgecolor': 'white'})
        pie_ax.axis("equal")
        st.pyplot(pie_fig)

        # 🧾 Stock vs Demand Audit
        st.subheader(T["audit"])
        stock_check = merged_df.groupby("Product").agg({
            "Ordered_Qty": "sum",
            "To_Give": "sum",
            "Available_Qty": "first"
        }).reset_index()
        stock_check["Unallocated_Stock"] = stock_check["Available_Qty"] - stock_check["To_Give"]
        stock_check["Unmet_Demand"] = stock_check["Ordered_Qty"] - stock_check["To_Give"]
        st.dataframe(stock_check)

        # Download Excel
        st.subheader(T["download_report"])
        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
            merged_df.to_excel(writer, index=False, sheet_name="Dispatch")
        st.download_button("Download Dispatch Report", data=buffer.getvalue(), file_name="dispatch_report.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    except Exception as e:
        st.error(f"{T['error']}: {e}")
else:
    st.warning(T["warning"])
