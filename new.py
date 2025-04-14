import streamlit as st

# ✅ MUST be first Streamlit command
st.set_page_config(
    page_title="Client Dispatch Assistant",
    layout="wide",
    page_icon="fav.png"
)

# Other imports
from streamlit_option_menu import option_menu
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from io import BytesIO

# Show logo
st.image("prg.png", width=250)

# Language translations
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

# Language selection
lang = option_menu(None, ["🇬🇧 English", "🇫🇷 Français"], orientation="horizontal")
lang_code = "en" if "English" in lang else "fr"
T = translations[lang_code]

st.title(T["title"])

# Upload
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

        product_col = st.sidebar.selectbox("Product Column (Orders)", orders_columns)
        client_col = st.sidebar.selectbox("Client Column", orders_columns)
        qty_ordered_col = st.sidebar.selectbox("Ordered Quantity Column", orders_columns)
        stock_product_col = st.sidebar.selectbox("Product Column (Stock)", stock_columns)
        stock_qty_col = st.sidebar.selectbox("Stock Quantity Column", stock_columns)

        # Rename
        orders_df = orders_df.rename(columns={
            product_col: "Product",
            client_col: "Client",
            qty_ordered_col: "Ordered_Qty"
        })
        stock_df = stock_df.rename(columns={
            stock_product_col: "Product",
            stock_qty_col: "Available_Qty"
        })

        # Merge
        merged_df = orders_df.merge(stock_df, on="Product", how="left")
        merged_df["Available_Qty"] = pd.to_numeric(merged_df["Available_Qty"], errors="coerce").fillna(0)
        merged_df["Ordered_Qty"] = pd.to_numeric(merged_df["Ordered_Qty"], errors="coerce").fillna(0)

        # Auto Dispatch Calculation
        merged_df["Auto_Dispatch_Qty"] = 0
        for product, group in merged_df.groupby("Product"):
            total_ordered = group["Ordered_Qty"].sum()
            stock_row = stock_df[stock_df["Product"] == product]
            total_stock = stock_row["Available_Qty"].sum() if not stock_row.empty else 0

            if total_ordered == 0 or total_stock == 0:
                merged_df.loc[group.index, "Auto_Dispatch_Qty"] = 0
                continue

            allocations = []
            for _, row in group.iterrows():
                prop = (row["Ordered_Qty"] / total_ordered) * total_stock
                to_give = min(row["Ordered_Qty"], int(round(prop)))
                allocations.append(to_give)

            # Adjustment loop
            total_allocated = sum(allocations)
            while total_allocated > total_stock:
                for i in range(len(allocations)):
                    if allocations[i] > 0:
                        allocations[i] -= 1
                        total_allocated -= 1
                        if total_allocated <= total_stock:
                            break

            merged_df.loc[group.index, "Auto_Dispatch_Qty"] = allocations

        # Set editable column
        merged_df["To_Give"] = merged_df["Auto_Dispatch_Qty"]

        # Client selector
        st.subheader(T["edit_quantities"])
        selected_client = st.selectbox(T["choose_client"], merged_df["Client"].unique())
        client_data = merged_df[merged_df["Client"] == selected_client].copy()

        st.markdown("### ✨ You can adjust 'To_Give'. Cannot exceed ordered quantity.")

        editor = st.data_editor(
            client_data[["Product", "Ordered_Qty", "Available_Qty", "To_Give"]],
            column_config={
                "To_Give": st.column_config.NumberColumn("To Give", min_value=0)
            },
            key="editor"
        )

        # Apply edits
        for i, row in editor.iterrows():
            max_allowed = client_data.iloc[i]["Ordered_Qty"]
            updated_val = min(row["To_Give"], max_allowed)
            merged_df.at[client_data.index[i], "To_Give"] = updated_val

        # Satisfaction calculation
        merged_df["Satisfaction (%)"] = round((merged_df["To_Give"] / merged_df["Ordered_Qty"]) * 100, 2).fillna(0)

        st.subheader(T["dispatch_summary"])
        st.dataframe(merged_df)

        # Satisfaction Chart
        st.subheader(T["satisfaction_chart"])
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
        st.subheader(T["fulfillment_pie"])
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

        # Stock Audit
        st.subheader(T["audit"])
        audit = merged_df.groupby("Product").agg({
            "Ordered_Qty": "sum",
            "To_Give": "sum",
            "Available_Qty": "first"
        }).reset_index()
        audit["Remaining_Stock"] = audit["Available_Qty"] - audit["To_Give"]
        audit["Unmet_Demand"] = audit["Ordered_Qty"] - audit["To_Give"]
        st.dataframe(audit)

        # Download report
        st.subheader(T["download_report"])
        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            merged_df.to_excel(writer, sheet_name="Dispatch", index=False)
        st.download_button(
            label="📥 Download Dispatch Report",
            data=output.getvalue(),
            file_name="dispatch_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"{T['error']}: {e}")
else:
    st.warning(T["warning"])
