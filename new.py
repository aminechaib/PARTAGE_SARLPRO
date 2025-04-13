import streamlit as st import pandas as pd import numpy as np import matplotlib.pyplot as plt import seaborn as sns from io import BytesIO

st.set_page_config(layout="wide") st.title("Smart Order Dispatcher")

Upload files

orders_file = st.file_uploader("Upload Orders File (Excel)", type=["xlsx"]) stock_file = st.file_uploader("Upload Stock File (Excel)", type=["xlsx"])

if orders_file and stock_file: # Read data orders_df = pd.read_excel(orders_file) stock_df = pd.read_excel(stock_file)

# Merge and compute allocation
merged_df = orders_df.merge(stock_df, on="Product")
merged_df["Total_Ordered"] = merged_df.groupby("Product")["Ordered_Qty"].transform("sum")
merged_df["Proportion"] = merged_df["Ordered_Qty"] / merged_df["Total_Ordered"]
merged_df["Allocated_Qty"] = (merged_df["Proportion"] * merged_df["Disponible_Qty"]).round().astype(int)

# Client satisfaction
client_satisfaction = merged_df.groupby("Client").agg(
    Total_Ordered=("Ordered_Qty", "sum"),
    Total_Allocated=("Allocated_Qty", "sum")
).reset_index()
client_satisfaction["Satisfaction_%"] = (client_satisfaction["Total_Allocated"] / client_satisfaction["Total_Ordered"] * 100).round(2)

# Product fulfillment
product_fulfillment = merged_df.groupby("Product").agg(
    Total_Ordered=("Ordered_Qty", "sum"),
    Total_Allocated=("Allocated_Qty", "sum")
).reset_index()
product_fulfillment["Fulfillment_%"] = (product_fulfillment["Total_Allocated"] / product_fulfillment["Total_Ordered"] * 100).round(2)

# Display data
st.subheader("Merged Allocation Table")
st.dataframe(merged_df)

st.subheader("Client Satisfaction")
st.dataframe(client_satisfaction)

st.subheader("Product Fulfillment")
st.dataframe(product_fulfillment)

# Charts
st.subheader("Charts")
col1, col2 = st.columns(2)
with col1:
    st.markdown("**Client Satisfaction Rate**")
    fig1, ax1 = plt.subplots(figsize=(10, 4))
    sns.barplot(data=client_satisfaction, x="Client", y="Satisfaction_%", palette="coolwarm", ax=ax1)
    plt.xticks(rotation=45)
    st.pyplot(fig1)

with col2:
    st.markdown("**Product Fulfillment Rate**")
    fig2, ax2 = plt.subplots(figsize=(10, 4))
    sns.barplot(data=product_fulfillment, x="Product", y="Fulfillment_%", palette="viridis", ax=ax2)
    plt.xticks(rotation=45)
    st.pyplot(fig2)

# Download button
st.subheader("Download Result Excel")
output = BytesIO()
with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
    merged_df.to_excel(writer, index=False, sheet_name="Allocation")
    client_satisfaction.to_excel(writer, index=False, sheet_name="ClientSatisfaction")
    product_fulfillment.to_excel(writer, index=False, sheet_name="ProductFulfillment")
    writer.save()
    st.download_button(label="Download Excel File",
                       data=output.getvalue(),
                       file_name="dispatch_results.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

else: st.info("Please upload both files to proceed.")

