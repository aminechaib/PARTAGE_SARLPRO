import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from io import BytesIO

st.set_page_config(layout="wide")
st.title("Smart Order Dispatcher")
# Upload files
orders_file = st.file_uploader("Upload Orders File (Excel)", type=["xlsx"])
stock_file = st.file_uploader("Upload Stock File (Excel)", type=["xlsx"])