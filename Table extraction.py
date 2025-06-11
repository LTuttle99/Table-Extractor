import streamlit as st
import pandas as pd
import openpyxl

st.title("Excel Named Table Extractor")

uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])

if uploaded_file:
    wb = openpyxl.load_workbook(uploaded_file, data_only=True)
    ws = wb["Sheet1"]
    table_names = list(ws.tables.keys())

    if table_names:
        selected_table = st.selectbox("Select a table to extract", table_names)
        table = ws.tables[selected_table]
        table_range = table.ref
        data = ws[table_range]
        data = [[cell.value for cell in row] for row in data]
        df = pd.DataFrame(data[1:], columns=data[0])
        st.subheader(f"Extracted {selected_table}")
        st.dataframe(df)
    else:
        st.warning("No named tables found in Sheet1.")
else:
    st.info("Please upload an Excel file to extract tables.")
