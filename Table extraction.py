import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO

st.set_page_config(page_title="Excel Table Editor", layout="wide")
st.title("📊 Excel Named Table Editor")

uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx"])

def combine_rows(df, selected_indices):
    if not selected_indices:
        return df

    selected_rows = df.loc[selected_indices]
    combined_row = {}

    for col in df.columns:
        if pd.api.types.is_numeric_dtype(df[col]):
            combined_row[col] = selected_rows[col].sum()
        else:
            combined_row[col] = " / ".join(selected_rows[col].astype(str))

    df = df.drop(index=selected_indices)
    df = pd.concat([df, pd.DataFrame([combined_row])], ignore_index=True)
    return df

def merge_columns(df, selected_columns, new_column_name):
    if not selected_columns or len(selected_columns) < 2:
        return df

    df[new_column_name] = df[selected_columns].astype(str).agg(" / ".join, axis=1)
    df = df.drop(columns=selected_columns)
    return df

def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='ModifiedTable')
    output.seek(0)
    return output

if uploaded_file:
    wb = openpyxl.load_workbook(uploaded_file, data_only=True)
    ws = wb["Sheet1"]
    table_names = list(ws.tables.keys())

    if table_names:
        selected_table = st.selectbox("Select a named table", table_names)
        table = ws.tables[selected_table]
        table_range = table.ref
        data = ws[table_range]
        data = [[cell.value for cell in row] for row in data]
        df = pd.DataFrame(data[1:], columns=data[0])

        st.subheader("✏️ Edit Table")
        edited_df = st.data_editor(df, num_rows="dynamic", use_container_width=True)

        st.subheader("🔗 Combine Rows")
        selected_rows = st.multiselect("Select rows to combine (by index)", edited_df.index.tolist())
        if st.button("Combine Selected Rows"):
            edited_df = combine_rows(edited_df, selected_rows)
            st.success("Rows combined successfully!")

        st.subheader("🧬 Merge Columns")
        selected_cols = st.multiselect("Select columns to merge", edited_df.columns.tolist(), key="merge_cols")
        new_col_name = st.text_input("New column name", value="MergedColumn")
        if st.button("Merge Selected Columns"):
            edited_df = merge_columns(edited_df, selected_cols, new_col_name)
            st.success(f"Columns merged into '{new_col_name}'")

        st.subheader("📋 Final Table")
        st.dataframe(edited_df, use_container_width=True)

        st.subheader("📥 Download Modified Table")
        excel_data = to_excel(edited_df)
        st.download_button("Download as Excel", data=excel_data, file_name="modified_table.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        st.warning("No named tables found in Sheet1.")
else:
    st.info("Please upload an Excel file to begin.")

