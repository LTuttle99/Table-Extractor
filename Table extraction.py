import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO

st.set_page_config(page_title="Manual Excel Table Editor", layout="wide")
st.title("✂️ Manual Excel Table Editor")

# --- Utility Functions (Cached) ---
@st.cache_resource(ttl=3600)
def load_workbook_from_bytesio(file_buffer):
    """Loads an OpenPyXL workbook from a BytesIO object."""
    file_buffer.seek(0)
    return openpyxl.load_workbook(file_buffer, data_only=True)

@st.cache_data(ttl=3600)
def get_initial_dataframe(_workbook, sheet_name, start_row, end_row, start_col, end_col, use_first_row_as_header):
    """
    Extracts a DataFrame from a specified range within an OpenPyXL worksheet.
    Handles header logic and ensures unique column names.
    """
    ws = _workbook[sheet_name]

    start_row = max(1, start_row)
    end_row = max(start_row, end_row)
    start_col = max(1, start_col)
    end_col = max(start_col, end_col)

    data = []
    try:
        for row in ws.iter_rows(min_row=start_row, max_row=end_row, min_col=start_col, max_col=end_col, values_only=True):
            data.append(list(row))
    except Exception as e:
        st.error(f"Error reading specified range from sheet: {e}. Please check your row/column inputs.")
        return pd.DataFrame()

    if not data:
        return pd.DataFrame()

    if use_first_row_as_header and len(data) > 0:
        raw_headers = list(data[0])
        rows = data[1:]
    else:
        raw_headers = [f"Column_{i+1}" for i in range(end_col - start_col + 1)]
        rows = data

    headers = []
    seen = {}
    for h in raw_headers:
        h_str = str(h) if h is not None and str(h).strip() != "" else "Unnamed"
        if h_str in seen:
            seen[h_str] += 1
            h_str = f"{h_str}_{seen[h_str]}"
        else:
            seen[h_str] = 0
        headers.append(h_str)
    
    adjusted_rows = []
    expected_cols = len(headers)
    for row in rows:
        if len(row) < expected_cols:
            adjusted_rows.append(list(row) + [None] * (expected_cols - len(row)))
        else:
            adjusted_rows.append(list(row[:expected_cols]))
            
    df_result = pd.DataFrame(adjusted_rows, columns=headers)
    df_result = df_result.dropna(how="all")

    if 'Order' not in df_result.columns:
        df_result.insert(0, 'Order', range(1, len(df_result) + 1))

    return df_result

# --- Main App Logic ---
with st.sidebar:
    st.header("Upload Excel File")
    uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx", "xls"])
    st.markdown("---")

if uploaded_file is not None:
    try:
        with st.spinner("Loading Excel file..."):
            wb = load_workbook_from_bytesio(uploaded_file)
        
        sheet_names = wb.sheetnames
        st.success("File loaded successfully!")

        with st.sidebar:
            st.header("Sheet Selection")
            selected_sheet = st.selectbox(
                "Select the relevant sheet", 
                sheet_names, 
                key="selected_sheet_sidebar_manual"
            )
            ws = wb[selected_sheet]
            max_row = ws.max_row
            max_column = ws.max_column
            st.info(f"Sheet dimensions: {max_row} rows, {max_column} columns")
            st.markdown("---")

        st.markdown("### 🔍 Define Table Range")

        col1, col2 = st.columns(2)
        with col1:
            start_row = st.number_input("Start Row", min_value=1, value=1, max_value=max_row, key="manual_start_row")
            end_row = st.number_input("End Row", min_value=start_row, value=max_row, max_value=max_row, key="manual_end_row")
        with col2:
            start_col = st.number_input("Start Column (e.g., 1 for A, 2 for B)", min_value=1, value=1, max_value=max_column, key="manual_start_col")
            end_col = st.number_input("End Column (e.g., 1 for A, 2 for B)", min_value=start_col, value=max_column, max_value=max_column, key="manual_end_col")
        
        use_first_row_as_header = st.checkbox("Use first row as header", value=True, key="manual_use_header")

        # Fetch initial DataFrame based on manual inputs
        df_initial = get_initial_dataframe(wb, selected_sheet,
                                             start_row, end_row,
                                             start_col, end_col,
                                             use_first_row_as_header)

        # Session State Management for current_df and history
        current_data_selection_id = (
            f"{uploaded_file.file_id}-"
            f"{selected_sheet}-"
            f"{start_row}-"
            f"{end_row}-"
            f"{start_col}-"
            f"{end_col}-"
            f"{use_first_row_as_header}"
        )

        # Initialize or reset the DataFrame and history based on selection ID
        if "last_processed_file_id_manual" not in st.session_state or st.session_state.last_processed_file_id_manual != current_data_selection_id:
            st.session_state.current_df_manual = df_initial.copy()
            st.session_state.history_manual = [] # Clear history on new selection
            st.session_state.last_processed_file_id_manual = current_data_selection_id
            if not df_initial.empty: # Only show info if actual data is loaded
                st.info("Table initialized from specified range.")
        elif st.session_state.current_df_manual.empty and not df_initial.empty:
            # Re-initialize if the previous data was empty but new detection isn't
            st.session_state.current_df_manual = df_initial.copy()
            st.session_state.history_manual = []
            st.session_state.last_processed_file_id_manual = current_data_selection_id
            st.info("Re-initializing table from file as previous data was empty.")

        st.markdown("---")

        # --- Display and Editing UI ---
        if not st.session_state.current_df_manual.empty:
            st.subheader("✏️ Review and Edit Table (Directly in Table)")
            st.info("You can directly edit cells in the table. To reorder rows, edit the numbers in the 'Order' column. To delete a row, click the 'X' button on the right of the row.")

            # Ensure 'Order' column is numeric for proper sorting and data editor
            st.session_state.current_df_manual['Order'] = pd.to_numeric(st.session_state.current_df_manual['Order'], errors='coerce').fillna(0).astype(int)

            edited_df_manual = st.data_editor(
                st.session_state.current_df_manual,
                num_rows="dynamic", # Allows adding/deleting rows directly in the editor
                use_container_width=True,
                column_config={
                    "Order": st.column_config.NumberColumn(
                        "Order",
                        help="Assign a number to reorder rows.",
                        default=0,
                        step=1,
                        format="%d"
                    )
                },
                key="main_data_editor_manual" # Unique key for the data editor
            )

            # Check if edited_df is different from current_df
            # This triggers a history save and success message
            if not edited_df_manual.equals(st.session_state.current_df_manual):
                st.session_state.history_manual.append(st.session_state.current_df_manual.copy())
                st.session_state.current_df_manual = edited_df_manual.copy()
                st.success("Changes detected. Apply order or continue editing.")
                st.rerun() # Rerun to reflect changes immediately and prevent edit conflicts

            if st.button("Apply Manual Row Order Changes", key="apply_order_manual"):
                if 'Order' in st.session_state.current_df_manual.columns:
                    temp_df = st.session_state.current_df_manual.copy()
                    temp_df['Order_temp_sort'] = temp_df['Order']
                    if temp_df['Order_temp_sort'].duplicated().any():
                        temp_df['Order_temp_sort'] = temp_df['Order'].astype(str) + '.' + temp_df.groupby('Order_temp_sort').cumcount().astype(str)
                        temp_df['Order_temp_sort'] = pd.to_numeric(temp_df['Order_temp_sort'], errors='coerce')
                    st.session_state.current_df_manual = temp_df.sort_values(by='Order_temp_sort', ascending=True).drop(columns=['Order_temp_sort']).reset_index(drop=True)
                    st.success("Rows reordered successfully!")
                    st.rerun()
                else:
                    st.warning("No 'Order' column found to reorder rows.")
            
            # Undo functionality (optional, but good for a manual editor)
            if len(st.session_state.history_manual) > 0:
                if st.button("Undo Last Action", key="undo_manual"):
                    st.session_state.current_df_manual = st.session_state.history_manual.pop()
                    st.warning("Last action undone.")
                    st.rerun()
            else:
                st.info("No actions to undo.")

            st.subheader("📋 Final Edited Table")
            # Display the final, non-all-NA rows of the table
            final_df_manual = st.session_state.current_df_manual.dropna(how="all").reset_index(drop=True)
            st.dataframe(final_df_manual, use_container_width=True)

            st.subheader("📥 Download Modified Table")
            def to_excel_manual(df_to_save):
                """Converts a DataFrame to an Excel file in BytesIO object."""
                output = BytesIO()
                if not df_to_save.empty:
                    with pd.ExcelWriter(output, engine="openpyxl") as writer:
                        df_to_save.to_excel(writer, index=False, sheet_name="EditedTable")
                output.seek(0)
                return output

            # Generate Excel data for download
            excel_data_manual = to_excel_manual(final_df_manual)
            st.download_button(
                "Download Edited Table as Excel",
                data=excel_data_manual,
                file_name="edited_excel_table.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="download_manual"
            )

        else:
            st.info("Please define the table range using the inputs above to load data.")

    except Exception as e:
        st.error(f"An unexpected error occurred while processing the Excel file: {e}")
        st.exception(e) # Display full traceback for debugging
        st.info("Please ensure it's a valid Excel file with readable content and try again.")
else:
    st.info("Please upload your Excel file to begin editing.")
