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

    # Ensure 'Order' column exists for manual reordering
    if 'Order' not in df_result.columns:
        df_result.insert(0, 'Order', range(1, len(df_result) + 1))
    else:
        # If 'Order' exists, ensure it's numeric and reset if needed
        df_result['Order'] = pd.to_numeric(df_result['Order'], errors='coerce').fillna(0).astype(int)
        df_result = df_result.sort_values(by='Order').reset_index(drop=True)
        df_result['Order'] = range(1, len(df_result) + 1) # Re-index for consistent order

    return df_result

# --- Callback function for data_editor changes ---
def on_data_editor_change():
    # This callback updates a dedicated session state variable for selected rows
    # and forces a rerun so the UI can react to the selection immediately.
    if 'main_data_editor_manual' in st.session_state:
        st.session_state.selected_rows_for_combine = st.session_state.main_data_editor_manual.get('selected_rows', [])
    else:
        st.session_state.selected_rows_for_combine = []
    st.rerun()

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
            st.session_state.selected_rows_for_combine = [] # Initialize selected rows
            if not df_initial.empty: # Only show info if actual data is loaded
                st.info("Table initialized from specified range.")
        elif st.session_state.current_df_manual.empty and not df_initial.empty:
            # Re-initialize if the previous data was empty but new detection isn't
            st.session_state.current_df_manual = df_initial.copy()
            st.session_state.history_manual = []
            st.session_state.last_processed_file_id_manual = current_data_selection_id
            st.session_state.selected_rows_for_combine = [] # Initialize selected rows
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
                key="main_data_editor_manual", # Unique key for the data editor
                on_change=on_data_editor_change # ADD THIS CALLBACK
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
                    # Handle duplicate order numbers by making them unique for sorting purposes
                    if temp_df['Order_temp_sort'].duplicated().any():
                        temp_df['Order_temp_sort'] = temp_df['Order'].astype(str) + '.' + temp_df.groupby('Order_temp_sort').cumcount().astype(str)
                        temp_df['Order_temp_sort'] = pd.to_numeric(temp_df['Order_temp_sort'], errors='coerce') # Convert back to numeric for sorting
                    
                    st.session_state.current_df_manual = temp_df.sort_values(by='Order_temp_sort', ascending=True).drop(columns=['Order_temp_sort']).reset_index(drop=True)
                    # Re-assign sequential Order numbers after sorting
                    st.session_state.current_df_manual['Order'] = range(1, len(st.session_state.current_df_manual) + 1)
                    st.success("Rows reordered successfully!")
                    st.rerun()
                else:
                    st.warning("No 'Order' column found to reorder rows.")
            
            # --- Manual Row Combination Feature ---
            st.markdown("---")
            st.subheader("🔗 Combine Selected Rows Manually")
            st.info("Select rows from the table above using the checkboxes on the left, then click 'Combine Selected Rows'. Numeric columns will be summed, text columns joined by ' / '.")

            # GET SELECTED ROWS FROM THE DEDICATED SESSION STATE VARIABLE, UPDATED BY CALLBACK
            selected_rows_indices = st.session_state.get('selected_rows_for_combine', [])
            
            # --- Always show the combine button, but disable it if not enough rows are selected ---
            new_row_name = st.text_input("Enter a name for the combined row (e.g., 'Combined Item')", "Combined Row", key="combined_row_name_manual")
            
            combine_button_disabled = len(selected_rows_indices) < 2
            
            if st.button("Combine Selected Rows", key="combine_selected_manual", disabled=combine_button_disabled):
                # Ensure we have enough rows before proceeding (should be true if button is not disabled)
                if len(selected_rows_indices) >= 2:
                    st.session_state.history_manual.append(st.session_state.current_df_manual.copy()) # Save current state

                    combined_row_data = {}
                    # Use .loc with the directly obtained indices from selected_rows_indices
                    selected_df_for_combine = st.session_state.current_df_manual.loc[selected_rows_indices]

                    for col_idx, col in enumerate(st.session_state.current_df_manual.columns):
                        if col == 'Order': # Special handling for 'Order' column
                            # Assign a new high order number, ensuring it's unique
                            combined_row_data[col] = st.session_state.current_df_manual['Order'].max() + 1 if not st.session_state.current_df_manual.empty else 1
                        elif pd.api.types.is_numeric_dtype(st.session_state.current_df_manual[col]):
                            combined_row_data[col] = selected_df_for_combine[col].sum()
                        else:
                            # Join non-numeric values, handling NaNs
                            joined_value = " / ".join(selected_df_for_combine[col].dropna().astype(str).tolist())
                            combined_row_data[col] = joined_value
                    
                    # Set the new name for the first non-order column, or the first column if 'Order' isn't present
                    if not st.session_state.current_df_manual.empty and not st.session_state.current_df_manual.columns.empty:
                        if 'Order' in st.session_state.current_df_manual.columns and len(st.session_state.current_df_manual.columns) > 1:
                            first_non_order_col = next((c for c in st.session_state.current_df_manual.columns if c != 'Order'), None)
                            if first_non_order_col:
                                combined_row_data[first_non_order_col] = new_row_name
                        else: # If 'Order' is the only column, or it's not present and we need to assign a name
                            combined_row_data[st.session_state.current_df_manual.columns[0]] = new_row_name


                    combined_df_new = pd.DataFrame([combined_row_data], columns=st.session_state.current_df_manual.columns)
                    
                    remaining_df = st.session_state.current_df_manual.drop(index=selected_rows_indices).reset_index(drop=True)
                    st.session_state.current_df_manual = pd.concat([remaining_df, combined_df_new], ignore_index=True)
                    
                    # After combining, re-assign 'Order' numbers to ensure they are sequential
                    if 'Order' in st.session_state.current_df_manual.columns:
                        st.session_state.current_df_manual['Order'] = range(1, len(st.session_state.current_df_manual) + 1)

                    st.success(f"Selected rows combined into '{new_row_name}'.")
                    # Clear the selected rows state after combining
                    st.session_state.selected_rows_for_combine = [] 
                    st.rerun()
                else: # This case should ideally not be hit if the button is disabled correctly
                    st.warning("Please select at least two rows to combine.")
            
            # Provide feedback when the button is disabled
            if combine_button_disabled:
                st.warning("Please select at least two rows to enable the 'Combine Selected Rows' button.")


            # --- Undo functionality ---
            st.markdown("---")
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
