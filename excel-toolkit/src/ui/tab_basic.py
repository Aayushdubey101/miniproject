"""
Tab 1: Basic Operations UI
Create, upload, modify, and password operations
"""

import streamlit as st
import os
import tempfile
from openpyxl import load_workbook
from src.features.basic_operations import (
    create_new_excel, modify_excel_cell, 
    set_password_excel, remove_password_excel
)
from src.utils.file_handlers import (
    load_excel_with_password, get_all_sheets, 
    load_sheet_data, create_download_link
)
from src.ui.components import show_dataframe_preview
from src.config.settings import SESSION_UPLOADED_FILE, SESSION_WORKBOOK, SESSION_FILE_PATH, SESSION_DF_DICT


def render_basic_operations_tab():
    """Render the Basic Operations tab"""
    st.header("📁 Basic Operations")
    
    col1, col2 = st.columns(2)
    
    # Create New File Section
    with col1:
        st.subheader("Create New Excel File")
        file_name = st.text_input("Enter file name:", key="new_file_name")
        if st.button("Create File", key="create_btn"):
            if file_name:
                create_new_excel(file_name)
            else:
                st.warning("Please enter a file name")
    
    # Upload File Section
    with col2:
        st.subheader("Upload Excel File")
        uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx", "xls"], key="file_uploader")
        password = st.text_input("Password (if protected):", type="password", key="file_password")
    
    # Process uploaded file
    if uploaded_file is not None:
        st.session_state[SESSION_UPLOADED_FILE] = uploaded_file
        
        file_bytes = uploaded_file.getvalue()
        file_io = load_excel_with_password(file_bytes, password if password else None)
        
        if file_io:
            # Save to temp file
            temp_path = os.path.join(tempfile.gettempdir(), uploaded_file.name)
            with open(temp_path, 'wb') as f:
                f.write(file_bytes)
            st.session_state[SESSION_FILE_PATH] = temp_path
            
            # Load workbook
            try:
                file_io.seek(0)
                wb = load_workbook(file_io)
                st.session_state[SESSION_WORKBOOK] = wb
                
                # Load all sheets
                file_io.seek(0)
                sheets = get_all_sheets(file_io)
                st.session_state[SESSION_DF_DICT] = {}
                
                for sheet in sheets:
                    file_io.seek(0)
                    df = load_sheet_data(file_io, sheet)
                    if df is not None:
                        st.session_state[SESSION_DF_DICT][sheet] = df
                
                st.success(f"✅ Loaded {len(sheets)} sheet(s)")
                
                # Preview Data
                st.subheader("📋 Preview Data")
                selected_sheet = st.selectbox("Select sheet to view:", sheets, key="preview_sheet")
                
                if selected_sheet in st.session_state[SESSION_DF_DICT]:
                    df = st.session_state[SESSION_DF_DICT][selected_sheet]
                    show_dataframe_preview(df)
                
                # Modify Cell
                st.subheader("✏️ Modify Cell")
                mod_col1, mod_col2, mod_col3 = st.columns(3)
                
                with mod_col1:
                    mod_sheet = st.selectbox("Sheet:", sheets, key="mod_sheet")
                with mod_col2:
                    cell_address = st.text_input("Cell address (e.g., A1):", key="cell_addr")
                with mod_col3:
                    new_value = st.text_input("New value:", key="new_val")
                
                if st.button("Modify Cell", key="modify_cell_btn"):
                    if cell_address and new_value:
                        wb = modify_excel_cell(st.session_state[SESSION_WORKBOOK], mod_sheet, cell_address.upper(), new_value)
                        st.session_state[SESSION_WORKBOOK] = wb
                        
                        st.download_button(
                            label="📥 Download Modified File",
                            data=create_download_link(wb, uploaded_file.name),
                            file_name=f"modified_{uploaded_file.name}",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                
                # Password Operations
                st.subheader("🔒 Password Operations")
                pw_col1, pw_col2 = st.columns(2)
                
                with pw_col1:
                    st.write("**Set Password**")
                    new_password = st.text_input("New password:", type="password", key="set_pw")
                    if st.button("Set Password", key="set_pw_btn"):
                        if new_password and st.session_state.get(SESSION_FILE_PATH):
                            set_password_excel(st.session_state[SESSION_FILE_PATH], new_password)
                
                with pw_col2:
                    st.write("**Remove Password**")
                    remove_pw = st.text_input("Current password:", type="password", key="remove_pw")
                    if st.button("Remove Password", key="remove_pw_btn"):
                        if remove_pw and st.session_state.get(SESSION_FILE_PATH):
                            remove_password_excel(st.session_state[SESSION_FILE_PATH], remove_pw)
                
            except Exception as e:
                st.error(f"Error loading workbook: {str(e)}")
