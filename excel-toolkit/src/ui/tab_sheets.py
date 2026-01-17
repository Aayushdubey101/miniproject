"""
Tab 4: Sheet Management UI
Add, delete, rename, reorder, hide, and protect sheets
"""

import streamlit as st
from src.features.sheet_management import (
    add_sheet, delete_sheet, rename_sheet, reorder_sheets,
    hide_unhide_sheet, protect_sheet, unprotect_sheet
)
from src.utils.file_handlers import create_download_link
from src.utils.excel_helpers import validate_sheet_name
from src.config.settings import SESSION_WORKBOOK


def render_sheet_management_tab():
    """Render the Sheet Management tab"""
    st.header("📑 Sheet Management")
    
    if not st.session_state.get(SESSION_WORKBOOK):
        st.warning("⚠️ Please upload an Excel file in the Basic Operations tab first")
        return
    
    wb = st.session_state[SESSION_WORKBOOK]
    
    # Add/Delete/Rename Sheets
    with st.expander("➕ Add / ✏️ Rename / 🗑️ Delete Sheets", expanded=True):
        sheet_col1, sheet_col2, sheet_col3 = st.columns(3)
        
        with sheet_col1:
            st.write("**Add Sheet**")
            new_sheet_name = st.text_input("New sheet name:", key="new_sheet_name")
            position = st.radio("Position:", ["End", "Beginning"], key="sheet_position", horizontal=True)
            
            if st.button("Add Sheet", key="add_sheet_btn"):
                if new_sheet_name:
                    valid, msg = validate_sheet_name(new_sheet_name, wb.sheetnames)
                    if valid:
                        wb = add_sheet(wb, new_sheet_name, position.lower())
                        st.session_state[SESSION_WORKBOOK] = wb
                        st.success(f"✅ Added sheet '{new_sheet_name}'")
                        st.rerun()
                    else:
                        st.error(msg)
                else:
                    st.warning("Please enter a sheet name")
        
        with sheet_col2:
            st.write("**Rename Sheet**")
            old_sheet_name = st.selectbox("Select sheet:", wb.sheetnames, key="rename_old")
            rename_new_name = st.text_input("New name:", key="rename_new")
            
            if st.button("Rename Sheet", key="rename_sheet_btn"):
                if rename_new_name:
                    valid, msg = validate_sheet_name(rename_new_name, [s for s in wb.sheetnames if s != old_sheet_name])
                    if valid:
                        wb = rename_sheet(wb, old_sheet_name, rename_new_name)
                        st.session_state[SESSION_WORKBOOK] = wb
                        st.success(f"✅ Renamed to '{rename_new_name}'")
                        st.rerun()
                    else:
                        st.error(msg)
                else:
                    st.warning("Please enter a new name")
        
        with sheet_col3:
            st.write("**Delete Sheet**")
            delete_sheet_name = st.selectbox("Select sheet:", wb.sheetnames, key="delete_sheet")
            
            if st.button("Delete Sheet", key="delete_sheet_btn"):
                if len(wb.sheetnames) > 1:
                    wb = delete_sheet(wb, delete_sheet_name)
                    st.session_state[SESSION_WORKBOOK] = wb
                    st.success(f"✅ Deleted sheet '{delete_sheet_name}'")
                    st.rerun()
                else:
                    st.error("Cannot delete the last sheet")
        
        if st.button("💾 Save All Sheet Changes", key="save_sheet_changes"):
            st.download_button(
                label="📥 Download Updated File",
                data=create_download_link(wb, "updated.xlsx"),
                file_name="sheets_updated.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    
    # Reorder Sheets
    with st.expander("🔀 Reorder Sheets"):
        st.write("**Current order:**")
        current_order = wb.sheetnames
        for i, sheet in enumerate(current_order):
            st.write(f"{i+1}. {sheet}")
        
        st.write("**New order (comma-separated):**")
        st.info("Enter sheet names in desired order, separated by commas")
        new_order_input = st.text_input("New order:", value=", ".join(current_order), key="new_order")
        
        if st.button("Apply New Order", key="reorder_btn"):
            new_order = [s.strip() for s in new_order_input.split(",")]
            if set(new_order) == set(current_order):
                wb = reorder_sheets(wb, new_order)
                st.session_state[SESSION_WORKBOOK] = wb
                st.success("✅ Sheets reordered successfully")
                st.download_button(
                    label="📥 Download Reordered File",
                    data=create_download_link(wb, "reordered.xlsx"),
                    file_name="sheets_reordered.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.error("Invalid order. Please include all sheets exactly once.")
    
    # Hide/Unhide Sheets
    with st.expander("👁️ Hide / Unhide Sheets"):
        st.write("**Sheet Visibility Status:**")
        for sheet in wb.worksheets:
            col1, col2, col3 = st.columns([3, 1, 1])
            with col1:
                st.write(f"**{sheet.title}**")
            with col2:
                status = "Visible" if sheet.sheet_state == 'visible' else "Hidden"
                st.write(status)
            with col3:
                if sheet.sheet_state == 'visible':
                    if st.button("Hide", key=f"hide_{sheet.title}"):
                        wb = hide_unhide_sheet(wb, sheet.title, hide=True)
                        st.session_state[SESSION_WORKBOOK] = wb
                        st.rerun()
                else:
                    if st.button("Unhide", key=f"unhide_{sheet.title}"):
                        wb = hide_unhide_sheet(wb, sheet.title, hide=False)
                        st.session_state[SESSION_WORKBOOK] = wb
                        st.rerun()
        
        if st.button("💾 Save Visibility Changes", key="save_visibility"):
            st.download_button(
                label="📥 Download Updated File",
                data=create_download_link(wb, "visibility.xlsx"),
                file_name="visibility_updated.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    
    # Protect/Unprotect Sheets
    with st.expander("🔒 Protect / Unprotect Sheets"):
        st.write("**Sheet Protection Status:**")
        for sheet in wb.worksheets:
            st.write(f"**{sheet.title}**")
            prot_col1, prot_col2 = st.columns(2)
            
            with prot_col1:
                is_protected = sheet.protection.sheet
                st.write(f"Status: {'🔒 Protected' if is_protected else '🔓 Unprotected'}")
            
            with prot_col2:
                if not is_protected:
                    protect_pw = st.text_input(f"Password for {sheet.title}:", type="password", key=f"protect_pw_{sheet.title}")
                    if st.button(f"Protect", key=f"protect_{sheet.title}"):
                        wb = protect_sheet(wb, sheet.title, protect_pw if protect_pw else None)
                        st.session_state[SESSION_WORKBOOK] = wb
                        st.success(f"✅ Protected '{sheet.title}'")
                        st.rerun()
                else:
                    if st.button(f"Unprotect", key=f"unprotect_{sheet.title}"):
                        wb = unprotect_sheet(wb, sheet.title)
                        st.session_state[SESSION_WORKBOOK] = wb
                        st.success(f"✅ Unprotected '{sheet.title}'")
                        st.rerun()
            st.markdown("---")
        
        if st.button("💾 Save Protection Changes", key="save_protection"):
            st.download_button(
                label="📥 Download Updated File",
                data=create_download_link(wb, "protection.xlsx"),
                file_name="protection_updated.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
