"""
Sheet Management Operations
Functions for adding, deleting, renaming, reordering, hiding, and protecting sheets
"""

import streamlit as st


def add_sheet(wb, sheet_name, position='end'):
    """
    Add new sheet to workbook
    
    Args:
        wb: openpyxl Workbook object
        sheet_name: Name for new sheet
        position: Position to add ('end' or 'beginning')
        
    Returns:
        Modified workbook
    """
    try:
        if position == 'end':
            wb.create_sheet(title=sheet_name)
        elif position == 'beginning':
            wb.create_sheet(title=sheet_name, index=0)
        return wb
    except Exception as e:
        st.error(f"Error adding sheet: {str(e)}")
        return wb


def delete_sheet(wb, sheet_name):
    """
    Delete sheet from workbook
    
    Args:
        wb: openpyxl Workbook object
        sheet_name: Name of sheet to delete
        
    Returns:
        Modified workbook
    """
    try:
        if len(wb.sheetnames) > 1:
            del wb[sheet_name]
        else:
            st.error("Cannot delete the last sheet in the workbook")
        return wb
    except Exception as e:
        st.error(f"Error deleting sheet: {str(e)}")
        return wb


def rename_sheet(wb, old_name, new_name):
    """
    Rename sheet in workbook
    
    Args:
        wb: openpyxl Workbook object
        old_name: Current sheet name
        new_name: New sheet name
        
    Returns:
        Modified workbook
    """
    try:
        sheet = wb[old_name]
        sheet.title = new_name
        return wb
    except Exception as e:
        st.error(f"Error renaming sheet: {str(e)}")
        return wb


def reorder_sheets(wb, new_order):
    """
    Reorder sheets in workbook
    
    Args:
        wb: openpyxl Workbook object
        new_order: List of sheet names in desired order
        
    Returns:
        Modified workbook
    """
    try:
        wb._sheets = [wb[name] for name in new_order]
        return wb
    except Exception as e:
        st.error(f"Error reordering sheets: {str(e)}")
        return wb


def hide_unhide_sheet(wb, sheet_name, hide=True):
    """
    Hide or unhide a sheet
    
    Args:
        wb: openpyxl Workbook object
        sheet_name: Name of sheet
        hide: True to hide, False to unhide
        
    Returns:
        Modified workbook
    """
    try:
        sheet = wb[sheet_name]
        if hide:
            # Check if it's the last visible sheet
            visible_count = sum(1 for s in wb.worksheets if s.sheet_state == 'visible')
            if visible_count <= 1:
                st.error("Cannot hide the last visible sheet")
                return wb
            sheet.sheet_state = 'hidden'
        else:
            sheet.sheet_state = 'visible'
        return wb
    except Exception as e:
        st.error(f"Error hiding/unhiding sheet: {str(e)}")
        return wb


def protect_sheet(wb, sheet_name, password=None):
    """
    Protect a sheet with optional password
    
    Args:
        wb: openpyxl Workbook object
        sheet_name: Name of sheet
        password: Optional password string
        
    Returns:
        Modified workbook
    """
    try:
        sheet = wb[sheet_name]
        sheet.protection.sheet = True
        if password:
            sheet.protection.password = password
        return wb
    except Exception as e:
        st.error(f"Error protecting sheet: {str(e)}")
        return wb


def unprotect_sheet(wb, sheet_name):
    """
    Unprotect a sheet
    
    Args:
        wb: openpyxl Workbook object
        sheet_name: Name of sheet
        
    Returns:
        Modified workbook
    """
    try:
        sheet = wb[sheet_name]
        sheet.protection.sheet = False
        sheet.protection.password = None
        return wb
    except Exception as e:
        st.error(f"Error unprotecting sheet: {str(e)}")
        return wb
