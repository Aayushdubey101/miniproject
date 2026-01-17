"""
Basic Excel Operations
Functions for creating, modifying, and password-protecting Excel files
"""

import streamlit as st
import openpyxl
from openpyxl import load_workbook, Workbook
from io import BytesIO
from win32com.client.gencache import EnsureDispatch
from win32com.client import Dispatch
import pythoncom
import os


def create_new_excel(name):
    """
    Create a new Excel file with one sheet
    
    Args:
        name: Base filename (without extension)
        
    Returns:
        BytesIO object containing the new Excel file
    """
    try:
        wb = Workbook()
        sheet = wb.active
        sheet.title = "Sheet1"
        output = BytesIO()
        wb.save(output)
        output.seek(0)
        
        st.download_button(
            label="📥 Download New Excel File",
            data=output.getvalue(),
            file_name=f"{name}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.success(f"Excel file '{name}.xlsx' created successfully!")
        return output
    except Exception as e:
        st.error(f"Error creating file: {str(e)}")
        return None


def modify_excel_cell(wb, sheet_name, address, value):
    """
    Modify a specific cell in Excel workbook
    
    Args:
        wb: openpyxl Workbook object
        sheet_name: Name of the sheet
        address: Cell address (e.g., 'A1')
        value: New value for the cell
        
    Returns:
        Modified workbook
    """
    try:
        sheet = wb[sheet_name]
        sheet[address] = value
        st.success(f"Value '{value}' has been written to cell '{address}' in sheet '{sheet_name}'.")
        return wb
    except Exception as e:
        st.error(f"Error modifying cell: {str(e)}")
        return wb


def set_password_excel(file_path, password):
    """
    Set password for Excel file (Windows only - requires Excel installed)
    
    Args:
        file_path: Full path to Excel file
        password: Password to set
    """
    try:
        # Initialize COM for this thread
        pythoncom.CoInitialize()
        
        xl_file = EnsureDispatch("Excel.Application")
        wb = xl_file.Workbooks.Open(file_path)
        xl_file.DisplayAlerts = False
        wb.Visible = False
        wb.SaveAs(file_path, Password=password)
        wb.Close()
        xl_file.Quit()
        
        # Uninitialize COM
        pythoncom.CoUninitialize()
        
        st.success("Password has been set successfully.")
    except Exception as e:
        pythoncom.CoUninitialize()
        st.error(f"Error setting password: {str(e)}")


def remove_password_excel(file_path, password):
    """
    Remove password from Excel file (Windows only - requires Excel installed)
    
    Args:
        file_path: Full path to Excel file
        password: Current password
    """
    try:
        # Initialize COM for this thread
        pythoncom.CoInitialize()
        
        excel_app = Dispatch("Excel.Application")
        workbook = excel_app.Workbooks.Open(file_path, False, True, None, password)
        for sheet in workbook.Worksheets:
            if sheet.ProtectContents:
                sheet.Unprotect(password)
        excel_app.DisplayAlerts = False
        
        # Save without password
        output_path = file_path.replace('.xlsx', '_unprotected.xlsx')
        workbook.SaveAs(output_path, FileFormat=51, Password="")
        workbook.Close(SaveChanges=True)
        excel_app.Quit()
        
        # Uninitialize COM
        pythoncom.CoUninitialize()
        
        # Read and provide download
        with open(output_path, 'rb') as f:
            st.download_button(
                label="📥 Download Unprotected File",
                data=f.read(),
                file_name=os.path.basename(output_path),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        st.success("Password removed successfully!")
        
        # Cleanup
        os.remove(output_path)
    except Exception as e:
        pythoncom.CoUninitialize()
        st.error(f"Error removing password: {str(e)}")
