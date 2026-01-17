import streamlit as st
import pandas as pd
import openpyxl
from openpyxl import load_workbook
import msoffcrypto
from io import BytesIO
from win32com.client.gencache import EnsureDispatch
from pathlib import PurePath
from win32com.client import Dispatch

st.title("Python Project: Working on Excel Using Python")

def create_new_excel(name):
    wb = openpyxl.Workbook()
    sheet = wb.active
    sheet.title = "Sheet1"
    wb.save(f"{name}.xlsx")
    st.success(f"Excel file '{name}.xlsx' created successfully!")

def read_data(file_path, password=None):
    if password:
        decrypted = BytesIO()
        with open(file_path, 'rb') as f:
            office_file = msoffcrypto.OfficeFile(f)
            office_file.load_key(password=password)
            office_file.decrypt(decrypted)
        decrypted.seek(0)
        data = pd.read_excel(decrypted, engine='openpyxl')
    else:
        data = pd.read_excel(file_path)
    st.write("Data from the Excel file:")
    st.dataframe(data)

def modify_excel(file_path, address, value):
    work = load_workbook(file_path)
    sheet = work.active
    sheet[address] = value
    work.save(file_path)
    st.success(f"Value '{value}' has been written to cell '{address}'.")

def set_password(file_path, password):
    xl_file = EnsureDispatch("Excel.Application")
    wb = xl_file.Workbooks.Open(file_path)
    xl_file.DisplayAlerts = False
    wb.Visible = False
    wb.SaveAs(file_path, Password=password)
    wb.Close()
    xl_file.Quit()
    st.success("Password has been set successfully.")

def remove_password(file_path, password, new_file):
    excel_app = Dispatch("Excel.Application")
    workbook = excel_app.Workbooks.Open(file_path, False, True, None, password)
    for sheet in workbook.Worksheets:
        if sheet.ProtectContents:
            sheet.Unprotect(password)
    excel_app.DisplayAlerts = False
    workbook.SaveAs(new_file, FileFormat=51, Password="")
    workbook.Close(SaveChanges=True)
    excel_app.Quit()
    st.success("Password removed and file saved without password.")

choice = st.radio("Choose an option:", ("Create New Excel File", "Upload Existing Excel File"))

if choice == "Create New Excel File":
    file_name = st.text_input("Enter the name for the new file:")
    if st.button("Create File"):
        create_new_excel(file_name)

elif choice == "Upload Existing Excel File":
    uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx"])
    password = st.text_input("Enter password if the file is protected (leave blank if not)", type="password")
    if uploaded_file is not None:
        file_path = uploaded_file.name
        with open(file_path, 'wb') as f:
            f.write(uploaded_file.getbuffer())
        read_data(file_path, password)
        
        modify = st.text_input("Enter the cell address you want to modify:")
        if modify:
            value = st.text_input("Enter the new value for the cell:")
            if st.button("Modify Excel"):
                modify_excel(file_path, modify.upper(), value)
        
        if st.button("Remove Password"):
            new_file = st.text_input("Enter name for new file without password:")
            if new_file:
                remove_password(file_path, password, new_file)
        
        new_password = st.text_input("Enter a new password to protect the file:")
        if new_password and st.button("Set Password"):
            set_password(file_path, new_password)
