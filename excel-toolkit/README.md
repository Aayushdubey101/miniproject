# Python Excel Manipulation Project

## Description
This is a Streamlit-based web application that provides a user-friendly interface for performing various Excel file operations using Python. The app allows users to create new Excel files, upload and read existing files (with optional password protection), modify cell values, set passwords, and remove passwords from Excel files.

## Features
- **Create New Excel Files**: Generate new Excel files with a default sheet.
- **Upload and Read Excel Files**: Upload existing Excel files and display their data in a table format.
- **Password Support**: Handle password-protected Excel files for reading and manipulation.
- **Modify Cell Values**: Update specific cell values in an Excel file.
- **Password Management**: Set new passwords or remove existing passwords from Excel files.
- **User-Friendly Interface**: Built with Streamlit for an intuitive web-based experience.

## Installation
1. Ensure you have Python installed on your system (version 3.7 or higher recommended).
2. Install the required dependencies using pip:
   ```
   pip install streamlit pandas openpyxl msoffcrypto win32com
   ```
   Note: The `win32com` library is Windows-specific and may not work on other operating systems.

3. Clone or download this project to your local machine.

## Usage
1. Navigate to the project directory in your terminal.
2. Run the Streamlit app:
   ```
   streamlit run app.py
   ```
3. Open your web browser and go to the URL displayed in the terminal (usually `http://localhost:8501`).

4. Use the radio buttons to choose between creating a new Excel file or uploading an existing one.
   - **Create New Excel File**: Enter a name for the new file and click "Create File".
   - **Upload Existing Excel File**: Upload an Excel file, enter a password if required, and perform various operations like reading data, modifying cells, setting/removing passwords.

## Requirements
- Python 3.7+
- Streamlit
- Pandas
- OpenPyXL
- MS Office Crypto (msoffcrypto)
- pywin32 (win32com) - Windows only
- Microsoft Excel (for password operations on Windows)

## Contributing
Contributions are welcome! Please feel free to submit a Pull Request.

## License
This project is open-source and available under the MIT License.
