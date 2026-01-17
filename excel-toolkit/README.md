# 📊 Excel Manipulation Tool - Professional Edition

A comprehensive, production-ready Streamlit application for advanced Excel file manipulation, data analysis, visualization, and automation. Built with a modular architecture for scalability and maintainability.

[![Python Version](https://img.shields.io/badge/python-3.7%2B-blue)](https://www.python.org/)
[![Streamlit](https://img.shields.io/badge/streamlit-latest-red)](https://streamlit.io/)
<!-- [![License](https://img.shields.io/badge/license-MIT-green)](LICENSE) -->

## 🌟 Features

### 📁 Basic Operations
- **Create New Excel Files** - Generate blank Excel workbooks with custom names
- **Upload & Read Files** - Support for .xlsx and .xls formats with password protection
- **Cell Modification** - Update individual cell values with validation
- **Password Management** - Set and remove file passwords (Windows only)

### 📈 Data Analysis & Visualization
- **Interactive Charts** - Create bar, line, pie, and scatter plots with Plotly
- **Statistical Analysis** - Calculate mean, median, mode, sum, standard deviation, min, max
- **Pivot Tables** - Dynamic pivot table generation with customizable aggregations
- **Advanced Filtering** - Filter data with multiple conditions (equals, contains, greater than, less than)
- **Smart Search** - Search across all sheets with case-sensitive/insensitive options

### ⚡ Bulk Operations & Automation
- **Batch Modifications** - Upload CSV to modify multiple cells at once
- **File Merging** - Combine multiple Excel files into one workbook
- **Smart Splitting** - Split files based on column values or criteria
- **Data Copy** - Copy data between sheets with range validation
- **Conditional Deletion** - Delete rows/columns based on custom conditions
- **Find & Replace** - Search and replace text across entire workbook with preview

### 📋 Sheet Management
- **CRUD Operations** - Add, delete, and rename sheets with validation
- **Sheet Reordering** - Reorganize sheet order visually
- **Visibility Control** - Hide/unhide sheets as needed
- **Sheet Protection** - Protect/unprotect individual sheets with passwords

## 🏗️ Project Structure
```
excel-toolkit/
├── app.py                          # Main application entry point (~100 lines)
├── pyproject.toml                  # Project dependencies
├── README.md                       # This file
├── uv.lock                         # Dependency lock file
│
└── src/                            # Source code directory
    ├── config/                     # Configuration module
    │   └── settings.py             # App settings and constants
    │
    ├── utils/                      # Utility functions
    │   ├── file_handlers.py        # File loading/saving utilities
    │   └── excel_helpers.py        # Excel-specific helpers
    │
    ├── features/                   # Feature modules
    │   ├── basic_operations.py     # Create, modify, password operations
    │   ├── data_analysis.py        # Charts, statistics, pivot tables
    │   ├── bulk_operations.py      # Batch, merge, split, find/replace
    │   └── sheet_management.py     # Sheet add/delete/rename/protect
    │
    └── ui/                         # UI components and tabs
        ├── components.py           # Reusable UI components
        ├── tab_basic.py            # Basic Operations UI
        ├── tab_analysis.py         # Data Analysis UI
        ├── tab_bulk.py             # Bulk Operations UI
        └── tab_sheets.py           # Sheet Management UI
```

See [project-structure.md](project-structure.md) for detailed documentation.

## 🚀 Installation

### Prerequisites
- Python 3.7 or higher
- [UV](https://github.com/astral-sh/uv) package manager (recommended) or pip

### Using UV (Recommended)
```bash
# Clone the repository
git clone https://github.com/Aayushdubey101/miniproject.git
cd miniproject/excel-toolkit

# Install dependencies
uv sync

# Run the application
uv run streamlit run app.py
```

### Using Pip
```bash
# Clone the repository
git clone https://github.com/Aayushdubey101/miniproject.git
cd miniproject/excel-toolkit

# Create virtual environment
python -m venv .venv
source .venv/bin/activate  # On Windows: .venv\Scripts\activate

# Install dependencies
pip install streamlit pandas openpyxl msoffcrypto-tool pywin32 plotly matplotlib seaborn

# Run the application
streamlit run app.py
```

## 💻 Usage

### Quick Start

1. **Launch the application:**
```bash
   uv run streamlit run app.py
```

2. **Open your browser:**
   Navigate to `http://localhost:8501`

3. **Select a feature category:**
   - Tab 1: Basic Operations
   - Tab 2: Data Analysis & Visualization
   - Tab 3: Bulk Operations
   - Tab 4: Sheet Management

### Example Workflows

#### Creating Charts
1. Go to "Data Analysis & Visualization" tab
2. Upload your Excel file
3. Select chart type (Bar, Line, Pie, Scatter)
4. Choose X and Y columns
5. Click "Generate Chart"
6. Download or embed in Excel

#### Batch Modifications
1. Go to "Bulk Operations" tab
2. Upload your Excel file
3. Upload CSV with modifications (format: Sheet, Cell, Value)
4. Preview changes
5. Apply modifications
6. Download updated file

#### Managing Sheets
1. Go to "Sheet Management" tab
2. Upload your Excel file
3. Add, delete, rename, or reorder sheets
4. Hide/unhide sheets
5. Protect sheets with password
6. Download modified workbook

## 📦 Dependencies

### Core Libraries
- **streamlit** - Web application framework
- **pandas** - Data manipulation and analysis
- **openpyxl** - Excel file operations
- **msoffcrypto-tool** - Password-protected file handling

### Visualization
- **plotly** - Interactive charts
- **matplotlib** - Static plotting
- **seaborn** - Statistical visualizations

### Windows-Specific
- **pywin32** - Excel COM automation (password features)

## ⚙️ Configuration

Edit `src/config/settings.py` to customize:
```python
APP_TITLE = "📊 Excel Manipulation Tool"
MAX_PREVIEW_ROWS = 100
SUPPORTED_EXTENSIONS = ["xlsx", "xls"]
```

## 🖥️ Platform Support

| Feature | Windows | macOS | Linux |
|---------|---------|-------|-------|
| Basic Operations | ✅ | ✅ | ✅ |
| Data Analysis | ✅ | ✅ | ✅ |
| Visualization | ✅ | ✅ | ✅ |
| Bulk Operations | ✅ | ✅ | ✅ |
| Sheet Management | ✅ | ✅ | ✅ |
| Password Set/Remove | ✅ | ❌ | ❌ |

**Note:** Password management features require Windows and Microsoft Excel installed due to `win32com` dependency.

## 🧪 Testing
```bash
# Test file upload
uv run streamlit run app.py

# Navigate to Basic Operations
# Upload a test Excel file
# Verify all features work correctly
```

## 🤝 Contributing

Contributions are welcome! Please follow these steps:

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'feat: add amazing feature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

### Development Guidelines
- Follow PEP 8 style guide
- Add docstrings to all functions
- Keep functions focused and single-purpose
- Test features before submitting PR

<!-- ## 📄 License

This project is open-source and available under the [MIT License](LICENSE). -->

## 🐛 Known Issues

- Password management features only work on Windows with Microsoft Excel installed
- Large files (>100MB) may experience performance degradation
- Some Excel formatting may not be preserved during operations

## 🔮 Roadmap

- [ ] Add support for CSV and Google Sheets
- [ ] Implement data validation rules
- [ ] Add macro support
- [ ] Create REST API endpoints
- [ ] Add unit and integration tests
- [ ] Support for cloud storage (Google Drive, OneDrive)
- [ ] Multi-language support

## 📧 Contact & Support

- **Repository:** [miniproject/excel-toolkit](https://github.com/Aayushdubey101/miniproject)
- **Issues:** [GitHub Issues](https://github.com/Aayushdubey101/miniproject/issues)
- **Author:** Aayush Dubey

## 🙏 Acknowledgments

- Built with [Streamlit](https://streamlit.io/)
- Excel operations powered by [openpyxl](https://openpyxl.readthedocs.io/)
- Data analysis with [Pandas](https://pandas.pydata.org/)
- Visualizations with [Plotly](https://plotly.com/python/)

---

**Made with ❤️ by Aayush Dubey**

**Last Updated:** January 2026  
**Version:** 2.0.0 (Modular Architecture)
