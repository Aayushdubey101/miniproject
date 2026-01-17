"""
Application Configuration and Settings
"""

# Application metadata
APP_TITLE = "📊 Excel Manipulation Tool - Professional Edition"
APP_ICON = "📊"
APP_LAYOUT = "wide"

# File handling settings
MAX_PREVIEW_ROWS = 100
MAX_FILE_SIZE_MB = 100
SUPPORTED_EXTENSIONS = ["xlsx", "xls"]

# Chart settings
DEFAULT_CHART_TEMPLATE = "plotly_white"
CHART_TYPES = ["Bar Chart", "Line Chart", "Pie Chart", "Scatter Plot"]

# Statistics options
NUMERIC_STATS = ['Mean', 'Median', 'Mode', 'Sum', 'Count', 'Min', 'Max', 'Std Dev']
TEXT_STATS = ['Count', 'Unique', 'Mode']

# Pivot table aggregations
PIVOT_AGGREGATIONS = ["sum", "mean", "count", "min", "max"]

# Filter conditions
FILTER_CONDITIONS = ["equals", "contains", "greater than", "less than", "not equals"]
DELETE_CONDITIONS = ["equals", "contains", "greater than", "less than", "empty"]

# Sheet name validation
MAX_SHEET_NAME_LENGTH = 31
INVALID_SHEET_CHARS = ['\\', '/', '*', '?', ':', '[', ']']

# Session state keys
SESSION_UPLOADED_FILE = 'uploaded_file'
SESSION_WORKBOOK = 'workbook'
SESSION_FILE_PATH = 'file_path'
SESSION_DF_DICT = 'df_dict'
