"""
Excel-Specific Helper Functions
Utilities for Excel operations and validation
"""

from copy import copy
from src.config.settings import MAX_SHEET_NAME_LENGTH, INVALID_SHEET_CHARS


def validate_sheet_name(name, existing_sheets):
    """
    Validate sheet name according to Excel rules
    
    Args:
        name: Proposed sheet name
        existing_sheets: List of existing sheet names
        
    Returns:
        Tuple of (is_valid: bool, message: str)
    """
    if not name or name.strip() == "":
        return False, "Sheet name cannot be empty"
    
    if name in existing_sheets:
        return False, "Sheet name already exists"
    
    if len(name) > MAX_SHEET_NAME_LENGTH:
        return False, f"Sheet name must be {MAX_SHEET_NAME_LENGTH} characters or less"
    
    if any(char in name for char in INVALID_SHEET_CHARS):
        return False, f"Sheet name cannot contain: {', '.join(INVALID_SHEET_CHARS)}"
    
    return True, "Valid"


def copy_cell_style(source_cell, target_cell):
    """
    Copy formatting from source cell to target cell
    
    Args:
        source_cell: Source openpyxl cell
        target_cell: Target openpyxl cell
    """
    if source_cell.has_style:
        target_cell.font = copy(source_cell.font)
        target_cell.border = copy(source_cell.border)
        target_cell.fill = copy(source_cell.fill)
        target_cell.number_format = copy(source_cell.number_format)
        target_cell.protection = copy(source_cell.protection)
        target_cell.alignment = copy(source_cell.alignment)
