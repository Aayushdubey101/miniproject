"""
Reusable UI Components
Common UI elements used across the application
"""

import streamlit as st
from src.config.settings import MAX_PREVIEW_ROWS


def render_file_uploader(label="Choose an Excel file", key="file_uploader"):
    """Render file uploader component"""
    return st.file_uploader(label, type=["xlsx", "xls"], key=key)


def render_sheet_selector(sheets, label="Select sheet:", key="sheet_selector"):
    """Render sheet selection dropdown"""
    return st.selectbox(label, sheets, key=key)


def render_download_button(data, filename, label="📥 Download File"):
    """Render download button with consistent styling"""
    return st.download_button(
        label=label,
        data=data,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


def show_dataframe_preview(df, max_rows=MAX_PREVIEW_ROWS):
    """Display DataFrame with pagination info"""
    st.dataframe(df.head(max_rows), use_container_width=True)
    if len(df) > max_rows:
        st.info(f"Showing first {max_rows} rows of {len(df)} total rows")
