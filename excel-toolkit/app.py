"""
Excel Manipulation Tool - Professional Edition
Main Application Entry Point

This is the refactored modular version with clean separation of concerns.
All features are organized into separate modules for better maintainability.
"""

import streamlit as st
from src.config.settings import (
    APP_TITLE, APP_ICON, APP_LAYOUT,
    SESSION_UPLOADED_FILE, SESSION_WORKBOOK, SESSION_FILE_PATH, SESSION_DF_DICT
)
from src.ui.tab_basic import render_basic_operations_tab
from src.ui.tab_analysis import render_data_analysis_tab
from src.ui.tab_bulk import render_bulk_operations_tab
from src.ui.tab_sheets import render_sheet_management_tab


# ==================== PAGE CONFIGURATION ====================
st.set_page_config(
    page_title="Excel Manipulation Tool",
    layout=APP_LAYOUT,
    page_icon=APP_ICON
)


# ==================== SESSION STATE INITIALIZATION ====================
def initialize_session_state():
    """Initialize session state variables"""
    if SESSION_UPLOADED_FILE not in st.session_state:
        st.session_state[SESSION_UPLOADED_FILE] = None
    if SESSION_WORKBOOK not in st.session_state:
        st.session_state[SESSION_WORKBOOK] = None
    if SESSION_FILE_PATH not in st.session_state:
        st.session_state[SESSION_FILE_PATH] = None
    if SESSION_DF_DICT not in st.session_state:
        st.session_state[SESSION_DF_DICT] = {}


# ==================== MAIN APPLICATION ====================
def main():
    """Main application function"""
    # Initialize session state
    initialize_session_state()
    
    # App title
    st.title(APP_TITLE)
    st.markdown("---")
    
    # Sidebar
    with st.sidebar:
        st.header("🎯 Navigation")
        st.info("Upload an Excel file to unlock all features")
        
        if st.session_state[SESSION_UPLOADED_FILE]:
            st.success("✅ File loaded successfully")
            if st.session_state[SESSION_DF_DICT]:
                st.metric("Sheets", len(st.session_state[SESSION_DF_DICT]))
    
    # Main tabs
    tab1, tab2, tab3, tab4 = st.tabs([
        "📁 Basic Operations",
        "📊 Data Analysis & Visualization",
        "⚡ Bulk Operations",
        "📑 Sheet Management"
    ])
    
    # Render each tab
    with tab1:
        render_basic_operations_tab()
    
    with tab2:
        render_data_analysis_tab()
    
    with tab3:
        render_bulk_operations_tab()
    
    with tab4:
        render_sheet_management_tab()
    
    # Footer
    st.markdown("---")
    st.markdown("""
    <div style='text-align: center; color: gray;'>
        <p>Excel Manipulation Tool - Professional Edition | Built with Streamlit & Python</p>
        <p>Modular Architecture | 21 Features | Production Ready</p>
    </div>
    """, unsafe_allow_html=True)


# ==================== ENTRY POINT ====================
if __name__ == "__main__":
    main()
