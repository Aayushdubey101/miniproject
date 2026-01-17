"""
Data Analysis and Visualization
Functions for charts, statistics, pivot tables, filtering, and search
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import re


def create_chart(df, chart_type, x_col, y_col, title="Chart"):
    """
    Create interactive charts using Plotly
    
    Args:
        df: pandas DataFrame
        chart_type: Type of chart (Bar Chart, Line Chart, Pie Chart, Scatter Plot)
        x_col: Column name for X-axis
        y_col: Column name for Y-axis
        title: Chart title
        
    Returns:
        Plotly figure object or None on error
    """
    try:
        if chart_type == "Bar Chart":
            fig = px.bar(df, x=x_col, y=y_col, title=title)
        elif chart_type == "Line Chart":
            fig = px.line(df, x=x_col, y=y_col, title=title)
        elif chart_type == "Pie Chart":
            fig = px.pie(df, names=x_col, values=y_col, title=title)
        elif chart_type == "Scatter Plot":
            fig = px.scatter(df, x=x_col, y=y_col, title=title)
        else:
            return None
        
        fig.update_layout(template="plotly_white")
        return fig
    except Exception as e:
        st.error(f"Error creating chart: {str(e)}")
        return None


def calculate_statistics(df, columns):
    """
    Calculate statistics for selected columns
    
    Args:
        df: pandas DataFrame
        columns: List of column names
        
    Returns:
        DataFrame with statistics or None on error
    """
    try:
        stats_dict = {}
        for col in columns:
            if pd.api.types.is_numeric_dtype(df[col]):
                stats_dict[col] = {
                    'Mean': df[col].mean(),
                    'Median': df[col].median(),
                    'Mode': df[col].mode()[0] if not df[col].mode().empty else None,
                    'Sum': df[col].sum(),
                    'Count': df[col].count(),
                    'Min': df[col].min(),
                    'Max': df[col].max(),
                    'Std Dev': df[col].std()
                }
            else:
                stats_dict[col] = {
                    'Count': df[col].count(),
                    'Unique': df[col].nunique(),
                    'Mode': df[col].mode()[0] if not df[col].mode().empty else None
                }
        
        stats_df = pd.DataFrame(stats_dict).T
        return stats_df
    except Exception as e:
        st.error(f"Error calculating statistics: {str(e)}")
        return None


def create_pivot_table(df, index_col, columns_col, values_col, aggfunc):
    """
    Create pivot table from DataFrame
    
    Args:
        df: pandas DataFrame
        index_col: Column for rows
        columns_col: Column for columns
        values_col: Column for values
        aggfunc: Aggregation function (sum, mean, count, min, max)
        
    Returns:
        Pivot table DataFrame or None on error
    """
    try:
        pivot = pd.pivot_table(df, index=index_col, columns=columns_col, 
                              values=values_col, aggfunc=aggfunc, fill_value=0)
        return pivot
    except Exception as e:
        st.error(f"Error creating pivot table: {str(e)}")
        return None


def filter_data(df, column, condition, value):
    """
    Filter DataFrame based on condition
    
    Args:
        df: pandas DataFrame
        column: Column name to filter
        condition: Filter condition (equals, contains, greater than, less than, not equals)
        value: Value to compare against
        
    Returns:
        Filtered DataFrame
    """
    try:
        if condition == "equals":
            return df[df[column] == value]
        elif condition == "contains":
            return df[df[column].astype(str).str.contains(str(value), case=False, na=False)]
        elif condition == "greater than":
            return df[df[column] > float(value)]
        elif condition == "less than":
            return df[df[column] < float(value)]
        elif condition == "not equals":
            return df[df[column] != value]
        else:
            return df
    except Exception as e:
        st.error(f"Error filtering data: {str(e)}")
        return df


def search_in_excel(wb, search_term, case_sensitive=False):
    """
    Search for term across all sheets in workbook
    
    Args:
        wb: openpyxl Workbook object
        search_term: Text to search for
        case_sensitive: Whether to match case
        
    Returns:
        DataFrame with search results (Sheet, Cell, Value)
    """
    results = []
    try:
        for sheet_name in wb.sheetnames:
            sheet = wb[sheet_name]
            for row in sheet.iter_rows():
                for cell in row:
                    if cell.value:
                        cell_value = str(cell.value)
                        search_value = search_term if case_sensitive else search_term.lower()
                        compare_value = cell_value if case_sensitive else cell_value.lower()
                        
                        if search_value in compare_value:
                            results.append({
                                'Sheet': sheet_name,
                                'Cell': cell.coordinate,
                                'Value': cell_value
                            })
        return pd.DataFrame(results)
    except Exception as e:
        st.error(f"Error searching: {str(e)}")
        return pd.DataFrame()
