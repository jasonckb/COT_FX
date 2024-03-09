import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import openpyxl

# Helper function to clean and format data
def clean_and_format_data(sheet_data):
    formatted_sheet = sheet_data.copy()
    # Remove commas for numeric columns and parentheses for negative numbers
    formatted_sheet = formatted_sheet.replace({'\,': '', '\(': '-', '\)': ''}, regex=True)
    
    for col in formatted_sheet.columns:
        try:
            # Attempt to convert to numeric, errors='coerce' will replace non-numeric values with NaN
            formatted_sheet[col] = pd.to_numeric(formatted_sheet[col], errors='coerce')
        except ValueError:
            # If the column cannot be converted to numeric, likely a string, skip it
            continue

    # Format percentage columns after numeric conversion
    for col in formatted_sheet.columns:
        # Identify columns that should be formatted as percentages
        if isinstance(col, str) and col.strip().endswith('%'):
            formatted_sheet[col] = formatted_sheet[col].astype(float) / 100

    # Parse the 'Date' column if it exists
    if 'Date' in formatted_sheet.columns:
        formatted_sheet['Date'] = pd.to_datetime(formatted_sheet['Date'], dayfirst=True, errors='coerce')

    # Trim whitespace from column names and values
    formatted_sheet.columns = formatted_sheet.columns.str.strip()
    formatted_sheet = formatted_sheet.applymap(lambda x: x.strip() if isinstance(x, str) else x)

    return formatted_sheet

# Read data from Dropbox
@st.cache_data(show_spinner=False)
def load_data():
    url = "https://www.dropbox.com/scl/fi/c50v70ob66syx58vtc028/COT-Report.xlsx?rlkey=3fu2xoqsln3gaj084hw0rfcw0&dl=1"
    xls = pd.ExcelFile(url, engine='openpyxl')
    all_sheets_data = {}
    for sheet_name in xls.sheet_names:
        sheet_data = pd.read_excel(xls, sheet_name=sheet_name)
        all_sheets_data[sheet_name] = clean_and_format_data(sheet_data)
    return all_sheets_data

data = load_data()

# Sidebar for sheet selection
sheet_names = list(data.keys())  # Maintain the order of sheets
sheet = st.sidebar.selectbox("Select a sheet:", options=sheet_names)

# Print column names and data types for debugging
st.write("Column names:")
st.write(data[sheet].columns)
st.write("Data types:")
st.write(data[sheet].dtypes)

# Display data table for the selected sheet with formatting applied
st.table(data[sheet])

