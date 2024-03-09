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

    return formatted_sheet

# Read data from Dropbox
@st.cache(allow_output_mutation=True)
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

# Display data table for the selected sheet with formatting applied
st.dataframe(data[sheet])
# Assuming all other code up to dataframe display remains unchanged.

# Display available columns for the selected sheet
if sheet.lower() != 'summary':
    st.write("Available columns in the selected sheet:")
    st.write(data[sheet].columns.tolist())

if sheet.lower() != 'summary':
    # Adjust these column names based on what is actually available
    date_column = 'Date'
    long_column = 'Long'
    short_column = 'Short'
    net_position_column = 'Net Position'  # Change this if the actual name is different

    if {date_column, long_column, short_column, net_position_column}.issubset(data[sheet].columns):
        chart_data = data[sheet].head(20)  # or .tail(20) for the most recent data
        
        # Chart 1: Long, Short, and Net Positions
        fig = go.Figure()
        fig.add_trace(go.Scatter(x=chart_data[date_column], y=chart_data[long_column], mode='lines', name='Long'))
        fig.add_trace(go.Scatter(x=chart_data[date_column], y=chart_data[short_column], mode='lines', name='Short'))
        fig.add_trace(go.Scatter(x=chart_data[date_column], y=chart_data[net_position_column], mode='lines', name='Net Position'))
        st.plotly_chart(fig, use_container_width=True)

        # Chart 2: Net Position and its 13-week Moving Average
        chart_data['13w MA'] = chart_data[net_position_column].rolling(window=13, min_periods=1).mean()
        fig2 = go.Figure()
        fig2.add_trace(go.Scatter(x=chart_data[date_column], y=chart_data[net_position_column], mode='lines', name='Net Position'))
        fig2.add_trace(go.Scatter(x=chart_data[date_column], y=chart_data['13w MA'], mode='lines+markers', name='13w MA', line=dict(dash='dot')))
        st.plotly_chart(fig2, use_container_width=True)
    else:
        st.error("Required columns for plotting are not found in the selected sheet.")

