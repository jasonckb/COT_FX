import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import openpyxl

# Helper function to clean and format data
def clean_and_format_data(sheet_data):
    formatted_sheet = sheet_data.copy()

    # Remove commas for numeric columns and parentheses for negative numbers
    formatted_sheet = formatted_sheet.replace({'\,' : '', '\(': '-', '\)': ''}, regex=True)

    # Convert '% Long' and '% Short' columns to numeric, assuming they are like '60%' and should become 60 (not 0.6)
    if '% Long' in formatted_sheet.columns:
        formatted_sheet['% Long'] = formatted_sheet['% Long'].str.rstrip('%').astype('float') / 100

    if '% Short' in formatted_sheet.columns:
        formatted_sheet['% Short'] = formatted_sheet['% Short'].str.rstrip('%').astype('float') / 100

    # Since now they are properly in decimal form representing percentages, no further formatting is required
    # If you need them in string format for some reason, multiply by 100 and append '%'
    # However, it's usually best to keep them as numeric for any calculations and only format them when displaying

    # Parse the 'Date' column if it exists
    if 'Date' in formatted_sheet.columns:
        formatted_sheet['Date'] = pd.to_datetime(formatted_sheet['Date'], format='%d/%m/%Y', errors='coerce').dt.date

    return formatted_sheet




# Read data from Dropbox
@st.cache_data(show_spinner=False)
# Inside your load_data function:
def load_data():
    url = "https://www.dropbox.com/scl/fi/c50v70ob66syx58vtc028/COT-Report.xlsx?rlkey=3fu2xoqsln3gaj084hw0rfcw0&dl=1"
    xls = pd.ExcelFile(url, engine='openpyxl')
    all_sheets_data = {}
    for sheet_name in xls.sheet_names:
        # Ensuring that the first row is used as header
        sheet_data = pd.read_excel(xls, sheet_name=sheet_name, header=0)  
        all_sheets_data[sheet_name] = clean_and_format_data(sheet_data)
    return all_sheets_data


data = load_data()

# Sidebar for sheet selection
sheet_names = list(data.keys())  # Maintain the order of sheets
sheet = st.sidebar.selectbox("Select a sheet:", options=sheet_names)

# Display data table for the selected sheet with formatting applied
st.dataframe(data[sheet])


if sheet.lower() != 'summary':
    chart_data = data[sheet].head(20)

    # Direct usage of the column name without reconstruction
    net_position_column = 'SP500 Net Positions'

    # Proceed with plotting if the column exists
    if net_position_column not in chart_data.columns:
        st.error(f"Expected column '{net_position_column}' not found. Available columns: {chart_data.columns.tolist()}")
    else:

        # Proceed with plotting if the column exists
        fig = go.Figure()
        fig.add_trace(go.Scatter(x=chart_data['Date'], y=chart_data['Long'], mode='lines', name='Long'))
        fig.add_trace(go.Scatter(x=chart_data['Date'], y=chart_data['Short'], mode='lines', name='Short'))
        fig.add_trace(go.Scatter(x=chart_data['Date'], y=chart_data[net_position_column], mode='lines', name=net_position_column))
        st.plotly_chart(fig, use_container_width=True)

        # Chart 2: Net Position and its 13-week Moving Average
        chart_data[net_position_column] = pd.to_numeric(chart_data[net_position_column], errors='coerce')
        chart_data['13w MA'] = chart_data[net_position_column].rolling(window=13, min_periods=1).mean()
        fig2 = go.Figure()
        fig2.add_trace(go.Scatter(x=chart_data['Date'], y=chart_data[net_position_column], mode='lines', name=net_position_column))
        fig2.add_trace(go.Scatter(x=chart_data['Date'], y=chart_data['13w MA'], mode='lines+markers', name='13w MA', line=dict(dash='dot')))
        st.plotly_chart(fig2, use_container_width=True)
