import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import openpyxl

# Helper function to clean and format data
def clean_and_format_data(sheet_data):
    formatted_sheet = sheet_data.copy()

    # Remove commas for numeric columns and parentheses for negative numbers
    formatted_sheet = formatted_sheet.replace({'\,' : '', '\(': '-', '\)': ''}, regex=True)

    # Convert '% Long' and '% Short' columns to numeric
    if '% Long' in formatted_sheet.columns:
        formatted_sheet['% Long'] = pd.to_numeric(formatted_sheet['% Long'], errors='coerce')

    if '% Short' in formatted_sheet.columns:
        formatted_sheet['% Short'] = pd.to_numeric(formatted_sheet['% Short'], errors='coerce')

    # Format '% Long' and '% Short' columns
    if '% Long' in formatted_sheet.columns:
        formatted_sheet['% Long'] = formatted_sheet['% Long'].apply(lambda x: '{:.1%}'.format(x ))

    if '% Short' in formatted_sheet.columns:
        formatted_sheet['% Short'] = formatted_sheet['% Short'].apply(lambda x: '{:.1%}'.format(x))

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
    chart_data = data[sheet].head(20)  # or use .tail(20) as needed

    # Chart 1: Long (blue), Short (red), and Net Positions (dark gray)
    fig = go.Figure()
    fig.add_trace(go.Scatter(x=chart_data['Date'], y=chart_data['Long'], mode='lines', name='Long', line=dict(color='blue')))
    fig.add_trace(go.Scatter(x=chart_data['Date'], y=chart_data['Short'], mode='lines', name='Short', line=dict(color='red')))
    net_position_column = f"{sheet.replace(' ', '')} Net Positions" if ' ' in sheet else f"{sheet} Net Positions"
    fig.add_trace(go.Scatter(x=chart_data['Date'], y=chart_data[net_position_column], mode='lines', name=net_position_column, line=dict(color='darkgray')))
    st.plotly_chart(fig, use_container_width=True)

    # Chart 2: Net Position (dark gray) and its 13-week Moving Average
    chart_data['13w MA'] = chart_data[net_position_column].rolling(window=13, min_periods=1).mean()
    fig2 = go.Figure()
    fig2.add_trace(go.Scatter(x=chart_data['Date'], y=chart_data[net_position_column], mode='lines', name=net_position_column, line=dict(color='black')))
    fig2.add_trace(go.Scatter(x=chart_data['Date'], y=chart_data['13 Week MA'], mode='lines+markers', name='13 Week MA', line=dict(dash='dot', color='darkgray')))
    st.plotly_chart(fig2, use_container_width=True)
