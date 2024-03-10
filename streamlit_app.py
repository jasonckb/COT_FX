import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import openpyxl

# Add a big heading
st.title('CFTC COT Report & FX Dashboard by Jason Chan')

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
        formatted_sheet['% Long'] = formatted_sheet['% Long'].apply(lambda x: '{:.1%}'.format(x))
    if '% Short' in formatted_sheet.columns:
        formatted_sheet['% Short'] = formatted_sheet['% Short'].apply(lambda x: '{:.1%}'.format(x))
    # Parse the 'Date' column if it exists
    if 'Date' in formatted_sheet.columns:
        formatted_sheet['Date'] = pd.to_datetime(formatted_sheet['Date'], format='%d/%m/%Y', errors='coerce').dt.date
    return formatted_sheet

# Load data function
@st.cache_data(show_spinner=False)
def load_data():
    url = "https://www.dropbox.com/scl/fi/c50v70ob66syx58vtc028/COT-Report.xlsx?rlkey=3fu2xoqsln3gaj084hw0rfcw0&dl=1"
    xls = pd.ExcelFile(url, engine='openpyxl')
    all_sheets_data = {}
    for sheet_name in xls.sheet_names:
        sheet_data = pd.read_excel(xls, sheet_name=sheet_name, header=0)
        all_sheets_data[sheet_name] = clean_and_format_data(sheet_data)
    return all_sheets_data

# Loading and displaying data
data = load_data()
sheet_names = list(data.keys())
sheet = st.sidebar.selectbox("Select a sheet:", options=sheet_names)

# Make table wider
st.dataframe(data[sheet], width=1000)

# Placing two charts side by side
if sheet.lower() != 'summary':
    chart_data = data[sheet].head(20)

    # Initialize two-column layout
    col1, col2 = st.columns(2)

    # Chart 1 in column 1
    with col1:
        fig = go.Figure()
        fig.add_trace(go.Scatter(x=chart_data['Date'], y=chart_data['Long'], mode='lines', name='Long', line=dict(color='blue')))
        fig.add_trace(go.Scatter(x=chart_data['Date'], y=chart_data['Short'], mode='lines', name='Short', line=dict(color='red')))
        net_position_column = f"{sheet.replace(' ', '')} Net Positions" if ' ' in sheet else f"{sheet} Net Positions"
        fig.add_trace(go.Scatter(x=chart_data['Date'], y=chart_data[net_position_column], mode='lines', name=net_position_column, line=dict(color='darkgray', width=3)))
        st.plotly_chart(fig, use_container_width=True)

    # Chart 2 in column 2
    with col2:
        fig2 = go.Figure()
        fig2.add_trace(go.Scatter(x=chart_data['Date'], y=chart_data[net_position_column], mode='lines', name=net_position_column, line=dict(color='black', width=3)))
        fig2.add_trace(go.Scatter(x=chart_data['Date'], y=chart_data['13w MA'], mode='lines+markers', name='13w MA', line=dict(dash='dot', color='darkgray')))
        st.plotly_chart(fig2, use_container_width=True)

