import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import openpyxl

# Helper function to clean and format data
def clean_and_format_data(sheet_data):
    formatted_sheet = sheet_data.copy()

    # Identify numeric columns
    numeric_columns = formatted_sheet.select_dtypes(include=['int64', 'float64']).columns

    # Remove commas for numeric columns and parentheses for negative numbers
    formatted_sheet[numeric_columns] = formatted_sheet[numeric_columns].replace({'\,' : '', '\(': '-', '\)': ''}, regex=True)

    for col in formatted_sheet.columns:
        if formatted_sheet[col].dtype != 'object':  # Skip the conversion for non-numeric columns
            try:
                # Attempt to convert to numeric, errors='coerce' will replace non-numeric values with NaN
                formatted_sheet[col] = pd.to_numeric(formatted_sheet[col], errors='coerce')
            except ValueError:
                # If the column cannot be converted to numeric, likely a string, skip it
                continue

    # Format '% Long' and '% Short' columns
    if '% Long' in formatted_sheet.columns:
        formatted_sheet['% Long'] = formatted_sheet['% Long'].apply(lambda x: '{:.1%}'.format(x / 100))

    if '% Short' in formatted_sheet.columns:
        formatted_sheet['% Short'] = formatted_sheet['% Short'].apply(lambda x: '{:.1%}'.format(x / 100))

    # Parse the 'Date' column if it exists
    if 'Date' in formatted_sheet.columns:
        formatted_sheet['Date'] = pd.to_datetime(formatted_sheet['Date'], dayfirst=True, errors='coerce').dt.date

    return formatted_sheet



# Read data from Dropbox
@st.cache_data(show_spinner=False)
def load_data():
    url = "https://www.dropbox.com/scl/fi/c50v70ob66syx58vtc028/COT-Report.xlsx?rlkey=3fu2xoqsln3gaj084hw0rfcw0&dl=1"
    xls = pd.ExcelFile(url, engine='openpyxl')
    all_sheets_data = {}
    for sheet_name in xls.sheet_names:
        sheet_data = pd.read_excel(xls, sheet_name=sheet_name, header=None)  # Specify that there is no header
        all_sheets_data[sheet_name] = clean_and_format_data(sheet_data)
    return all_sheets_data


data = load_data()

# Sidebar for sheet selection
sheet_names = list(data.keys())  # Maintain the order of sheets
sheet = st.sidebar.selectbox("Select a sheet:", options=sheet_names)

# Display data table for the selected sheet with formatting applied
st.dataframe(data[sheet])




# Display interactive charts for selected sheet, excluding 'Summary'
if sheet.lower() != 'summary':
    # Chart 1: Long, Short, and Net Positions
    chart_data = data[sheet].tail(20)
    fig = go.Figure()
    fig.add_trace(go.Scatter(x=chart_data['Date'], y=chart_data['Long'], mode='lines', name='Long'))
    fig.add_trace(go.Scatter(x=chart_data['Date'], y=chart_data['Short'], mode='lines', name='Short'))
    fig.add_trace(go.Scatter(x=chart_data['Date'], y=chart_data['Net Position'], mode='lines', name='Net Position'))
    st.plotly_chart(fig, use_container_width=True)

    # Chart 2: Net Position and its 13-week Moving Average
    chart_data['Net Position'] = pd.to_numeric(chart_data['Net Position'], errors='coerce')  # Ensure numeric for MA calculation
    chart_data['MA'] = chart_data['Net Position'].rolling(window=13).mean()
    fig2 = go.Figure()
    fig2.add_trace(go.Scatter(x=chart_data['Date'], y=chart_data['Net Position'], mode='lines', name='Net Position'))
    fig2.add_trace(go.Scatter(x=chart_data['Date'], y=chart_data['MA'], mode='lines+markers', name='13w MA', line=dict(dash='dot')))
    st.plotly_chart(fig2, use_container_width=True)


