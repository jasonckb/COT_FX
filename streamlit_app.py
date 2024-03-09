import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import openpyxl

# Helper function to format percentage columns
def format_percentage_columns(sheet_data):
    formatted_sheet = sheet_data.copy()
    for col in formatted_sheet.columns:
        # Check if '%' is in the column name (more flexible than endswith)
        if isinstance(col, str) and '%' in col:
            # Remove extra whitespace and format as percentage
            formatted_sheet[col.strip()] = pd.to_numeric(formatted_sheet[col], errors='coerce').apply(
                lambda x: f"{x:.2%}" if pd.notnull(x) else x)
    return formatted_sheet


# Read data from Dropbox
@st.cache(allow_output_mutation=True)
def load_data():
    url = "https://www.dropbox.com/scl/fi/c50v70ob66syx58vtc028/COT-Report.xlsx?rlkey=3fu2xoqsln3gaj084hw0rfcw0&dl=1"
    xls = pd.ExcelFile(url, engine='openpyxl')
    all_sheets_data = {}
    for sheet_name in xls.sheet_names:
        # Use the first row as header by default, and manually skip for 'Summary' if necessary
        header_row = 0 #if sheet_name != "Summary" else 1
        sheet_data = pd.read_excel(xls, sheet_name=sheet_name, header=header_row)
        all_sheets_data[sheet_name] = format_percentage_columns(sheet_data)
    return all_sheets_data

data = load_data()

# Sidebar for sheet selection
sheet_names = list(data.keys())  # Maintain the order of sheets
sheet = st.sidebar.selectbox("Select a sheet:", options=sheet_names)

# Display data table for the selected sheet with percentage formatting applied
st.dataframe(data[sheet])

# Rest of your Streamlit code...


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


