import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import openpyxl

# Function to format percentage columns
def format_percentage_columns(sheet_data):
    formatted_sheet = sheet_data.copy()
    for col in formatted_sheet.columns:
        # Check if column header indicates a percentage value
        # Stripping whitespace to ensure correct identification
        if "%" in col.strip():
            formatted_sheet[col] = formatted_sheet[col].apply(lambda x: "{:.2%}".format(x) if pd.notnull(x) else x)
    return formatted_sheet

# Read data from Dropbox and apply formatting
@st.cache(allow_output_mutation=True)
def load_data():
    url = "https://www.dropbox.com/scl/fi/fj9ovd8bn7c6ntbp0i0uw/COT-Report.xlsx?rlkey=9ag1xpfm8v1wvkg1xqm0m2bun&dl=1"
    all_sheets_data = pd.read_excel(url, sheet_name=None, engine='openpyxl')
    
    for sheet_name, sheet_data in all_sheets_data.items():
        all_sheets_data[sheet_name] = format_percentage_columns(sheet_data)

    return all_sheets_data

data = load_data()

# Sidebar for sheet selection
sheet_names = list(data.keys())  # Maintain the order of sheets
sheet = st.sidebar.selectbox("Select a sheet:", options=sheet_names)

# Display data table for the selected sheet with percentage formatting applied
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

