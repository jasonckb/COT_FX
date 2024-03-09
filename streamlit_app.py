import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import openpyxl

# Helper function to format percentage columns
def format_percentage_columns(sheet_data):
    for col in sheet_data.columns:
        # Identify columns that should be formatted as percentages
        if col.endswith('%'):
            sheet_data[col] = pd.to_numeric(sheet_data[col], errors='coerce').apply(lambda x: f"{x:.2%}" if pd.notnull(x) else x)
    return sheet_data

# Read data from Dropbox and skip the first row for "Summary"
@st.cache(allow_output_mutation=True)
def load_data():
    url = "https://www.dropbox.com/scl/fi/c50v70ob66syx58vtc028/COT-Report.xlsx?rlkey=3fu2xoqsln3gaj084hw0rfcw0&dl=1"
    # Use `None` to load all sheets
    all_sheets_data = pd.read_excel(url, sheet_name=None, engine='openpyxl', skiprows=[0] if sheet_name == "Summary" else None)
    
    # Format the data for all sheets
    for sheet_name, sheet_data in all_sheets_data.items():
        all_sheets_data[sheet_name] = format_percentage_columns(sheet_data)

    return all_sheets_data

data = load_data()

# Sidebar for sheet selection
sheet_names = list(data.keys())  # Maintain the order of sheets
sheet = st.sidebar.selectbox("Select a sheet:", options=sheet_names)

# Display data table for the selected sheet with percentage formatting applied
if sheet == "Summary":
    st.dataframe(data[sheet].iloc[1:])  # Skip the first row for "Summary"
else:
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

