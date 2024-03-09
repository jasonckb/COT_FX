import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import openpyxl


# Helper function to format percentage columns and clean numeric data
def clean_and_format_data(sheet_data):
    # Make a copy of the dataframe to prevent changes to the original data
    formatted_sheet = sheet_data.copy()

    # Convert 'Date' column to datetime
    formatted_sheet['Date'] = pd.to_datetime(formatted_sheet['Date'], errors='coerce')

    for col in formatted_sheet.columns:
        # Remove commas for thousands and convert to float
        if formatted_sheet[col].dtype == 'object':
            formatted_sheet[col] = formatted_sheet[col].replace({',': '', '\(': '-', '\)': ''}, regex=True)
            formatted_sheet[col] = pd.to_numeric(formatted_sheet[col], errors='coerce')

        # Format percentage columns
        if '%' in col:
            formatted_sheet[col] = formatted_sheet[col] / 100

    return formatted_sheet

# Read data from Dropbox
@st.cache(allow_output_mutation=True)
def load_data():
    url = "YOUR_DROPBOX_LINK_HERE"  # Replace with your actual Dropbox link to the Excel file
    # Read the Excel file
    all_sheets_data = pd.read_excel(url, sheet_name=None, engine='openpyxl')
    
    # Clean and format all sheets data
    for sheet_name, sheet_data in all_sheets_data.items():
        all_sheets_data[sheet_name] = clean_and_format_data(sheet_data)

    return all_sheets_data

data = load_data()

# Sidebar for sheet selection
sheet_names = list(data.keys())  # Maintain the order of sheets
sheet = st.sidebar.selectbox("Select a sheet:", options=sheet_names)

# Display the data table for the selected sheet with formatting applied
st.dataframe(data[sheet])

# Display interactive charts for selected sheet, excluding 'Summary'
if sheet.lower() != 'summary':
    # Extract the last 20 data points for charting
    chart_data = data[sheet].tail(20)

    # Create Chart 1: Long, Short, and Net Positions
    fig = go.Figure()
    fig.add_trace(go.Scatter(x=chart_data['Date'], y=chart_data['Long'], mode='lines', name='Long'))
    fig.add_trace(go.Scatter(x=chart_data['Date'], y=chart_data['Short'], mode='lines', name='Short'))
    fig.add_trace(go.Scatter(x=chart_data['Date'], y=chart_data['Nasdaq 100 Net Positions'], mode='lines', name='Net Position'))

    # Show Chart 1
    st.plotly_chart(fig, use_container_width=True)

    # Create Chart 2: Net Position with 13-week Moving Average
    fig2 = go.Figure()
    fig2.add_trace(go.Scatter(x=chart_data['Date'], y=chart_data['Nasdaq 100 Net Positions'], mode='lines', name='Net Position'))
    fig2.add_trace(go.Scatter(
        x=chart_data['Date'], 
        y=chart_data['Nasdaq 100 Net Positions'].rolling(window=13, min_periods=1).mean(), 
        mode='lines',
        name='13w MA',
        line=dict(dash='dot')
    ))

    # Show Chart 2
    st.plotly_chart(fig2, use_container_width=True)



