import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import openpyxl
import yfinance as yf

st.set_page_config(page_title="CFTC COT Report & FX Dashboard", layout="wide")
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
st.dataframe(data[sheet], width=None)

# Determine if the selected sheet should not display charts
sheets_without_charts = ['summary', 'fx_supply_demand_swing']

# Check if the selected sheet should plot charts
if sheet.lower() not in sheets_without_charts:
    # Load the last 20 entries of the selected sheet for chart plotting
    chart_data = data[sheet].head(20)
    # Generate the dynamic column name for net positions
    net_position_column = f"{sheet.replace(' ', '')} Net Positions" if ' ' in sheet else f"{sheet} Net Positions"
    
    # Check if the dynamically generated column name is in the DataFrame
    if net_position_column not in chart_data.columns:
        # If not, print an error message and list available columns
        st.error(f"Column {net_position_column} not found in the data.")
        st.write("Available columns: ", chart_data.columns.tolist())
        # Optionally, halt further execution for this sheet
        st.stop()
    else:
        # Proceed with plotting since the column exists
        col1, col2 = st.columns(2)

    with col1:
        fig = go.Figure()
        fig.add_trace(go.Scatter(x=chart_data['Date'], y=chart_data['Long'], mode='lines', name='Long', line=dict(color='blue')))
        fig.add_trace(go.Scatter(x=chart_data['Date'], y=chart_data['Short'], mode='lines', name='Short', line=dict(color='red')))
        # Use the net_position_column safely as it is defined above
        fig.add_trace(go.Scatter(x=chart_data['Date'], y=chart_data[net_position_column], mode='lines', name=net_position_column, line=dict(color='darkgray', width=3)))
        st.plotly_chart(fig, use_container_width=True)

    with col2:
        # Ensure '13w MA' is based on net_position_column and computed right before this plot
        chart_data['13w MA'] = chart_data[net_position_column].rolling(window=13, min_periods=1).mean()
        fig2 = go.Figure()
        fig2.add_trace(go.Scatter(x=chart_data['Date'], y=chart_data[net_position_column], mode='lines', name=net_position_column, line=dict(color='black', width=3)))
        fig2.add_trace(go.Scatter(x=chart_data['Date'], y=chart_data['13w MA'], mode='lines+markers', name='13 Week MA', line=dict(dash='dot', color='darkgray')))
        st.plotly_chart(fig2, use_container_width=True)

# Define a function to get the latest price from Yahoo Finance
def get_latest_price(symbol):
    try:
        # Fetching the latest price for the symbol
        latest_price_data = yf.download(symbol, period="1d")
        return latest_price_data['Close'].iloc[-1]
    except Exception as e:
        st.error(f"Error fetching data for {symbol}: {e}")
        return None

# Your main app logic
# Assuming 'data' is your main DataFrame and 'sheet' represents the active sheet name
if 'fx_supply_demand_swing' == sheet.lower():
    # Sidebar button to refresh FX rates
    if st.sidebar.button('Refresh FX Rate'):
        # Refresh the latest prices for the 'FX_Supply_Demand_Swing' sheet
        for idx, symbol in enumerate(data[sheet]['Symbol']):
            yahoo_symbol = f"{symbol}=X"
            data[sheet].at[idx, 'Latest Price'] = get_latest_price(yahoo_symbol)

    # Assuming setup levels are the next four columns after 'Symbol'
    # Display the data including the latest price and setup levels
    st.table(data[sheet])