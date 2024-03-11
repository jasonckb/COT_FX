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

# Extract unique symbols from the 'fx_supply_demand_swing' sheet
symbols = data['fx_supply_demand_swing']['Symbol'].unique()

# Sidebar for sheet selection
sheet_names = list(data.keys())  # Maintain the order of sheets
sheet = st.sidebar.selectbox("Select a sheet:", options=sheet_names)

# Add a select box for symbols when 'fx_supply_demand_swing' is selected
if sheet.lower() == 'fx_supply_demand_swing':
    selected_symbol = st.sidebar.selectbox("Select a symbol:", options=symbols)

# Display data table for the selected sheet with formatting applied
if sheet.lower() != 'fx_supply_demand_swing':
    st.dataframe(data[sheet], width=None)

# If 'fx_supply_demand_swing' is selected, fetch historical data and plot the interactive chart
if sheet.lower() == 'fx_supply_demand_swing':
    # Fetch historical data for the selected symbol
    selected_symbol_data = data['fx_supply_demand_swing'][data['fx_supply_demand_swing']['Symbol'] == selected_symbol]
    historical_data = get_historical_data(selected_symbol_data.iloc[0]['Symbol'].replace("/", "") + "=X")

    # Get setup levels for the selected symbol
    setup_levels = [selected_symbol_data.iloc[0]['1st Long Setup'],
                    selected_symbol_data.iloc[0]['2nd Long Setup'],
                    selected_symbol_data.iloc[0]['1st short Setup'],
                    selected_symbol_data.iloc[0]['2nd short Setup']]

    # Plot the interactive chart with horizontal levels
    plot_interactive_chart(historical_data, setup_levels)

# Define a function to fetch historical data for a given symbol
def get_historical_data(symbol):
    try:
        historical_data = yf.download(symbol, period="3mo")
        return historical_data[['Open', 'High', 'Low', 'Close']]
    except Exception as e:
        st.error(f"Error fetching data for {symbol}: {e}")
        return None

# Define a function to plot an interactive chart with horizontal levels
def plot_interactive_chart(historical_data, setup_levels):
    if historical_data is not None:
        fig = go.Figure()
        fig.add_trace(go.Candlestick(x=historical_data.index,
                                     open=historical_data['Open'],
                                     high=historical_data['High'],
                                     low=historical_data['Low'],
                                     close=historical_data['Close']))

        # Add horizontal lines for setup levels
        for level in setup_levels:
            if level is not None and level != '':
                fig.add_shape(type="line",
                              x0=historical_data.index[0],
                              y0=level,
                              x1=historical_data.index[-1],
                              y1=level,
                              line=dict(color="red", width=2, dash="dot"))

        fig.update_layout(title=f"Interactive Chart for {selected_symbol}")
        st.plotly_chart(fig, use_container_width=True)

    

