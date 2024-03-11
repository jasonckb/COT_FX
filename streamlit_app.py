import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
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
if sheet.lower() != 'fx_supply_demand_swing':
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
    # Sort the data by 'Date' in ascending order for correct moving average calculation
        chart_data_sorted = chart_data.sort_values('Date', ascending=True)

    # Calculate the 13-week moving average using the sorted data
        chart_data_sorted['13w MA'] = chart_data_sorted[net_position_column].rolling(window=13, min_periods=1).mean()

    # Create the figure for plotting
        fig2 = go.Figure()

    # Add the net position line
        fig2.add_trace(go.Scatter(x=chart_data_sorted['Date'], y=chart_data_sorted[net_position_column],
                              mode='lines', name=net_position_column, line=dict(color='black', width=3)))

    # Add the 13-week moving average line
        fig2.add_trace(go.Scatter(x=chart_data_sorted['Date'], y=chart_data_sorted['13w MA'],
                              mode='lines+markers', name='13 Week MA', line=dict(dash='dot', color='darkgray')))

    # Plot the chart using the original order of dates (descending if that's your requirement)
        st.plotly_chart(fig2, use_container_width=True)


# Define a function to color the rows based on conditions
def color_rows(df):
    # Create a new DataFrame with the same index and columns as the input DataFrame
    colored_df = pd.DataFrame('', index=df.index, columns=df.columns)

    # Loop through each row in the input DataFrame
    for idx, row in df.iterrows():
        # Check if any of the required values are NA
        if pd.isna(row['Latest Price']) or pd.isna(row['1st Long Setup']) or pd.isna(row['2nd Long Setup']) or pd.isna(row['1st short Setup']) or pd.isna(row['2nd short Setup']):
            continue

        # Check if the latest price is close to 0.1% of one of the long set up levels
        if row['Latest Price'] >= row['1st Long Setup'] * 0.999 and row['Latest Price'] <= row['1st Long Setup'] * 1.001:
            colored_df.loc[idx, :] = 'background-color:  #89ef97'
        elif row['Latest Price'] >= row['2nd Long Setup'] * 0.999 and row['Latest Price'] <= row['2nd Long Setup'] * 1.001:
            colored_df.loc[idx, :] = 'background-color:  #89ef97'
        # Check if the latest price is close to 0.1% of one of the short set up levels
        elif row['Latest Price'] >= row['1st short Setup'] * 0.999 and row['Latest Price'] <= row['1st short Setup'] * 1.001:
            colored_df.loc[idx, :] = 'background-color: #efb4bd'
        elif row['Latest Price'] >= row['2nd short Setup'] * 0.999 and row['Latest Price'] <= row['2nd short Setup'] * 1.001:
            colored_df.loc[idx, :] = 'background-color: #efb4bd'

    return colored_df

# Define a function to fetch historical data for a given symbol
def get_historical_data(symbol):
    try:
        historical_data = yf.download(symbol, period="3mo")
        return historical_data[['Open', 'High', 'Low', 'Close']]
    except Exception as e:
        st.error(f"Error fetching data for {symbol}: {e}")
        return None
# Define a function to fetch the latest price for a given symbol
def get_latest_price(symbol):
    try:
        latest_price_data = yf.download(symbol, period="1d")
        return latest_price_data['Close'].iloc[-1]
    except Exception as e:
        st.error(f"Error fetching data for {symbol}: {e}")
        return None

if sheet.lower() == 'fx_supply_demand_swing':
    # Refresh button in the sidebar
    if st.sidebar.button('Refresh FX Rate'):
        # Fetch and update latest prices
        for idx, row in data[sheet].iterrows():
            symbol = row['Symbol'].replace("/", "") + "=X"
            data[sheet].at[idx, 'Latest Price'] = get_latest_price(symbol)

    # Make sure 'Latest Price' exists before reordering columns
    if 'Latest Price' not in data[sheet].columns:
        data[sheet]['Latest Price'] = pd.NA

    # Ensure the 'Latest Price' column is second
    # First, get all columns excluding 'Symbol' and 'Latest Price'
    other_cols = [col for col in data[sheet].columns if col not in ['Symbol', 'Latest Price']]
    # Define new column order
    new_cols = ['Symbol', 'Latest Price'] + other_cols
    # Reorder the DataFrame
    data[sheet] = data[sheet][new_cols]

    # Apply the coloring function to the DataFrame
    styled_data = data[sheet].style.apply(color_rows, axis=None)

    # Display the full length table without scrolling
    st.write(styled_data)
    
if sheet.lower() == 'fx_supply_demand_swing' and 'fx_supply_demand_swing' in data:
    # Extract unique symbols from the 'fx_supply_demand_swing' sheet
    symbols = data['fx_supply_demand_swing']['Symbol'].unique()

    # Add a select box for symbols
    selected_symbol = st.selectbox("Select a symbol:", options=symbols)

    if selected_symbol:
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



