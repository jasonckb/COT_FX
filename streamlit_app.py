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

# Separate handling for 'FX_Supply_Demand_Swing' at the end, without an else
if sheet.lower() == 'fx_supply_demand_swing':
    # Display the original data table
    st.dataframe(data[sheet])

    dashboard_data = data[sheet].copy()
    dashboard_data['Latest Price'] = pd.NA  # Initialize Latest Price column

    for idx, row in dashboard_data.iterrows():
        symbol = f"{row['Symbol']}=X"  # Adjust for Forex symbol format
        try:
            # Fetch the latest price
            latest_price = yf.download(symbol, period="1d")['Close'].iloc[-1]
            dashboard_data.at[idx, 'Latest Price'] = latest_price
        except Exception as e:
            st.error(f"Error fetching data for {row['Symbol']}: {e}")

    # Move 'Latest Price' to the second column position
    cols = list(dashboard_data.columns)
    cols.insert(1, cols.pop(cols.index('Latest Price')))
    dashboard_data = dashboard_data.loc[:, cols]

    # Determine highlight color based on comparison with setup values
    for idx, row in dashboard_data.iterrows():
        highlight_color = ''
        for setup_col in ['1st Long Setup', '2nd Long Setup']:
            if pd.notnull(row[setup_col]) and 0 < abs((row['Latest Price'] - row[setup_col]) / row[setup_col]) <= 0.001:
                highlight_color = 'green'
                break

        if not highlight_color:
            for setup_col in ['1st short Setup', '2nd short Setup']:
                if pd.notnull(row[setup_col]) and 0 < abs((row['Latest Price'] - row[setup_col]) / row[setup_col]) <= 0.001:
                    highlight_color = 'red'
                    break

        dashboard_data.at[idx, 'Highlight'] = highlight_color

    # Define function to apply conditional formatting
    def apply_highlight(row):
        # Apply the color to all cells in the row based on 'Highlight' value
        color = row['Highlight']
        return ['background-color: ' + color if color else '' for _ in row[:-1]] + ['']  # Empty style for 'Highlight'

    # Apply styling and exclude 'Highlight' column for display
    styled_df = dashboard_data.drop(columns='Highlight').style.apply(apply_highlight, axis=1)
    st.dataframe(styled_df)