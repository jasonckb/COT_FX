import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import openpyxl  # Explicit import for clarity

# Function to format the dataframe for display
def format_dataframe(df, percentage_columns):
    formatted_df = df.copy()
    for col in percentage_columns:
        formatted_df[col] = pd.to_numeric(formatted_df[col], errors='coerce').apply(lambda x: f'{x:.2%}')
    return formatted_df

# Read data from Dropbox and apply formatting
@st.cache
def load_data():
    url = "https://www.dropbox.com/scl/fi/fj9ovd8bn7c6ntbp0i0uw/COT-Report.xlsx?rlkey=9ag1xpfm8v1wvkg1xqm0m2bun&dl=1"
    data = pd.read_excel(url, sheet_name=None, engine='openpyxl')

    # Determine percentage columns for each sheet
    percentage_columns = {
        sheet_name: [col for col in sheet_data.columns if '%' in col or sheet_data[col].dtype in ['float64', 'float32']]
        for sheet_name, sheet_data in data.items()
    }

    # Format each sheet appropriately
    for sheet_name, sheet_data in data.items():
        if sheet_name.lower() == 'summary':
            # Create custom header by merging the first row for the 'Summary' sheet
            #merged_header = ' '.join(sheet_data.iloc[0].dropna().astype(str))
            sheet_data.columns = [merged_header] + sheet_data.columns.tolist()[1:]
            data[sheet_name] = sheet_data.iloc[1:]
        # Apply percentage formatting to specified columns
        data[sheet_name] = format_dataframe(sheet_data, percentage_columns[sheet_name])

    return data

data = load_data()

# Sidebar for sheet selection
sheet = st.sidebar.selectbox("Select a sheet:", options=['Summary'] + [s for s in data.keys() if s.lower() != 'summary'])

# Display data table
if sheet.lower() == 'summary':
    # Display the 'Summary' sheet with merged header and formatted percentages
    st.write(data[sheet].columns[0])  # Write out the merged header
    st.dataframe(data[sheet].iloc[:, 1:])  # Display the rest of the dataframe excluding the merged header
else:
    # Display data for other sheets with percentage formatting
    st.dataframe(data[sheet])

    # Chart 1: Long, Short, and Net Positions
    # Convert percentage strings back to floats for plotting
    chart_data = data[sheet].tail(20).apply(pd.to_numeric, errors='coerce')
    fig = go.Figure()
    fig.add_trace(go.Scatter(x=chart_data.index, y=chart_data['Long'], mode='lines', name='Long'))
    fig.add_trace(go.Scatter(x=chart_data.index, y=chart_data['Short'], mode='lines', name='Short'))
    fig.add_trace(go.Scatter(x=chart_data.index, y=chart_data['Net Position'], mode='lines', name='Net Position'))
    st.plotly_chart(fig, use_container_width=True)

    # Chart 2: Net Position and its 13-week Moving Average
    chart_data['MA'] = chart_data['Net Position'].rolling(window=13).mean()
    fig2 = go.Figure()
    fig2.add_trace(go.Scatter(x=chart_data.index, y=chart_data['Net Position'], mode='lines', name='Net Position'))
    fig2.add_trace(go.Scatter(x=chart_data.index, y=chart_data['MA'], mode='lines', name='13w MA', line=dict(dash='dot')))
    st.plotly_chart(fig2, use_container_width=True)
