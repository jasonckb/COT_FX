import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import openpyxl  # Explicit import for clarity

# Function to format the dataframe for display
def format_dataframe(df):
    formatted_df = df.copy()
    for col in formatted_df.columns[1:]:  # Skip the first column
        formatted_df[col] = pd.to_numeric(formatted_df[col], errors='coerce').apply(lambda x: f'{x:.2%}')
    return formatted_df

# Read data from Dropbox and apply formatting
@st.cache
def load_data():
    url = "https://www.dropbox.com/scl/fi/fj9ovd8bn7c6ntbp0i0uw/COT-Report.xlsx?rlkey=9ag1xpfm8v1wvkg1xqm0m2bun&dl=1"
    data = pd.read_excel(url, sheet_name=None, engine='openpyxl')
    
    # Format the 'Summary' sheet differently from the rest
    for sheet_name in data:
        if sheet_name.lower() == 'summary':
            # Format 'Summary' sheet and merge first row by creating a custom header
            summary_df = data[sheet_name]
            header = " ".join(summary_df.iloc[0].dropna().astype(str))
            data[sheet_name] = summary_df[1:]  # Skip the first row as it's now in the header
            data[sheet_name].columns = [header] + summary_df.columns.tolist()[1:]  # Add the merged header
        else:
            # Format other sheets with percentage values
            data[sheet_name] = format_dataframe(data[sheet_name])
    
    return data

data = load_data()

# Sidebar for sheet selection
sheet = st.sidebar.selectbox("Select a sheet:", options=['Summary'] + [s for s in data.keys() if s.lower() != 'summary'])

# Display data table
if sheet.lower() == 'summary':
    # Apply specific formatting for 'Summary' sheet
    st.write(data[sheet].columns[0])  # Custom header
    st.dataframe(data[sheet].iloc[:, 1:])  # Exclude the first column which is the custom header
else:
    # Display data for other sheets normally
    st.dataframe(data[sheet])

# Dashboard and plotting for the selected asset
if sheet != 'Summary':
    st.header(f"Dashboard for {sheet}")

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
