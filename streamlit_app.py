import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import openpyxl  # Explicit import for clarity

# Read data from Dropbox and apply formatting
@st.cache
def load_data():
    url = "https://www.dropbox.com/scl/fi/fj9ovd8bn7c6ntbp0i0uw/COT-Report.xlsx?rlkey=9ag1xpfm8v1wvkg1xqm0m2bun&dl=1"
    data = pd.read_excel(url, sheet_name=None, engine='openpyxl')
    
    # Format all numeric data as percentages
    for sheet_name, sheet_data in data.items():
        for col in sheet_data.select_dtypes(include=['float64']).columns:
            sheet_data[col] = sheet_data[col].apply(lambda x: f'{x:.2%}')
    
    return data

data = load_data()

# Sidebar for sheet selection
sheet = st.sidebar.selectbox("Select a sheet:", options=['Summary'] + [s for s in data.keys() if s != 'Summary'])

# Display data table
if sheet == 'Summary':
    # Apply specific formatting for 'Summary' sheet
    st.dataframe(data[sheet].style.set_properties(**{'text-align': 'left'}), width=None, height=None)
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
