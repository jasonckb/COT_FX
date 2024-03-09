import streamlit as st
import pandas as pd
import plotly.graph_objects as go

# Read data from Dropbox
@st.cache
def load_data():
    url = "https://www.dropbox.com/scl/fi/fj9ovd8bn7c6ntbp0i0uw/COT-Report.xlsx?rlkey=9ag1xpfm8v1wvkg1xqm0m2bun&dl=1"
    return pd.read_excel(url, sheet_name=None)

data = load_data()

# Sidebar for sheet selection
sheet = st.sidebar.selectbox("Select a sheet:", options=['Summary'] + [s for s in data.keys() if s != 'Summary'])

# Display data table
st.dataframe(data[sheet])

# Dashboard for the selected asset, excluding summary
if sheet != 'Summary':
    st.header(f"Dashboard for {sheet}")

    # Assuming you need to show headers in the first column of the dashboard
    st.write("Headers:", data[sheet].columns.tolist())

    # Example Placeholder for second column - modify as per your needs
    st.write("Details or additional data could be shown here.")

    # Chart 1: Long, Short, and Net Positions
    chart_data = data[sheet].tail(20)
    fig = go.Figure()
    fig.add_trace(go.Scatter(x=chart_data.index, y=chart_data['Long'], mode='lines', name='Long'))
    fig.add_trace(go.Scatter(x=chart_data.index, y=chart_data['Short'], mode='lines', name='Short'))
    fig.add_trace(go.Scatter(x=chart_data.index, y=chart_data['Net Position'], mode='lines', name='Net Position'))
    st.plotly_chart(fig, use_container_width=True)

    # Chart 2: Net Position and 13-week MA
    chart_data['MA'] = chart_data['Net Position'].rolling(window=13).mean()
    fig2 = go.Figure()
    fig2.add_trace(go.Scatter(x=chart_data.index, y=chart_data['Net Position'], mode='lines', name='Net Position'))
    fig2.add_trace(go.Scatter(x=chart_data.index, y=chart_data['MA'], mode='lines', name='13w MA', line=dict(dash='dot')))
    st.plotly_chart(fig2, use_container_width=True)
