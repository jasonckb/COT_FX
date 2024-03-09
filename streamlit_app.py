import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import openpyxl
import plotly.express as px
from pandas.tseries.offsets import DateOffset

# Function to clean and format the data
def clean_and_format_data(df):
    # Convert percentage columns
    for col in df.columns:
        if '%' in col:
            df[col] = df[col].str.rstrip('%').astype(float) / 100.0
    return df

# Load data from Dropbox
@st.cache
def load_data():
    url = "https://www.dropbox.com/scl/fi/c50v70ob66syx58vtc028/COT-Report.xlsx?dl=1"
    data = pd.read_excel(url, sheet_name=None, engine='openpyxl')
    for name, sheet in data.items():
        data[name] = clean_and_format_data(sheet)
    return data

# Load data
data = load_data()

# Sidebar for sheet selection
sheet_names = list(data.keys())
sheet = st.sidebar.selectbox("Select a sheet:", options=sheet_names)

# Display data for the selected sheet
df_selected = data[sheet]
st.write(f"Data for {sheet}:")
st.dataframe(df_selected)

# If the selected sheet is not 'Summary', plot the charts
if sheet != 'Summary':
    # Chart 1: Long, Short, and Net Position for the past 20 data points
    st.write("Chart 1: Long, Short, and Net Position")
    fig = px.line(df_selected[-20:], x='Date', y=['Long', 'Short', 'Net Position'])
    st.plotly_chart(fig, use_container_width=True)

    # Chart 2: Net Position with its 13-week Moving average
    st.write("Chart 2: Net Position with 13-week Moving average")
    df_selected['13-Week MA'] = df_selected['Net Position'].rolling(window=13).mean()
    fig2 = px.line(df_selected[-20:], x='Date', y='Net Position')
    fig2.add_scatter(x=df_selected['Date'][-20:], y=df_selected['13-Week MA'][-20:], mode='lines', name='13-Week MA')
    st.plotly_chart(fig2, use_container_width=True)

# Mini interactive chart
# Here you would define your mini interactive chart
# This could be another line chart or a different type of chart, such as a bar chart

# Please replace "YOUR_COLUMN_NAME" with the actual column name you want to plot
# st.write("Mini Interactive Chart")
# fig_mini = px.line(df_selected, x='Date', y='YOUR_COLUMN_NAME')
# st.plotly_chart(fig_mini, use_container_width=True)



