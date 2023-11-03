import pandas as pd
import plotly.express as px
import streamlit as st

import warnings
warnings.filterwarnings('ignore')

st.set_page_config(page_title = 'Dashborard for Fashom',
                  page_icon = ":bar_chart:",
                  layout='wide')

df = pd.read_csv("/Users/kensong/Desktop/Data_for_Bangzhuo.csv")



#Sidebar
st.sidebar.header("Please Filter Here:")
stylist = st.sidebar.multiselect(
    "Select the Stylist:",
    options = df.Stylist.unique(),
    default = df.Stylist.unique()
)


State = st.sidebar.multiselect(
    "Select the State:",
    options = df.State.unique(),
    default = df.State.unique()
)

Age = st.sidebar.multiselect(
    "Select the Age:",
    options = df.Age.unique(),
    default = df.Age.unique()
)

df_selection =df.query(
    "Stylist == @stylist & State ==@State & Age == @Age"
)

st.dataframe(df_selection)


#shirt size
shirt_size = (
    df_selection.groupby('Shirt Size').size().reset_index(name = 'count')
)

fig_shirt_size = px.bar(
    shirt_size,
    x = "Shirt Size",
    y= 'count',
    title= 'Shirt Size'
)

st.plotly_chart(fig_shirt_size)


#pants size
pant_size = (
    df_selection.groupby('Pants Size').size().reset_index(name = 'count')
)

fig_pant_size = px.bar(
    pant_size,
    x = "Pants Size",
    y= 'count',
    title= 'Pant Size'
)

fig_pant_size.update_layout(
    yaxis=dict(
        range=[0, 1200],  # Adjust this range to fit your data
    )
)







st.plotly_chart(fig_pant_size)