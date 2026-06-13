import pandas as pd
import plotly.express as px
import streamlit as st
import requests
from io import BytesIO

URL = "https://www.dropbox.com/scl/fi/0nejigu8olvzhzm179cef/vm_2026_resultater.xlsx?rlkey=tdoi40028u4ve6nvqsow4zurt&st=ja9nbu4m&dl=1"

@st.cache_data(ttl=60)
def load_data():
    r = requests.get(URL)
    return pd.read_excel(BytesIO(r.content), sheet_name="Poeng")

df = load_data()

df["tid"] = pd.to_datetime(df["tid"], format="%d.%m %H:%M")

latest = df.iloc[-1]

ranking_df = latest.drop("tid").reset_index()
ranking_df.columns = ["Deltaker", "Poeng"]

# sikre numerisk
ranking_df["Poeng"] = pd.to_numeric(ranking_df["Poeng"], errors="coerce")

ranking_df = ranking_df.sort_values("Poeng", ascending=False)

ranking_df["Plass"] = range(1, len(ranking_df) + 1)

ranking_df = ranking_df[["Plass", "Deltaker", "Poeng"]]

st.subheader("🏆 Rangering")
st.dataframe(ranking_df, use_container_width=True)

df_long = df.melt(id_vars="tid", var_name="Deltaker", value_name="Poeng")

fig = px.line(df_long, x="tid", y="Poeng", color="Deltaker", markers=True)

st.plotly_chart(fig, use_container_width=True)
