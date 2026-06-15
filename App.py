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

# -------------------------------------------------
# KONVERTER TID (Excel serial -> datetime)
# -------------------------------------------------
df["tid"] = pd.to_datetime(df["tid"], errors="coerce", unit="D", origin="1899-12-30")

# fjern ugyldige rader
df = df.dropna(subset=["tid"])

# behold Excel-rekkefølge (viktig for stabil plot)
df["row_id"] = range(len(df))

# -------------------------------------------------
# RANGERING (siste rad)
# -------------------------------------------------
latest = df.iloc[-1]

ranking_df = latest.drop(["tid", "row_id"]).reset_index()
ranking_df.columns = ["Deltaker", "Poeng"]

ranking_df["Poeng"] = pd.to_numeric(ranking_df["Poeng"], errors="coerce")
ranking_df = ranking_df.sort_values("Poeng", ascending=False).reset_index(drop=True)
ranking_df["Plass"] = range(1, len(ranking_df) + 1)
ranking_df = ranking_df[["Plass", "Deltaker", "Poeng"]]

st.subheader("🏆 Rangering")

st.dataframe(
    ranking_df,
    use_container_width=False,   # ikke strekk ut unødvendig
    hide_index=True,             # skjuler indeksen
    height=420,                  # nok for 12 rader
    column_config={
        "Plass": st.column_config.NumberColumn(
            "Plass",
            width="small",
            format="%d",
        ),
        "Deltaker": st.column_config.TextColumn(
            "Deltaker",
            width="large",
        ),
        "Poeng": st.column_config.NumberColumn(
            "Poeng",
            width="small",
            format="%.0f",
        ),
    },
)

# -------------------------------------------------
# LONG FORMAT FOR PLOT
# -------------------------------------------------
df_long = df.melt(
    id_vars=["tid", "row_id"],
    var_name="Deltaker",
    value_name="Poeng"
)

df_long = df_long.sort_values("row_id")

# -------------------------------------------------
# STEP CHART
# -------------------------------------------------
fig = px.line(
    df_long,
    x="tid",
    y="Poeng",
    color="Deltaker",
    markers=True
)

fig.update_traces(line_shape="hv")  # 🔥 STEP CHART

fig.update_layout(
    hovermode="x unified"
)

st.plotly_chart(fig, use_container_width=True)
