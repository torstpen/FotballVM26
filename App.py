import pandas as pd
import plotly.express as px
import streamlit as st
import requests
from io import BytesIO

# -------------------------------------------------
# SIDELAYOUT
# -------------------------------------------------
st.set_page_config(layout="wide")

URL = "https://www.dropbox.com/scl/fi/0nejigu8olvzhzm179cef/vm_2026_resultater.xlsx?rlkey=tdoi40028u4ve6nvqsow4zurt&dl=1"


@st.cache_data(ttl=60)
def load_data():
    r = requests.get(URL)
    return pd.read_excel(BytesIO(r.content), sheet_name="Poeng")


df = load_data()

# -------------------------------------------------
# KONVERTER TID (Excel serial -> datetime)
# -------------------------------------------------
df["tid"] = pd.to_datetime(df["tid"], errors="coerce", unit="D", origin="1899-12-30")
df = df.dropna(subset=["tid"])
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
ranking_df = ranking_df[["Plass", "Deltaker", "Poeng"]].head(12)

rows_html = "\n".join(
    f"<tr><td>{row.Plass}</td><td>{row.Deltaker}</td><td>{int(row.Poeng) if pd.notna(row.Poeng) else ''}</td></tr>"
    for _, row in ranking_df.iterrows()
)

ranking_html = f"""
<style>
.ranking-wrap {{
    display: inline-block;
}}
.ranking-table {{
    width: auto;
    border-collapse: collapse;
    font-size: 0.95rem;
}}
.ranking-table th,
.ranking-table td {{
    white-space: nowrap;
    padding: 0.35rem 0.6rem;
    text-align: left;
    border-bottom: 1px solid rgba(49, 51, 63, 0.15);
}}
.ranking-table th {{
    font-weight: 600;
    background-color: rgba(250, 250, 250, 0.95);
}}
.ranking-table tr:nth-child(even) td {{
    background-color: rgba(0, 0, 0, 0.02);
}}
</style>

<div class="ranking-wrap">
<table class="ranking-table">
    <thead>
        <tr>
            <th>Plass</th>
            <th>Deltaker</th>
            <th>Poeng</th>
        </tr>
    </thead>
    <tbody>
        {rows_html}
    </tbody>
</table>
</div>
"""

# -------------------------------------------------
# LONG FORMAT FOR PLOT
# -------------------------------------------------
df_long = df.melt(
    id_vars=["tid", "row_id"],
    var_name="Deltaker",
    value_name="Poeng"
).sort_values("row_id")

# -------------------------------------------------
# PLOT
# -------------------------------------------------
fig = px.line(
    df_long,
    x="tid",
    y="Poeng",
    color="Deltaker",
    markers=True
)

fig.update_traces(line_shape="hv")
fig.update_layout(
    hovermode="x unified",
    height=620,
    margin=dict(l=20, r=20, t=30, b=20),
    legend_title_text=""
)

# -------------------------------------------------
# LAYOUT: GRAF VELDIG STOR, TABELL SMAL
# -------------------------------------------------
col1, col2 = st.columns([5, 1], gap="small")

with col1:
    st.subheader("📈 Poenggraf")
    st.plotly_chart(fig, use_container_width=True)

with col2:
    st.subheader("🏆 Rangering")
    st.markdown(ranking_html, unsafe_allow_html=True)
