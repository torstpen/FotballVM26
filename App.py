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
    r = requests.get(URL, timeout=20)
    r.raise_for_status()

    poeng_df = pd.read_excel(BytesIO(r.content), sheet_name="Poeng")
    toppscorere_df = pd.read_excel(BytesIO(r.content), sheet_name="Toppscorere")

    # Fjern tomme / unnødvendige kolonner
    poeng_df = poeng_df.loc[:, ~poeng_df.columns.astype(str).str.contains(r"^Unnamed")]
    poeng_df = poeng_df.dropna(axis=1, how="all")

    toppscorere_df = toppscorere_df.loc[:, ~toppscorere_df.columns.astype(str).str.contains(r"^Unnamed")]
    toppscorere_df = toppscorere_df.dropna(axis=1, how="all")

    return poeng_df, toppscorere_df


poeng_df, toppscorere_df = load_data()

# -------------------------------------------------
# KONVERTER TID (Excel serial -> datetime)
# -------------------------------------------------
poeng_df["tid"] = pd.to_datetime(poeng_df["tid"], errors="coerce", unit="D", origin="1899-12-30")
poeng_df = poeng_df.dropna(subset=["tid"])
poeng_df["row_id"] = range(len(poeng_df))

# -------------------------------------------------
# RANGERING (siste rad) MED DELTE PLASSER + MEDALJER
# -------------------------------------------------
latest = poeng_df.iloc[-1]

ranking_df = latest.drop(["tid", "row_id"]).reset_index()
ranking_df.columns = ["Deltaker", "Poeng"]

ranking_df["Poeng"] = pd.to_numeric(ranking_df["Poeng"], errors="coerce")
ranking_df = ranking_df.dropna(subset=["Poeng"])
ranking_df = ranking_df.sort_values(["Poeng", "Deltaker"], ascending=[False, True]).reset_index(drop=True)

# Delte plasser
ranking_df["Plass"] = ranking_df["Poeng"].rank(method="min", ascending=False).astype("Int64")

# Medalje-kolonne
medals = {1: "🥇", 2: "🥈", 3: "🥉"}
ranking_df["Medalje"] = ranking_df["Plass"].map(medals).fillna("")

ranking_df = ranking_df[["Plass", "Medalje", "Deltaker", "Poeng"]].head(12)

rows_html = "\n".join(
    f"<tr><td>{row.Plass}</td><td>{row.Medalje}</td><td>{row.Deltaker}</td><td>{int(row.Poeng) if pd.notna(row.Poeng) else ''}</td></tr>"
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
            <th></th>
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
# TOPPSCORERE (TOPP 3)
# -------------------------------------------------
toppscorere_df.columns = toppscorere_df.columns.astype(str).str.strip()

# Tilpass hvis kolonnenavnene avviker
# Forventer: Plassering, Navn, Land, Mål
# Rydd kolonnenavn
toppscorere_df.columns = toppscorere_df.columns.astype(str).str.strip()

# Finn kolonner mer fleksibelt
col_map = {}
for col in toppscorere_df.columns:
    low = col.lower()
    if "plass" in low or "rank" in low:
        col_map["Plassering"] = col
    elif "navn" in low or "player" in low:
        col_map["Navn"] = col
    elif "land" in low or "country" in low:
        col_map["Land"] = col
    elif "mål" in low or "maal" in low or "goals" in low:
        col_map["Mål"] = col

missing = [k for k in ["Plassering", "Navn", "Land", "Mål"] if k not in col_map]

if missing:
    st.error(f"Mangler kolonner i arket Toppscorere: {missing}")
    st.write("Fant disse kolonnene:", toppscorere_df.columns.tolist())
    st.stop()

toppscorere_top3 = toppscorere_df[[col_map["Plassering"], col_map["Navn"], col_map["Land"], col_map["Mål"]]].head(3)
toppscorere_top3.columns = ["Plassering", "Navn", "Land", "Mål"]

# -------------------------------------------------
# LONG FORMAT FOR PLOT
# -------------------------------------------------
df_long = poeng_df.melt(
    id_vars=["tid", "row_id"],
    var_name="Deltaker",
    value_name="Poeng"
).sort_values("row_id")

df_long["Poeng"] = pd.to_numeric(df_long["Poeng"], errors="coerce")

# -------------------------------------------------
# PLOT UTEN MARKØRER
# -------------------------------------------------
fig = px.line(
    df_long,
    x="tid",
    y="Poeng",
    color="Deltaker"
)

fig.update_traces(line_shape="hv")
fig.update_layout(
    hovermode="x unified",
    height=620,
    margin=dict(l=20, r=20, t=30, b=20),
    legend_title_text=""
)

# -------------------------------------------------
# LAYOUT
# -------------------------------------------------
col1, col2 = st.columns([5, 1], gap="small")

with col1:
    st.subheader("📈 Poenggraf")
    st.plotly_chart(fig, use_container_width=True)

with col2:
    st.subheader("🏆 Rangering")
    st.markdown(ranking_html, unsafe_allow_html=True)

    st.subheader("⚽ Toppscorere")
    st.dataframe(
        toppscorere_top3,
        use_container_width=True,
        hide_index=True
    )
