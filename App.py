import pandas as pd
import plotly.graph_objects as go
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

    xls = pd.ExcelFile(BytesIO(r.content))
    sheet_names = xls.sheet_names

    poeng_df = pd.read_excel(xls, sheet_name="Poeng")
    toppscorere_df = pd.read_excel(xls, sheet_name="Toppscorere")

    if "Hendelser" in sheet_names:
        hendelser_df = pd.read_excel(xls, sheet_name="Hendelser")
    else:
        hendelser_df = None

    poeng_df = poeng_df.loc[:, ~poeng_df.columns.astype(str).str.contains(r"^Unnamed")]
    poeng_df = poeng_df.dropna(axis=1, how="all")

    toppscorere_df = toppscorere_df.loc[:, ~toppscorere_df.columns.astype(str).str.contains(r"^Unnamed")]
    toppscorere_df = toppscorere_df.dropna(axis=1, how="all")

    if hendelser_df is not None:
        hendelser_df = hendelser_df.loc[:, ~hendelser_df.columns.astype(str).str.contains(r"^Unnamed")]
        hendelser_df = hendelser_df.dropna(axis=1, how="all")

    return poeng_df, toppscorere_df, hendelser_df, sheet_names


poeng_df, toppscorere_df, hendelser_df, sheet_names = load_data()

# -------------------------------------------------
# KONVERTER TID
# -------------------------------------------------
poeng_df["tid"] = pd.to_datetime(poeng_df["tid"], errors="coerce", unit="D", origin="1899-12-30")
poeng_df = poeng_df.dropna(subset=["tid"]).copy()
poeng_df["row_id"] = range(len(poeng_df))

# -------------------------------------------------
# LEGG TIL EKSTRA PUNKT MED NÅTID (NORSK TID)
# -------------------------------------------------
now = (pd.Timestamp.now(tz="Europe/Oslo") + pd.Timedelta(hours=2)).floor("min").tz_convert(None)
latest_row = poeng_df.iloc[-1].copy()
latest_row["tid"] = now
latest_row["row_id"] = poeng_df["row_id"].max() + 1

poeng_df = pd.concat([poeng_df, pd.DataFrame([latest_row])], ignore_index=True)

# -------------------------------------------------
# RANGERING
# -------------------------------------------------
latest = poeng_df.iloc[-2]
ranking_df = latest.drop(["tid", "row_id"]).reset_index()
ranking_df.columns = ["Deltaker", "Poeng"]

ranking_df["Poeng"] = pd.to_numeric(ranking_df["Poeng"], errors="coerce")
ranking_df = ranking_df.dropna(subset=["Poeng"])
ranking_df = ranking_df.sort_values(["Poeng", "Deltaker"], ascending=[False, True]).reset_index(drop=True)

ranking_df["Plass"] = ranking_df["Poeng"].rank(method="min", ascending=False).astype("Int64")

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
    width: 100%;
}}
.ranking-table {{
    width: 100%;
    border-collapse: collapse;
    font-size: 0.88rem;
    table-layout: fixed;
}}
.ranking-table th,
.ranking-table td {{
    white-space: nowrap;
    padding: 0.28rem 0.35rem;
    text-align: left;
    border-bottom: 1px solid rgba(49, 51, 63, 0.14);
    overflow: hidden;
    text-overflow: ellipsis;
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
            <th style="width:14%;">Plass</th>
            <th style="width:12%;"></th>
            <th style="width:52%;">Deltaker</th>
            <th style="width:22%;">Poeng</th>
        </tr>
    </thead>
    <tbody>
        {rows_html}
    </tbody>
</table>
</div>
"""

# -------------------------------------------------
# TOPPSCORERE
# -------------------------------------------------
toppscorere_df.columns = toppscorere_df.columns.astype(str).str.strip()

col_map = {}
for col in toppscorere_df.columns:
    low = col.lower()
    if "spiller" in low or "navn" in low or "player" in low:
        col_map["Navn"] = col
    elif "lag" in low or "land" in low or "team" in low or "country" in low:
        col_map["Land"] = col
    elif "mål" in low or "maal" in low or "goals" in low:
        col_map["Mål"] = col

missing = [k for k in ["Navn", "Land", "Mål"] if k not in col_map]

if missing:
    st.error(f"Mangler kolonner i arket Toppscorere: {missing}")
    st.write("Fant disse kolonnene:", toppscorere_df.columns.tolist())
    st.stop()

toppscorere_top3 = toppscorere_df[[col_map["Navn"], col_map["Land"], col_map["Mål"]]].head(3)
toppscorere_top3.columns = ["Navn", "Land", "Mål"]

# -------------------------------------------------
# HENDELSER
# -------------------------------------------------
hendelser_vis = None
hendelser_lookup_df = None

if hendelser_df is not None:
    hendelser_df.columns = hendelser_df.columns.astype(str).str.strip()

    tid_col = None
    tekst_col = None
    type_col = None

    for col in hendelser_df.columns:
        low = col.lower()
        if "tid" in low or "time" in low or "dato" in low:
            tid_col = col
        elif "hend" in low or "event" in low or "beskjed" in low or "tekst" in low:
            tekst_col = col
        elif "type" in low:
            type_col = col

    if tid_col is None and len(hendelser_df.columns) >= 1:
        tid_col = hendelser_df.columns[0]
    if tekst_col is None and len(hendelser_df.columns) >= 2:
        tekst_col = hendelser_df.columns[1]

    if tid_col is not None and tekst_col is not None:
        cols = [tid_col, tekst_col]
        if type_col is not None:
            cols.append(type_col)

        hendelser_vis = hendelser_df[cols].copy()
        hendelser_vis["DatoTid"] = pd.to_datetime(hendelser_vis[tid_col], errors="coerce")
        hendelser_vis = hendelser_vis.dropna(subset=["DatoTid", tekst_col])

        cutoff = pd.Timestamp.now() - pd.Timedelta(hours=24)
        siste_24t = hendelser_vis[hendelser_vis["DatoTid"] >= cutoff]

        if len(siste_24t) < 20:
            mangler = 20 - len(siste_24t)
            eldre = hendelser_vis[hendelser_vis["DatoTid"] < cutoff].head(mangler)
            hendelser_vis = pd.concat([siste_24t, eldre], ignore_index=True)
        else:
            hendelser_vis = siste_24t

        hendelser_vis = hendelser_vis.sort_values("DatoTid", ascending=False)

        hendelser_vis["Tid"] = hendelser_vis["DatoTid"].dt.strftime("%d.%m %H:%M")
        hendelser_vis["Hendelse"] = hendelser_vis[tekst_col].astype(str)

        if type_col is not None:
            hendelser_vis["Type"] = hendelser_vis[type_col].astype(str)
        else:
            hendelser_vis["Type"] = ""

        hendelser_vis = hendelser_vis[["Tid", "Type", "Hendelse"]]

        hendelser_lookup_df = hendelser_df[[tid_col, tekst_col]].copy()
        hendelser_lookup_df.columns = ["DatoTid", "Hendelse"]
        hendelser_lookup_df["DatoTid"] = pd.to_datetime(hendelser_lookup_df["DatoTid"], errors="coerce")
        hendelser_lookup_df = hendelser_lookup_df.dropna(subset=["DatoTid", "Hendelse"]).sort_values("DatoTid").reset_index(drop=True)

# -------------------------------------------------
# FUNKSJON FOR NÆRMESTE HENDELSE
# -------------------------------------------------
def finn_nærmeste_hendelse(ts, events_df, max_diff="45s"):
    if events_df is None or events_df.empty or pd.isna(ts):
        return ""

    diffs = (events_df["DatoTid"] - ts).abs()
    idx = diffs.idxmin()

    if diffs.loc[idx] <= pd.Timedelta(max_diff):
        return str(events_df.loc[idx, "Hendelse"])
    return ""

# -------------------------------------------------
# PLOT
# -------------------------------------------------
poeng_plot = poeng_df.copy()
poeng_plot["tid"] = pd.to_datetime(poeng_plot["tid"], errors="coerce")
poeng_plot = poeng_plot.dropna(subset=["tid"]).copy()

deltaker_cols = [c for c in poeng_plot.columns if c not in ["tid", "row_id"]]

fig = go.Figure()

for deltaker in deltaker_cols:
    y = pd.to_numeric(poeng_plot[deltaker], errors="coerce")

    fig.add_trace(
        go.Scatter(
            x=poeng_plot["tid"],
            y=y,
            mode="lines",
            name=str(deltaker),
            line_shape="hv",
            hoverinfo="skip"
        )
    )

event_texts = [
    finn_nærmeste_hendelse(ts, hendelser_lookup_df, max_diff="45s")
    for ts in poeng_plot["tid"]
]

hover_y = pd.to_numeric(poeng_plot[deltaker_cols[0]], errors="coerce")

fig.add_trace(
    go.Scatter(
        x=poeng_plot["tid"],
        y=hover_y,
        mode="markers",
        marker=dict(size=18, opacity=0.001),
        showlegend=False,
        customdata=event_texts,
        hovertemplate=(
            "<b>%{x|%d.%m %H:%M:%S}</b><br>"
            "%{customdata}<extra></extra>"
        )
    )
)

last_row = poeng_plot.iloc[-1]
last_points = []

for deltaker in deltaker_cols:
    poeng = pd.to_numeric(pd.Series([last_row[deltaker]]), errors="coerce").iloc[0]
    if pd.notna(poeng):
        last_points.append({"Deltaker": str(deltaker), "Poeng": poeng})

last_points = pd.DataFrame(last_points)

label_df = (
    last_points.groupby("Poeng")["Deltaker"]
    .apply(lambda s: ", ".join(s.sort_values().astype(str)))
    .reset_index()
)

if not label_df.empty:
    fig.add_trace(
        go.Scatter(
            x=[now] * len(label_df),
            y=label_df["Poeng"] + 0.3,
            mode="text",
            text=label_df["Deltaker"],
            textposition="middle left",
            showlegend=False,
            hoverinfo="skip"
        )
    )

fig.update_xaxes(
    range=[now - pd.Timedelta(hours=24), now],
    showspikes=True,
    spikemode="across",
    spikesnap="data",
    spikecolor="rgba(80,80,80,0.7)",
    spikethickness=1.2
)

fig.update_layout(
    hovermode="x",
    hoverdistance=20,
    spikedistance=-1,
    height=560,
    margin=dict(l=10, r=20, t=20, b=10),
    legend_title_text="",
    legend=dict(
        orientation="h",
        yanchor="bottom",
        y=1.02,
        xanchor="left",
        x=0
    )
)

# -------------------------------------------------
# LAYOUT
# -------------------------------------------------
main_col, side_col = st.columns([5.8, 1.8], gap="large")

with main_col:
    st.subheader("📈 Poenggraf")
    st.plotly_chart(fig, use_container_width=True)

    st.subheader("📝 Hendelser")
    if hendelser_vis is None:
        st.info(f"Fant ikke brukbare kolonner i arket 'Hendelser'. Tilgjengelige ark: {sheet_names}")
    elif hendelser_vis.empty:
        st.info("Ingen hendelser funnet i arket.")
    else:
        html_items = "".join(
            f"""
            <div style="
                display:inline-block;
                vertical-align:top;
                min-width:180px;
                max-width:250px;
                margin-right:10px;
                padding:8px 10px;
                border:1px solid rgba(49,51,63,0.15);
                border-radius:10px;
                background:#fafafa;
                font-size:0.90rem;
            ">
                <div style="font-weight:600; margin-bottom:3px;">{row.Tid}</div>
                <div style="font-size:0.82rem; color:#666; margin-bottom:3px;">{row.Type}</div>
                <div style="line-height:1.3;">{row.Hendelse}</div>
            </div>
            """
            for _, row in hendelser_vis.iterrows()
        )

        st.markdown(
            f"""
            <div style="display:flex; flex-wrap:nowrap; overflow-x:auto; padding-bottom:6px;">
                {html_items}
            </div>
            """,
            unsafe_allow_html=True
        )

with side_col:
    st.subheader("🏆 Rangering")
    st.markdown(ranking_html, unsafe_allow_html=True)

    st.subheader("⚽ Toppscorere")
    st.dataframe(
        toppscorere_top3,
        use_container_width=True,
        hide_index=True
    )
