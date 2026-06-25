import pandas as pd
import plotly.graph_objects as go
import streamlit as st
import requests
from io import BytesIO

graph_height = 700

# -------------------------------------------------
# SIDELAYOUT
# -------------------------------------------------
st.set_page_config(layout="wide")
OSLO = "Europe/Oslo"
URL = "https://www.dropbox.com/scl/fi/0nejigu8olvzhzm179cef/vm_2026_resultater.xlsx?rlkey=tdoi40028u4ve6nvqsow4zurt&dl=1"

def excel_tid_til_datetime(series):
    if pd.api.types.is_numeric_dtype(series):
        dt = pd.to_datetime(series, unit="D", origin="1899-12-30", errors="coerce")
    else:
        dt = pd.to_datetime(series, errors="coerce")

    return dt.dt.tz_localize("Europe/Oslo", nonexistent="NaT", ambiguous="NaT")

def to_oslo_datetime(series):
    """
    Standardiserer ALL input til tz-aware Europe/Oslo datetime.
    Støtter:
    - Excel serial numbers
    - strings
    - datetime
    """

    # 1) forsøk vanlig parsing først (strings/datetime)
    dt = pd.to_datetime(series, errors="coerce")

    # 2) hvis dette gir mye NaT → fallback til Excel serial
    if dt.notna().sum() == 0:
        dt = pd.to_datetime(series, unit="D", origin="1899-12-30", errors="coerce")

    # 3) gjør ALT til Oslo tz-aware
    if getattr(dt.dt, "tz", None) is None:
        return dt.dt.tz_localize(OSLO)

    return dt.dt.tz_convert(OSLO)

@st.cache_data(ttl=30)
def load_data():
    r = requests.get(URL, timeout=20)
    r.raise_for_status()

    xls = pd.ExcelFile(BytesIO(r.content))
    sheet_names = xls.sheet_names

    poeng_df = pd.read_excel(xls, sheet_name="Poeng")
    toppscorere_df = pd.read_excel(xls, sheet_name="Toppscorere")

    # Les hendelser fra Poeng-arket (kolonnene Hendelse, Type og tid)
    poeng_raw = pd.read_excel(xls, sheet_name="Poeng")
    poeng_raw = poeng_raw.loc[:, ~poeng_raw.columns.astype(str).str.contains(r"^Unnamed")]
    pcols = [c.strip() for c in poeng_raw.columns.astype(str)]
    poeng_raw.columns = pcols
    h_cols = [c for c in pcols if c.lower() == "hendelse"]
    t_cols = [c for c in pcols if c.lower() == "type"]
    if h_cols:
        rows = poeng_raw[poeng_raw[h_cols[0]].notna()].copy()
        keep = {"tid": pcols[0], "Hendelse": h_cols[0]}
        if t_cols:
            keep["Type"] = t_cols[0]
        hendelser_df = rows[[keep["tid"], keep["Hendelse"]] + ([keep["Type"]] if "Type" in keep else [])].copy()
        hendelser_df.columns = ["tid", "Hendelse"] + (["Type"] if "Type" in keep else [])
    else:
        hendelser_df = None

    if "Neste kamp" in sheet_names:
        neste_kamp_df = pd.read_excel(xls, sheet_name="Neste kamp", header=0)
        neste_kamp_df = neste_kamp_df.dropna(how="all")
    else:
        neste_kamp_df = None

    if "Aktiv kamp" in sheet_names:
        aktiv_kamp_df = pd.read_excel(xls, sheet_name="Aktiv kamp", header=0)
        aktiv_kamp_df = aktiv_kamp_df.dropna(how="all")
    else:
        aktiv_kamp_df = None

    if "Kamptips" in sheet_names:
        kamptips_df = pd.read_excel(xls, sheet_name="Kamptips", header=0)
        kamptips_df = kamptips_df.dropna(how="all")
    else:
        kamptips_df = None

    poeng_df = poeng_df.loc[:, ~poeng_df.columns.astype(str).str.contains(r"^Unnamed")]
    poeng_df = poeng_df.dropna(axis=1, how="all")

    toppscorere_df = toppscorere_df.loc[:, ~toppscorere_df.columns.astype(str).str.contains(r"^Unnamed")]
    toppscorere_df = toppscorere_df.dropna(axis=1, how="all")

    if hendelser_df is not None:
        hendelser_df = hendelser_df.loc[:, ~hendelser_df.columns.astype(str).str.contains(r"^Unnamed")]
        hendelser_df = hendelser_df.dropna(axis=1, how="all")

    return poeng_df, toppscorere_df, hendelser_df, neste_kamp_df, aktiv_kamp_df, kamptips_df, sheet_names


poeng_df, toppscorere_df, hendelser_df, neste_kamp_df, aktiv_kamp_df, kamptips_df, sheet_names = load_data()

# -------------------------------------------------
# KONVERTER TID
# -------------------------------------------------
def _parse_tid_value(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return pd.NaT
    if isinstance(v, (int, float)):
        return pd.Timestamp("1899-12-30") + pd.Timedelta(days=float(v))
    if hasattr(v, "year"):  # datetime object
        return pd.Timestamp(v)
    try:
        return pd.to_datetime(str(v), dayfirst=True)
    except Exception:
        return pd.NaT

poeng_df["tid"] = pd.to_datetime([_parse_tid_value(v) for v in poeng_df["tid"]], utc=False)
if poeng_df["tid"].dt.tz is None:
    poeng_df["tid"] = poeng_df["tid"].dt.tz_localize("Europe/Oslo")
else:
    poeng_df["tid"] = poeng_df["tid"].dt.tz_convert("Europe/Oslo")
poeng_df = poeng_df.dropna(subset=["tid"]).copy()

NON_DELTAKER_COLS = {"tid", "row_id", "Hendelse", "hendelse", "Type", "type", "matchid", "Matchid", "MatchId"}
DELTAKER_COLS = [c for c in poeng_df.columns if c not in NON_DELTAKER_COLS and c != "tid"][:12]
poeng_df = poeng_df[["tid"] + DELTAKER_COLS].copy()
poeng_df["row_id"] = range(len(poeng_df))

# -------------------------------------------------
# LEGG TIL EKSTRA PUNKT MED NÅTID (NORSK TID)
# -------------------------------------------------
now = pd.Timestamp.now(tz="Europe/Oslo").floor("min")
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

cutoff_24h = now - pd.Timedelta(hours=24)
row_24h = poeng_df[poeng_df["tid"] <= cutoff_24h].tail(1)
if not row_24h.empty:
    row_24h = row_24h.iloc[0]
    deltaker_cols_rank = DELTAKER_COLS
    rang_24h = (
        pd.Series({d: pd.to_numeric(row_24h[d], errors="coerce") for d in deltaker_cols_rank})
        .dropna()
        .rank(method="min", ascending=False)
        .astype(int)
    )

    def endring_24h(deltaker):
        n = pd.to_numeric(latest[deltaker], errors="coerce")
        g = pd.to_numeric(row_24h[deltaker], errors="coerce")
        if pd.notna(n) and pd.notna(g):
            diff = int(n - g)
            if diff > 0:
                return f"+{diff}"
            elif diff < 0:
                return str(diff)
        return ""

    def rangendring_24h(deltaker, nå_plass):
        if deltaker not in rang_24h.index:
            return ""
        gammel_plass = rang_24h[deltaker]
        diff = int(nå_plass) - gammel_plass
        if diff < 0:
            return "▲"
        elif diff > 0:
            return "▼"
        return ""

    ranking_df["24t"] = ranking_df["Deltaker"].map(endring_24h)
    ranking_df["Trend"] = ranking_df.apply(lambda r: rangendring_24h(r["Deltaker"], r["Plass"]), axis=1)
else:
    ranking_df["24t"] = ""
    ranking_df["Trend"] = ""

ranking_df = ranking_df[["Plass", "Medalje", "Deltaker", "Poeng", "24t", "Trend"]].head(12)

def trend_html(trend):
    if trend == "▲":
        return "<span style='color:#2a2;'>▲</span>"
    elif trend == "▼":
        return "<span style='color:#c33;'>▼</span>"
    return ""

def parse_endring(val):
    try:
        return int(str(val).replace("+", ""))
    except (ValueError, TypeError):
        return None

endrings_verdier = [parse_endring(r["24t"]) for _, r in ranking_df.iterrows()]
max_endring = max((v for v in endrings_verdier if v is not None), default=None)

rows_html = "\n".join(
    f"<tr>"
    f"<td>{row.Plass} {trend_html(row['Trend'])}</td>"
    f"<td>{row.Medalje}</td>"
    f"<td>{row.Deltaker}</td>"
    f"<td>{int(row.Poeng) if pd.notna(row.Poeng) else ''}</td>"
    f"<td style='color:{'#2a2' if str(row['24t']).startswith('+') else '#c33' if str(row['24t']).startswith('-') else '#888'};{' font-weight:700;' if max_endring is not None and parse_endring(row['24t']) == max_endring and max_endring > 0 else ''}'>{row['24t']}</td>"
    f"</tr>"
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
    font-size: 0.80rem;
    table-layout: fixed;
}}
.ranking-table th,
.ranking-table td {{
    white-space: nowrap;
    padding: 0.15rem 0.30rem;
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
            <th style="width:20%;">Plass</th>
            <th style="width:10%;"></th>
            <th style="width:42%;">Deltaker</th>
            <th style="width:14%;">Poeng</th>
            <th style="width:14%;">24t</th>
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

toppscorere_top3 = toppscorere_df[[col_map["Navn"], col_map["Land"], col_map["Mål"]]].head(10)
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
        hendelser_vis["DatoTid"] = excel_tid_til_datetime(hendelser_vis[tid_col])
        hendelser_vis = hendelser_vis.dropna(subset=["DatoTid", tekst_col])

        cutoff = pd.Timestamp.now(tz="Europe/Oslo") - pd.Timedelta(hours=24)
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
        hendelser_vis["Type"] = hendelser_vis[type_col].astype(str) if type_col is not None else ""
        # BEHOLD DatoTid for hover-match
        hendelser_vis = hendelser_vis[["DatoTid", "Tid", "Type", "Hendelse"]]

        lookup_cols = [tid_col, tekst_col] + ([type_col] if type_col else [])
        hendelser_lookup_df = hendelser_df[lookup_cols].copy()
        hendelser_lookup_df.columns = ["DatoTid", "Hendelse"] + (["Type"] if type_col else [])
        if "Type" not in hendelser_lookup_df.columns:
            hendelser_lookup_df["Type"] = ""
        hendelser_lookup_df["DatoTid"] = excel_tid_til_datetime(hendelser_lookup_df["DatoTid"])
        hendelser_lookup_df = hendelser_lookup_df.dropna(subset=["DatoTid", "Hendelse"]).sort_values("DatoTid").reset_index(drop=True)

# -------------------------------------------------
# HJELPEFUNKSJONER
# -------------------------------------------------
def finn_nærmeste_hendelse(ts, events_df, max_diff="45s"):
    if events_df is None or events_df.empty or pd.isna(ts):
        return ""

    diffs = (events_df["DatoTid"] - ts).abs()
    idx = diffs.idxmin()
    if diffs.loc[idx] <= pd.Timedelta(max_diff):
        row = events_df.loc[idx]
        type_str = str(row.get("Type", "")) if "Type" in events_df.columns else ""
        hendelse_str = str(row["Hendelse"])
        if type_str and type_str != "nan":
            return f"{type_str}<br>{hendelse_str}"
        return hendelse_str
    return ""

def finn_nærmeste_hendelse_rad(ts, events_df, max_diff="45s"):
    if events_df is None or events_df.empty or pd.isna(ts):
        return None

    diffs = (events_df["DatoTid"] - ts).abs()
    idx = diffs.idxmin()

    if diffs.loc[idx] <= pd.Timedelta(max_diff):
        return events_df.loc[idx]
    return None

def poengendring_ved_rad(ts, poeng_df, deltaker_cols):
    after_rows  = poeng_df[poeng_df["tid"] == ts]
    before_rows = poeng_df[poeng_df["tid"] < ts]

    if after_rows.empty or before_rows.empty:
        return ""

    after_row  = after_rows.iloc[0]
    before_row = before_rows.iloc[-1]

    diffs = {}
    for deltaker in deltaker_cols:
        b = pd.to_numeric(before_row[deltaker], errors="coerce")
        a = pd.to_numeric(after_row[deltaker], errors="coerce")
        if pd.notna(b) and pd.notna(a) and a - b != 0:
            diffs.setdefault(int(a - b), []).append(str(deltaker))

    parts = []
    for diff in sorted(diffs.keys(), reverse=True):
        sign = "+" if diff > 0 else ""
        parts.append(f"{sign}{diff} {', '.join(sorted(diffs[diff]))}")

    return "<br>".join(parts)

# -------------------------------------------------
# PLOT
# -------------------------------------------------
poeng_plot = poeng_df.copy()
poeng_plot["tid"] = pd.to_datetime(poeng_plot["tid"], errors="coerce")
poeng_plot = poeng_plot.dropna(subset=["tid"]).copy()

deltaker_cols = DELTAKER_COLS

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

event_texts = []
for ts in poeng_plot["tid"]:
    event_text = finn_nærmeste_hendelse(ts, hendelser_lookup_df, max_diff="45s")
    if event_text:
        changes_text = poengendring_ved_rad(ts, poeng_plot, deltaker_cols)
        if changes_text:
            event_text = f"{event_text}<br>{changes_text}"
    event_texts.append(event_text)

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
            y=label_df["Poeng"] + 1,
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

# -------------------------------------------------
# Y-AKSE: zoom inn til nærmeste 5 poeng under/over
# -------------------------------------------------
visible_df = poeng_plot[
    (poeng_plot["tid"] >= now - pd.Timedelta(hours=24)) &
    (poeng_plot["tid"] <= now)
].copy()

y_values = visible_df[deltaker_cols].apply(pd.to_numeric, errors="coerce").to_numpy().ravel()
y_values = y_values[~pd.isna(y_values)]

if len(y_values) > 0:
    y_min = float(y_values.min())
    y_max = float(y_values.max())

    # NED: ned til nærmeste 5 under minimum
    y_lower = int((y_min-1) // 5) * 5

    # OPP: opp til nærmeste 5 over maksimum
    y_upper = int((-(-(y_max+1) // 5))) * 5  # ceil til nærmeste 5

    if y_lower == y_upper:
        y_lower -= 5
        y_upper += 5
else:
    y_lower = 0
    y_upper = 10

fig.update_layout(
    height=graph_height,
    margin=dict(l=0, r=0, t=0, b=0),
    hovermode="x",
    hoverdistance=20,
    spikedistance=-1,
    legend_title_text="",
    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0),
    hoverlabel=dict(
        bgcolor="rgba(255,255,255,0.75)",
        bordercolor="rgba(0,0,0,0.15)",
        font=dict(color="#222")
    ),
    yaxis=dict(
        range=[y_lower, y_upper]
    )
)


# -------------------------------------------------
# LAYOUT
# -------------------------------------------------
main_col, side_col = st.columns([5.8, 1.8], gap="large")

with main_col:
    st.plotly_chart(fig, use_container_width=True)

    # Neste kamp-boks
    neste_kamp_html = ""
    if neste_kamp_df is not None and not neste_kamp_df.empty:
        cols = neste_kamp_df.columns.tolist()
        hjemmetla_col   = cols[1] if len(cols) > 1 else None
        hjemmeflagg_col = cols[2] if len(cols) > 2 else None
        borteflagg_col  = cols[3] if len(cols) > 3 else None
        bortetla_col    = cols[4] if len(cols) > 4 else None
        tidspunkt_col   = cols[5] if len(cols) > 5 else None

        kamp_linjer = ""
        for _, row in neste_kamp_df.iterrows():
            hjemmetla   = row[hjemmetla_col]   if hjemmetla_col   else ""
            hjemmeflagg = row[hjemmeflagg_col] if hjemmeflagg_col else ""
            borteflagg  = row[borteflagg_col]  if borteflagg_col  else ""
            bortetla    = row[bortetla_col]     if bortetla_col    else ""
            tidspunkt   = row[tidspunkt_col]    if tidspunkt_col   else ""
            match_id    = row[cols[0]]

            if pd.notna(tidspunkt):
                try:
                    tidspunkt_vis = pd.to_datetime(tidspunkt, utc=True).tz_convert("Europe/Oslo").strftime("%d.%m %H:%M")
                except Exception:
                    tidspunkt_vis = str(tidspunkt)
            else:
                tidspunkt_vis = ""

            flagg_h = f'<img src="{hjemmeflagg}" style="height:18px;width:auto;vertical-align:middle;">' if pd.notna(hjemmeflagg) and hjemmeflagg else ""
            flagg_b = f'<img src="{borteflagg}" style="height:18px;width:auto;vertical-align:middle;">' if pd.notna(borteflagg) and borteflagg else ""

            tips_html = ""
            if kamptips_df is not None and pd.notna(match_id):
                try:
                    mid_int = int(float(str(match_id)))
                    kt = kamptips_df.copy()
                    kt["_mid"] = pd.to_numeric(kt.iloc[:, 0], errors="coerce").astype("Int64")
                    tips_row = kt[kt["_mid"] == mid_int]
                    if not tips_row.empty:
                        tip_cols = kamptips_df.columns[1:]
                        tip_lines = "".join(
                            f'<div>{col}: {tips_row.iloc[0][col]}</div>'
                            for col in tip_cols
                            if pd.notna(tips_row.iloc[0][col]) and str(tips_row.iloc[0][col]).strip()
                        )
                        tips_html = f'<div class="neste-tip">{tip_lines}</div>'
                except Exception:
                    pass

            kamp_linjer += (
                f'<div class="neste-kamp-rad">'
                f'<div style="display:flex;align-items:center;gap:4px;margin-bottom:2px;">'
                f'<span style="flex:1;text-align:right;font-weight:600;">{hjemmetla} {flagg_h}</span>'
                f'<span style="font-weight:600;">–</span>'
                f'<span style="flex:1;text-align:left;font-weight:600;">{flagg_b} {bortetla}</span>'
                f'<span style="color:#444;font-size:0.82rem;white-space:nowrap;margin-left:6px;">{tidspunkt_vis}</span>'
                f'</div>'
                f'{tips_html}'
                f'</div>'
            )

        neste_kamp_html = (
            f'<div style="width:100%;padding:8px 10px;border:1px solid rgba(49,51,63,0.15);border-radius:10px;background:#f0f4ff;font-size:0.90rem;box-sizing:border-box;">'
            f'<div style="font-size:0.78rem;color:#666;margin-bottom:4px;">{"Neste kamper" if len(neste_kamp_df) > 1 else "Neste kamp"}</div>'
            f'{kamp_linjer}'
            f'</div>'
        )

    if hendelser_vis is None:
        st.info(f"Fant ikke brukbare kolonner i arket 'Hendelser'. Tilgjengelige ark: {sheet_names}")
    else:
        hendelser_html = ""
        if hendelser_vis is not None and not hendelser_vis.empty:
            hendelser_html = "".join(
                f'<div style="display:inline-block;vertical-align:top;min-width:180px;max-width:250px;margin-right:10px;padding:8px 10px;border:1px solid rgba(49,51,63,0.15);border-radius:10px;background:#fafafa;font-size:0.90rem;">'
                f'<div style="font-weight:600;margin-bottom:3px;">{row.Tid}</div>'
                f'<div style="font-size:0.82rem;color:#666;margin-bottom:3px;">{row.Type}</div>'
                f'<div style="line-height:1.3;">{row.Hendelse}</div>'
                f'</div>'
                for _, row in hendelser_vis.iterrows()
            )

        st.markdown(
            '<div style="display:flex;flex-wrap:nowrap;gap:10px;overflow-x:auto;padding-bottom:6px;">'
            + hendelser_html +
            '</div>',
            unsafe_allow_html=True
        )

with side_col:
    st.markdown(ranking_html, unsafe_allow_html=True)

    st.markdown("""<style>
.aktiv-kamp-rad{position:relative;}
.aktiv-tip{display:none;position:absolute;right:0;top:110%;background:#222;color:#fff;padding:6px 10px;border-radius:6px;font-size:0.78rem;font-weight:400;white-space:nowrap;text-align:right;z-index:9999;min-width:160px;}
.aktiv-kamp-rad:hover .aktiv-tip{display:block;}
.neste-kamp-rad{position:relative;cursor:default;}
.neste-tip{display:none;position:absolute;left:0;top:100%;background:#222;color:#fff;padding:6px 10px;border-radius:6px;font-size:0.78rem;font-weight:400;white-space:nowrap;text-align:right;z-index:9999;min-width:160px;}
.neste-kamp-rad:hover .neste-tip{display:block;}
@keyframes rec-blink{0%,100%{opacity:1;}50%{opacity:0.15;}}
.rec-dot{display:inline-block;width:8px;height:8px;border-radius:50%;background:#e00;margin-right:5px;vertical-align:middle;animation:rec-blink 1.2s ease-in-out infinite;}
</style>""", unsafe_allow_html=True)

    aktiv_kamp_html = ""
    if aktiv_kamp_df is not None and not aktiv_kamp_df.empty:
        cols = aktiv_kamp_df.columns.tolist()
        kamp_linjer = ""
        for _, row in aktiv_kamp_df.iterrows():
            hjemmetla   = row[cols[1]] if len(cols) > 1 else ""
            hjemmeflagg = row[cols[2]] if len(cols) > 2 else ""
            hjemmemål   = row[cols[3]] if len(cols) > 3 else ""
            bortemål    = row[cols[4]] if len(cols) > 4 else ""
            borteflagg  = row[cols[5]] if len(cols) > 5 else ""
            bortetla    = row[cols[6]] if len(cols) > 6 else ""
            flagg_h = f'<img src="{hjemmeflagg}" style="height:18px;width:auto;vertical-align:middle;">' if pd.notna(hjemmeflagg) and hjemmeflagg else ""
            flagg_b = f'<img src="{borteflagg}" style="height:18px;width:auto;vertical-align:middle;">' if pd.notna(borteflagg) and borteflagg else ""
            hm = int(hjemmemål) if pd.notna(hjemmemål) else 0
            bm = int(bortemål)  if pd.notna(bortemål)  else 0
            status_val = row[cols[7]] if len(cols) > 7 else ""
            status_str = str(status_val).strip() if pd.notna(status_val) else ""
            if status_str == "IN_PLAY":
                status_label = "Pågår"
            elif status_str == "PAUSED":
                status_label = "Pause"
            else:
                status_label = status_str

            match_id = row[cols[0]] if cols else None
            tips_linjer = ""
            if kamptips_df is not None and match_id is not None:
                try:
                    mid_int = int(float(str(match_id)))
                    kt = kamptips_df.copy()
                    kt["_mid"] = pd.to_numeric(kt.iloc[:, 0], errors="coerce").astype("Int64")
                    tips_row = kt[kt["_mid"] == mid_int]
                    if not tips_row.empty:
                        tip_cols = kamptips_df.columns[1:]
                        tips_linjer = "\n".join(
                            f"{col}: {tips_row.iloc[0][col]}"
                            for col in tip_cols
                            if pd.notna(tips_row.iloc[0][col]) and str(tips_row.iloc[0][col]).strip()
                        )
                except Exception:
                    pass

            header_line = f"{hjemmetla} {hm} – {bm} {bortetla} · {status_label}"
            tip_html = ""
            if tips_linjer:
                tip_rows = "".join(
                    f'<div>{line}</div>'
                    for line in (header_line + "\n\nTips:\n" + tips_linjer).split("\n")
                )
                tip_html = f'<div class="aktiv-tip">{tip_rows}</div>'

            kamp_linjer += (
                f'<div class="aktiv-kamp-rad">'
                f'<div style="display:flex;align-items:center;gap:4px;font-weight:600;cursor:default;">'
                f'<span style="flex:1;text-align:right;">{hjemmetla} {flagg_h}</span>'
                f'<span>{hm} – {bm}</span>'
                f'<span style="flex:1;text-align:left;">{flagg_b} {bortetla}</span>'
                f'</div>'
                f'{tip_html}'
                f'</div>'
            )
        aktiv_kamp_html = (
            f'<div style="width:100%;padding:8px 10px;border:1px solid rgba(49,51,63,0.15);border-radius:10px;background:#fff4e0;font-size:0.90rem;box-sizing:border-box;">'
            f'<div style="font-size:0.78rem;color:#666;margin-bottom:4px;"><span class="rec-dot"></span>Aktiv kamp</div>'
            f'{kamp_linjer}'
            f'</div>'
        )

    if aktiv_kamp_html:
        st.markdown(aktiv_kamp_html, unsafe_allow_html=True)
    elif neste_kamp_html:
        st.markdown(neste_kamp_html, unsafe_allow_html=True)

    st.markdown("<div style='font-size:0.82rem;font-weight:600;margin-top:4px;margin-bottom:2px;'>⚽ Toppscorere</div>", unsafe_allow_html=True)
    st.dataframe(toppscorere_top3, use_container_width=True, hide_index=True, height=35 + 35 * len(toppscorere_top3))
