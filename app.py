import re
from pathlib import Path
from typing import Optional, Tuple, List

import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import streamlit as st

st.set_page_config(page_title="IB Macro Futures Event Dashboard", layout="wide")

DEFAULT_MARKET_FILE = "YIBc_Trades_30min_Event_Matches(3).xlsx"
DEFAULT_MACRO_FILE = "Australia_Economic_Indicators_2022_2024(1).xlsx"

IMPACT_ORDER = ["High", "Medium", "Low"]
IMPACT_COLORS = {"High": "red", "Medium": "orange", "Low": "blue", "Unknown": "gray"}

# -----------------------------
# Helpers
# -----------------------------

def clean_col(c):
    if c is None:
        return ""
    return str(c).strip()


def find_col(df: pd.DataFrame, candidates: List[str]) -> Optional[str]:
    lower_map = {str(c).strip().lower(): c for c in df.columns}
    for cand in candidates:
        if cand.lower() in lower_map:
            return lower_map[cand.lower()]
    # loose contains fallback
    for c in df.columns:
        c_low = str(c).strip().lower()
        for cand in candidates:
            if cand.lower() in c_low:
                return c
    return None


def parse_datetime_series(s: pd.Series) -> pd.Series:
    return pd.to_datetime(s, errors="coerce")


def load_excel(uploaded_file, default_path: str) -> pd.ExcelFile:
    if uploaded_file is not None:
        return pd.ExcelFile(uploaded_file)
    path = Path(default_path)
    if not path.exists():
        st.error(f"Default file not found: {default_path}. Upload it from the sidebar.")
        st.stop()
    return pd.ExcelFile(path)


def read_sheet(xls: pd.ExcelFile, sheet_name: str) -> pd.DataFrame:
    df = pd.read_excel(xls, sheet_name=sheet_name)
    df.columns = [clean_col(c) for c in df.columns]
    df = df.loc[:, [c for c in df.columns if c != ""]]
    return df


def normalize_market_df(df: pd.DataFrame) -> pd.DataFrame:
    dt_col = find_col(df, ["AUS_Local_DateTime", "Trade Date-Time", "UTC_DateTime_Clean", "Date-Time", "Start"])
    if not dt_col:
        raise ValueError("Could not find a datetime column. Expected AUS_Local_DateTime, Trade Date-Time, UTC_DateTime_Clean, or Date-Time.")

    col_open = find_col(df, ["Open"])
    col_high = find_col(df, ["High"])
    col_low = find_col(df, ["Low"])
    col_close = find_col(df, ["Last", "Close"])
    col_vol = find_col(df, ["Volume"])

    required = [col_open, col_high, col_low, col_close]
    if any(c is None for c in required):
        raise ValueError("Could not find OHLC columns. Expected Open, High, Low, Last/Close.")

    out = pd.DataFrame()
    out["datetime"] = parse_datetime_series(df[dt_col])
    out["open"] = pd.to_numeric(df[col_open], errors="coerce")
    out["high"] = pd.to_numeric(df[col_high], errors="coerce")
    out["low"] = pd.to_numeric(df[col_low], errors="coerce")
    out["close"] = pd.to_numeric(df[col_close], errors="coerce")
    out["volume"] = pd.to_numeric(df[col_vol], errors="coerce") if col_vol else 0

    ric_col = find_col(df, ["#RIC", "Alias Underlying RIC"])
    type_col = find_col(df, ["Type"])
    out["ric"] = df[ric_col].astype(str) if ric_col else ""
    out["type"] = df[type_col].astype(str) if type_col else ""

    event_col = find_col(df, ["Matched Event(s)", "Matched_Event(s)", "Event", "Name"])
    impact_col = find_col(df, ["Matched_Impact", "Impact"])
    nearest_col = find_col(df, ["Nearest Event Time", "Nearest_Event_Time"])
    min_col = find_col(df, ["Minutes From Nearest Event", "Minutes_From_Nearest_Event"])

    out["event"] = df[event_col].fillna("").astype(str) if event_col else ""
    out["impact"] = df[impact_col].fillna("Unknown").astype(str) if impact_col else "Unknown"
    out["nearest_event_time"] = parse_datetime_series(df[nearest_col]) if nearest_col else pd.NaT
    out["minutes_from_event"] = pd.to_numeric(df[min_col], errors="coerce") if min_col else pd.NA

    out = out.dropna(subset=["datetime", "open", "high", "low", "close"]).sort_values("datetime")
    out = out.drop_duplicates(subset=["datetime", "open", "high", "low", "close", "volume"], keep="first")
    return out


def normalize_calendar_df(df: pd.DataFrame) -> pd.DataFrame:
    start_col = find_col(df, ["Start", "Nearest Event Time", "Date-Time", "AUS_Local_DateTime"])
    name_col = find_col(df, ["Name", "Matched Event(s)", "Event"])
    impact_col = find_col(df, ["Impact", "Matched_Impact"])
    curr_col = find_col(df, ["Currency"])
    if not start_col or not name_col:
        return pd.DataFrame(columns=["event_time", "name", "impact", "currency"])
    out = pd.DataFrame()
    out["event_time"] = parse_datetime_series(df[start_col])
    out["name"] = df[name_col].fillna("").astype(str)
    out["impact"] = df[impact_col].fillna("Unknown").astype(str) if impact_col else "Unknown"
    out["currency"] = df[curr_col].fillna("").astype(str) if curr_col else ""
    out = out.dropna(subset=["event_time"]).sort_values("event_time")
    out = out[out["name"].str.strip() != ""]
    return out


def resample_ohlcv(df: pd.DataFrame, rule: str) -> pd.DataFrame:
    if rule == "Raw / no merge":
        return df.copy()
    tmp = df.set_index("datetime").sort_index()
    agg = tmp.resample(rule).agg(
        open=("open", "first"),
        high=("high", "max"),
        low=("low", "min"),
        close=("close", "last"),
        volume=("volume", "sum"),
        event=("event", lambda x: "; ".join(sorted(set([str(v) for v in x if str(v).strip() and str(v) != "nan"])))[:300]),
        impact=("impact", lambda x: ", ".join(sorted(set([str(v) for v in x if str(v).strip() and str(v) != "nan"]))) or "Unknown"),
    ).dropna(subset=["open", "high", "low", "close"])
    return agg.reset_index()


def impact_rank(x: str) -> int:
    x = str(x)
    if "High" in x:
        return 0
    if "Medium" in x:
        return 1
    if "Low" in x:
        return 2
    return 3


def filter_events(events: pd.DataFrame, impacts: List[str], start, end) -> pd.DataFrame:
    if events.empty:
        return events
    ev = events[(events["event_time"] >= start) & (events["event_time"] <= end)].copy()
    if impacts and "All" not in impacts:
        ev = ev[ev["impact"].isin(impacts)]
    return ev.sort_values(["event_time", "impact"], key=lambda s: s.map(impact_rank) if s.name == "impact" else s)


def make_chart(candles: pd.DataFrame, events: pd.DataFrame, title: str, show_volume=True, show_events=True):
    fig = make_subplots(
        rows=2 if show_volume else 1,
        cols=1,
        shared_xaxes=True,
        vertical_spacing=0.03,
        row_heights=[0.72, 0.28] if show_volume else [1.0],
    )

    fig.add_trace(
        go.Candlestick(
            x=candles["datetime"],
            open=candles["open"],
            high=candles["high"],
            low=candles["low"],
            close=candles["close"],
            name="Candles",
            increasing_line_color="#16a34a",
            decreasing_line_color="#dc2626",
        ),
        row=1,
        col=1,
    )

    if show_volume:
        fig.add_trace(
            go.Bar(x=candles["datetime"], y=candles["volume"], name="Volume", opacity=0.45),
            row=2,
            col=1,
        )

    if show_events and not events.empty:
        y_top = candles["high"].max()
        y_bottom = candles["low"].min()
        for _, ev in events.iterrows():
            color = IMPACT_COLORS.get(str(ev["impact"]), "gray")
            fig.add_vline(x=ev["event_time"], line_dash="dot", line_width=1, line_color=color, row=1, col=1)
        fig.add_trace(
            go.Scatter(
                x=events["event_time"],
                y=[y_top] * len(events),
                mode="markers",
                marker=dict(size=9, color=[IMPACT_COLORS.get(str(v), "gray") for v in events["impact"]], symbol="diamond"),
                name="Economic Events",
                customdata=events[["name", "impact", "currency"]].values,
                hovertemplate="<b>%{customdata[0]}</b><br>Impact: %{customdata[1]}<br>Currency: %{customdata[2]}<br>Time: %{x}<extra></extra>",
            ),
            row=1,
            col=1,
        )

    fig.update_layout(
        title=title,
        height=760,
        hovermode="x unified",
        xaxis_rangeslider_visible=False,
        margin=dict(l=20, r=20, t=60, b=20),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
    )
    fig.update_yaxes(title_text="Price", row=1, col=1)
    if show_volume:
        fig.update_yaxes(title_text="Volume", row=2, col=1)
    return fig


def calculate_reaction(df: pd.DataFrame, event_time: pd.Timestamp, before_min=30, after_min=30):
    window = df[(df["datetime"] >= event_time - pd.Timedelta(minutes=before_min)) & (df["datetime"] <= event_time + pd.Timedelta(minutes=after_min))]
    pre = df[df["datetime"] <= event_time].tail(1)
    post = df[df["datetime"] >= event_time + pd.Timedelta(minutes=after_min)].head(1)
    if pre.empty or post.empty:
        return None
    start_price = float(pre["close"].iloc[0])
    end_price = float(post["close"].iloc[0])
    change = end_price - start_price
    pct = change / start_price * 100 if start_price else None
    return {
        "start_price": start_price,
        "end_price": end_price,
        "change": change,
        "pct_change": pct,
        "window_high": float(window["high"].max()) if not window.empty else None,
        "window_low": float(window["low"].min()) if not window.empty else None,
        "window_volume": float(window["volume"].sum()) if not window.empty else None,
    }


def macro_bias(indicator_name: str, actual, forecast) -> str:
    name = indicator_name.lower()
    if pd.isna(actual) or pd.isna(forecast):
        return "Forecast not available"
    surprise = actual - forecast
    if any(k in name for k in ["cpi", "inflation", "wage", "employment", "gdp", "pmi"]):
        return "Hawkish / growth-positive surprise" if surprise > 0 else "Dovish / slowdown surprise" if surprise < 0 else "In line"
    if "unemployment" in name or "unemployed" in name:
        return "Dovish labour-market weakness" if surprise > 0 else "Hawkish labour-market strength" if surprise < 0 else "In line"
    return "Positive surprise" if surprise > 0 else "Negative surprise" if surprise < 0 else "In line"

# -----------------------------
# Sidebar
# -----------------------------

st.title("IB Futures Assignment — Macro Event Reaction Dashboard")
st.caption("Candlesticks + volume + merged candles + economic event overlays + Australia macro context")

with st.sidebar:
    st.header("1) Upload Files")
    market_upload = st.file_uploader("Upload YIBc trades/event workbook", type=["xlsx"], key="market")
    macro_upload = st.file_uploader("Upload Australia macro indicators workbook", type=["xlsx"], key="macro")

market_xls = load_excel(market_upload, DEFAULT_MARKET_FILE)
macro_xls = None
try:
    macro_xls = load_excel(macro_upload, DEFAULT_MACRO_FILE)
except Exception:
    macro_xls = None

market_sheets = market_xls.sheet_names
candle_sheets = [s for s in market_sheets if re.search(r"YIBc|Trades", s, re.I)]
calendar_sheet = "Economic Calendar" if "Economic Calendar" in market_sheets else None

with st.sidebar:
    st.header("2) Chart Controls")
    selected_sheet = st.selectbox("Choose candle / trade sheet", candle_sheets, index=0 if candle_sheets else None)
    merge_rule = st.selectbox("Merge candles / resample interval", ["Raw / no merge", "5min", "15min", "30min", "1h", "4h", "1D"], index=2)
    selected_impacts = st.multiselect("Event impact category", ["All"] + IMPACT_ORDER, default=["All"])
    show_volume = st.checkbox("Show volume", value=True)
    show_events = st.checkbox("Show event markers", value=True)
    event_window_hours = st.slider("Event replay window ± hours", 1, 24, 3)

# -----------------------------
# Load data
# -----------------------------

try:
    raw_market = read_sheet(market_xls, selected_sheet)
    market_df = normalize_market_df(raw_market)
except Exception as e:
    st.error(f"Could not load selected market sheet: {e}")
    st.stop()

if calendar_sheet:
    calendar_df = normalize_calendar_df(read_sheet(market_xls, calendar_sheet))
else:
    # fallback: derive event list from market rows
    derived = market_df.dropna(subset=["nearest_event_time"]).copy()
    calendar_df = derived.rename(columns={"nearest_event_time": "event_time", "event": "name"})[["event_time", "name", "impact"]]
    calendar_df["currency"] = ""
    calendar_df = calendar_df.drop_duplicates()

# Date range controls after loading
min_dt, max_dt = market_df["datetime"].min(), market_df["datetime"].max()
with st.sidebar:
    st.header("3) Date Range")
    start_dt, end_dt = st.slider(
        "Select time range",
        min_value=min_dt.to_pydatetime(),
        max_value=max_dt.to_pydatetime(),
        value=(min_dt.to_pydatetime(), max_dt.to_pydatetime()),
        format="YYYY-MM-DD HH:mm",
    )

filtered_market = market_df[(market_df["datetime"] >= pd.Timestamp(start_dt)) & (market_df["datetime"] <= pd.Timestamp(end_dt))].copy()
if filtered_market.empty:
    st.warning("No candle data in selected date range.")
    st.stop()

candles = resample_ohlcv(filtered_market, merge_rule)
events = filter_events(calendar_df, selected_impacts, candles["datetime"].min(), candles["datetime"].max())

# -----------------------------
# Top KPI row
# -----------------------------

c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("Rows", f"{len(filtered_market):,}")
c2.metric("Candles after merge", f"{len(candles):,}")
c3.metric("Events shown", f"{len(events):,}")
c4.metric("Total volume", f"{filtered_market['volume'].sum():,.0f}")
c5.metric("Price change", f"{filtered_market['close'].iloc[-1] - filtered_market['close'].iloc[0]:,.2f}")

# -----------------------------
# Tabs
# -----------------------------

tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "Candlestick + Events",
    "Event Replay",
    "Trades Within 30min",
    "Australia Macro Context",
    "IB Assignment Notes",
])

with tab1:
    st.subheader("Candlestick chart with volume and macro event markers")
    fig = make_chart(candles, events, f"{selected_sheet} | {merge_rule} candles", show_volume, show_events)
    st.plotly_chart(fig, use_container_width=True)

    st.subheader("Events in selected range")
    if events.empty:
        st.info("No events found for the selected impact filter and date range.")
    else:
        st.dataframe(events, use_container_width=True, hide_index=True)

with tab2:
    st.subheader("Replay market reaction around a selected event")
    if events.empty:
        st.info("Select a wider date range or choose All impacts to see events.")
    else:
        event_labels = events.apply(lambda r: f"{r['event_time']} | {r['impact']} | {r['name']}", axis=1).tolist()
        choice = st.selectbox("Choose event", event_labels)
        selected_event = events.iloc[event_labels.index(choice)]
        etime = selected_event["event_time"]
        replay_df = market_df[(market_df["datetime"] >= etime - pd.Timedelta(hours=event_window_hours)) & (market_df["datetime"] <= etime + pd.Timedelta(hours=event_window_hours))].copy()
        replay_candles = resample_ohlcv(replay_df, merge_rule)
        replay_events = events[(events["event_time"] >= etime - pd.Timedelta(hours=event_window_hours)) & (events["event_time"] <= etime + pd.Timedelta(hours=event_window_hours))]
        st.markdown(f"**Selected event:** {selected_event['name']}  ")
        st.markdown(f"**Impact:** {selected_event['impact']} | **Time:** {selected_event['event_time']}")
        if replay_candles.empty:
            st.warning("No market candles available around this event window.")
        else:
            st.plotly_chart(make_chart(replay_candles, replay_events, f"Event replay ±{event_window_hours}h", show_volume, True), use_container_width=True)

            r30 = calculate_reaction(market_df, etime, 30, 30)
            r60 = calculate_reaction(market_df, etime, 60, 60)
            col_a, col_b = st.columns(2)
            with col_a:
                st.markdown("### 30-minute reaction")
                st.json(r30 or {"message": "Not enough data"})
            with col_b:
                st.markdown("### 1-hour reaction")
                st.json(r60 or {"message": "Not enough data"})

with tab3:
    st.subheader("Trades / candles matched within 30 minutes of events")
    if "Trades Within 30min" in market_sheets:
        trades_raw = read_sheet(market_xls, "Trades Within 30min")
        st.dataframe(trades_raw, use_container_width=True, hide_index=True)
        # Summary if columns exist
        impact_col = find_col(trades_raw, ["Impact", "Matched_Impact"])
        event_col = find_col(trades_raw, ["Matched Event(s)", "Matched_Event(s)"])
        if impact_col:
            st.markdown("### Count by event impact")
            impact_summary = trades_raw[impact_col].fillna("Unknown").value_counts().reset_index()
            impact_summary.columns = ["Impact", "Count"]
            st.bar_chart(impact_summary.set_index("Impact"))
        if event_col:
            st.markdown("### Most frequent matched events")
            event_summary = trades_raw[event_col].fillna("Unknown").value_counts().head(20).reset_index()
            event_summary.columns = ["Event", "Count"]
            st.dataframe(event_summary, use_container_width=True, hide_index=True)
    else:
        st.info("No 'Trades Within 30min' sheet found.")

with tab4:
    st.subheader("Australia macro context: actual vs forecast")
    if macro_xls is None:
        st.info("Upload the Australia macro workbook to use this tab.")
    else:
        macro_sheets = macro_xls.sheet_names
        default_inds = [s for s in macro_sheets if any(k in s.lower() for k in ["cpi", "unemployment", "gdp", "employment", "pmi"])]
        selected_indicators = st.multiselect("Choose macro indicators", macro_sheets, default=default_inds[:4])
        if not selected_indicators:
            st.info("Select at least one macro indicator.")
        else:
            summary_rows = []
            for ind in selected_indicators:
                mdf = read_sheet(macro_xls, ind)
                date_col = find_col(mdf, ["Release Date", "Reference Period"])
                actual_col = find_col(mdf, ["Actual"])
                forecast_col = find_col(mdf, ["Median_Forecast", "Average_Forecast", "Forecast"])
                if not date_col or not actual_col:
                    continue
                mdf["date"] = parse_datetime_series(mdf[date_col])
                mdf["actual"] = pd.to_numeric(mdf[actual_col], errors="coerce")
                mdf["forecast"] = pd.to_numeric(mdf[forecast_col], errors="coerce") if forecast_col else pd.NA
                mdf = mdf.dropna(subset=["date", "actual"]).sort_values("date")
                if mdf.empty:
                    continue
                latest = mdf.iloc[-1]
                summary_rows.append({
                    "Indicator": ind,
                    "Latest Date": latest["date"],
                    "Actual": latest["actual"],
                    "Forecast": latest["forecast"],
                    "Surprise": latest["actual"] - latest["forecast"] if pd.notna(latest["forecast"]) else pd.NA,
                    "Macro Bias": macro_bias(ind, latest["actual"], latest["forecast"]),
                })
                fig_m = go.Figure()
                fig_m.add_trace(go.Scatter(x=mdf["date"], y=mdf["actual"], mode="lines+markers", name="Actual"))
                if forecast_col:
                    fig_m.add_trace(go.Scatter(x=mdf["date"], y=mdf["forecast"], mode="lines+markers", name="Forecast"))
                fig_m.update_layout(title=ind, height=360, margin=dict(l=20, r=20, t=50, b=20), hovermode="x unified")
                st.plotly_chart(fig_m, use_container_width=True)
            if summary_rows:
                st.markdown("### Latest macro summary")
                st.dataframe(pd.DataFrame(summary_rows), use_container_width=True, hide_index=True)

with tab5:
    st.subheader("How to use this for the IB Futures Assignment")
    st.markdown(
        """
        **Main idea:** Use the dashboard to connect macro events with futures price action.

        **Workflow:**
        1. Select the correct YIBc sheet.
        2. Choose `Raw`, `5min`, `15min`, `30min`, or `1h` candles.
        3. Filter events by `High`, `Medium`, or `Low` impact.
        4. Use the event replay tab to study the reaction before and after the release.
        5. Check whether volume expanded after the event.
        6. Use the macro context tab to explain why the reaction happened.

        **Interpretation examples:**
        - Hot CPI → higher rate expectations → bonds down / yields up → equity futures may fall.
        - Weak employment → rate-cut expectations → bonds up / yields down → equity futures may rise.
        - Strong GDP/PMI → growth optimism, but also possible hawkish central-bank repricing.
        - High volume after an event means the move had stronger market participation.
        - Large wick + high volume can indicate liquidity sweep or absorption.
        """
    )
