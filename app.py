import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots

st.set_page_config(page_title="IB Macro Event Replay Dashboard", layout="wide")

st.title("IB Macro Event Replay Dashboard")
st.caption("Candles only around HIGH impact macro events")

uploaded_file = st.file_uploader("Upload YIBc Excel file", type=["xlsx"])

if uploaded_file is not None:
    xls = pd.ExcelFile(uploaded_file)
    sheets = [s for s in xls.sheet_names if s.startswith("YIBc")]

    sheet = st.sidebar.selectbox("Select Futures Sheet", sheets)

    df = pd.read_excel(uploaded_file, sheet_name=sheet)

    df["AUS_Local_DateTime"] = pd.to_datetime(df["AUS_Local_DateTime"], errors="coerce")
    df = df.dropna(subset=["AUS_Local_DateTime"])
    df = df.sort_values("AUS_Local_DateTime")

    required_cols = ["Open", "High", "Low", "Last", "Volume", "Matched_Event(s)", "Matched_Impact"]

    missing = [col for col in required_cols if col not in df.columns]

    if missing:
        st.error(f"Missing columns: {missing}")
        st.stop()

    high_events = df[
        df["Matched_Impact"].astype(str).str.lower().str.contains("high", na=False)
    ].copy()

    high_events = high_events.dropna(subset=["Matched_Event(s)"])

    if high_events.empty:
        st.warning("No High impact matched events found in this sheet.")
        st.stop()

    st.sidebar.header("Event Controls")

    event_list = sorted(high_events["Matched_Event(s)"].astype(str).unique())

    selected_event = st.sidebar.selectbox("Select High Impact Event", event_list)

    event_rows = high_events[
        high_events["Matched_Event(s)"].astype(str) == selected_event
    ]

    selected_time = st.sidebar.selectbox(
        "Select Event Time",
        event_rows["AUS_Local_DateTime"].drop_duplicates().sort_values()
    )

    before_hours = st.sidebar.slider("Hours Before Event", 1, 6, 2)
    after_hours = st.sidebar.slider("Hours After Event", 1, 6, 2)

    timeframe = st.sidebar.selectbox(
        "Merge Candles",
        ["1min", "5min", "15min", "30min", "1h"],
        index=2
    )

    window = df[
        (df["AUS_Local_DateTime"] >= selected_time - pd.Timedelta(hours=before_hours)) &
        (df["AUS_Local_DateTime"] <= selected_time + pd.Timedelta(hours=after_hours))
    ].copy()

    if window.empty:
        st.warning("No candle data found around this event.")
        st.stop()

    window = window.set_index("AUS_Local_DateTime")

    merged = window.resample(timeframe).agg({
        "Open": "first",
        "High": "max",
        "Low": "min",
        "Last": "last",
        "Volume": "sum"
    }).dropna()

    col1, col2, col3, col4 = st.columns(4)

    col1.metric("Selected Sheet", sheet)
    col2.metric("Impact", "High")
    col3.metric("Candles Shown", len(merged))
    col4.metric("Merge", timeframe)

    st.subheader(f"Event Replay: {selected_event}")
    st.write("Event Time:", selected_time)

    fig = make_subplots(
        rows=2,
        cols=1,
        shared_xaxes=True,
        vertical_spacing=0.03,
        row_heights=[0.75, 0.25]
    )

    fig.add_trace(
        go.Candlestick(
            x=merged.index,
            open=merged["Open"],
            high=merged["High"],
            low=merged["Low"],
            close=merged["Last"],
            name="Candles"
        ),
        row=1,
        col=1
    )

    fig.add_trace(
        go.Bar(
            x=merged.index,
            y=merged["Volume"],
            name="Volume"
        ),
        row=2,
        col=1
    )

    fig.add_vline(
        x=selected_time,
        line_width=2,
        line_dash="dash",
        line_color="red"
    )

    fig.add_annotation(
        x=selected_time,
        y=merged["High"].max(),
        text="HIGH IMPACT EVENT",
        showarrow=True,
        arrowhead=2,
        bgcolor="red",
        font=dict(color="white")
    )

    fig.update_layout(
        title=f"{selected_event} | {sheet}",
        height=850,
        xaxis_rangeslider_visible=False,
        template="plotly_dark",
        showlegend=False
    )

    st.plotly_chart(fig, use_container_width=True)

    st.subheader("Event Details")

    show_cols = [
        "AUS_Local_DateTime",
        "Matched_Event(s)",
        "Matched_Impact",
        "Minutes_From_Nearest_Event"
    ]

    available_cols = [c for c in show_cols if c in event_rows.columns]

    st.dataframe(
        event_rows[available_cols].drop_duplicates(),
        use_container_width=True
    )

else:
    st.info("Upload your YIBc Excel file to begin.")
