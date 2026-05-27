# IB Futures Assignment — Macro Event Reaction Dashboard

This dashboard uses your YIBc futures/trades workbook and Australia macro indicators workbook.

## Files needed

Put these files in the same folder:

- `app.py`
- `requirements.txt`
- `YIBc_Trades_30min_Event_Matches(3).xlsx`
- `Australia_Economic_Indicators_2022_2024(1).xlsx`

You can also upload both Excel files from the dashboard sidebar.

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Main features

- Candlestick chart
- Volume bars
- Merge candles / resampling: 5min, 15min, 30min, 1h, 4h, 1D
- High / Medium / Low event filtering
- Economic event markers on the chart
- Event replay window around macro releases
- 30-minute and 1-hour reaction calculation
- Trades within 30min table
- Australia macro context with Actual vs Forecast

## Suggested assignment workflow

1. Open the candlestick tab.
2. Filter only High impact events first.
3. Select 5min or 15min candles.
4. Go to Event Replay.
5. Choose an event like CPI, employment, GDP, or RBA event.
6. Compare price change, volume spike, and macro surprise.
7. Write the explanation in terms of rates, inflation, growth, bond yields, and futures reaction.
