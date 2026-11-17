# Harmony Manufacturing Forecast Engine — v2

## What changed in v2

This version is built around the provided manufacturing workbook structure, rather than a generic sales CSV:
- Detects plant sheets such as `Manesar` and `Gujrat` and excludes the aggregated `Combined Sales` tab to prevent double-counting.
- Reads customer, model, vehicle, part number, SAP code, product, segment, material, specification, weight and pricing fields.
- Detects the monthly date headers in row 6 (the reference workbook uses AJ:CQ).
- Provides visible global method selection plus per-part method overrides.
- Provides editable FY1 and FY2 percentage changes per part (e.g. 1, 5, -5).
- Displays last available month quantity, last 3-month mean and last 12-month mean.
- Compares Seasonal Naive, ETS/Holt-Winters, ARIMA, Ridge, Random Forest, Croston SBA, Moving Average, Ensemble and an Existing Plan (NPI) fallback when selected.
- Uses chronological holdout WAPE, RMSE and bias for method comparison.
- Preserves every original worksheet. Creates copied plant sheets for each selected forecast year and writes forecasted monthly quantities only into those copies.
- Adds `Forecast Summary`, `Model Comparison` and `Forecast Controls` sheets with part / plant / customer / price / method / % change / annual quantity / value details.

## Deploy to Streamlit Community Cloud

1. Create a GitHub repository for this app.
2. Upload `app.py`, `requirements.txt`, and `README.md` to the repository root.
3. Deploy at https://share.streamlit.io/
4. Select `app.py` as the main file.
5. Upload your manufacturing workbook in the deployed app.

No database, Redis, Docker, Node.js, secrets or API keys are required.

## Input workbook layout

The supplied reference workbook has these plant tabs:
- `Manesar`
- `Gujrat`

The part-level structure is:
- Customer: column C
- Model: column D
- Vehicle name plate: column E
- Part No.: column F
- SAP code: column H
- Product: column I
- Segment: column J
- Material: column L
- Specification: column M
- Weight: column R
- Monthly dates: row 6, AJ:CQ
- Monthly quantity values: part rows from row 11 onward

The app detects the actual plant sheets and dates; it does not hard-code the forecast year labels.

## Important data interpretation

The provided plant sheets label monthly quantities as `Plan`. This version therefore uses those values as an existing plan baseline and explicitly labels them as such. It does not misrepresent them as actual demand. The workbook's `Sheet1 (2)` contains FY24-25 customer-level actuals, but those are not part-number-level actuals. For genuinely demand-validated SKU forecasts, provide part-number-by-month actual dispatch/sales history.

Backtest metrics are chronological holdout diagnostics on the available source series. If the source is a plan series, the metrics are plan-baseline validation, not proof of real-world demand accuracy.

## Forecasting methods

- Seasonal Naive: repeats the prior-year month pattern when enough history exists.
- ETS / Holt-Winters: exponential smoothing; additive seasonality is used when history supports it.
- ARIMA: compact ARIMA(1,1,1) candidate.
- Ridge / Random Forest: lag and calendar-feature models when sufficient history exists.
- Croston SBA: intermittent / sparse demand.
- Moving Average: benchmark fallback.
- Ensemble: averages the selected available candidate forecasts.

The app uses non-negative quantity forecasts. Methods that cannot be fitted because history is insufficient fall back to the moving-average benchmark for the forecast grid; validation notes are shown.

## Output workbook

Original sheets are retained unchanged. For each selected forecast year and detected plant, the app creates a copied sheet such as `FCST FY-27-28 Manesar`. It updates only the selected year's monthly cells in the copy and appends forecast diagnostics. The output also adds:
- `Forecast Summary`
- `Model Comparison`
- `Forecast Controls`

The copied plant sheets retain the source layout, metadata and pricing columns. Excel recalculates workbook formulas when opened.


## Optional external business variables (v2.1)

The UI includes optional assumptions for each forecast year. Every variable defaults to disabled and 0%, so it has no effect unless the user turns it on:
- Car / vehicle market outlook: annual demand outlook adjustment.
- Working days / weekday availability: net change in operating weekdays and working days.
- Customer nominations / confirmed orders: customer schedule and order-book adjustment.
- Holiday / festival season effect: selected-month adjustment.
- Plant shutdown / maintenance: selected-month production or dispatch reduction.
- Promotions / customer campaigns: selected-month temporary adjustment.

Month-scoped factors require the user to select the applicable months. Multiple enabled factors are compounded multiplicatively. The part-level FY1/FY2 percentage is applied as well. The UI preview and exported plant forecast sheets, Forecast Summary and Forecast Assumptions sheet show the resulting effects.
