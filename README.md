# Harmony Manufacturing Forecast Engine

A lightweight, Excel-first Streamlit Cloud application. Upload an Excel workbook or CSV and the app automatically:
- scans worksheets and detects monthly date columns or transaction-level Date + Quantity data;
- identifies customer, SKU, product and plant fields where column names make them recognisable;
- excludes columns labelled Budget, Forecast, Plan, Target, Projection, Variance or Increment from historical actuals;
- classifies demand patterns as Seasonal, Growing, Declining, Sparse, New NPI / Limited history or Stable;
- compares benchmark, ARIMA, ETS/Holt-Winters, Croston SBA, an ensemble and selected ML methods where enough history exists;
- selects a model using chronological holdout WAPE (with RMSE and bias also reported);
- produces an editable 24-month forecast and Excel export.

## Deploy to Streamlit Community Cloud

1. Create a GitHub repository, for example `harmony-manufacturing-forecast-engine`.
2. Upload `app.py`, `requirements.txt`, and this `README.md` to the repository root.
3. Open https://share.streamlit.io/ and deploy the repository.
4. Set the Main file path to `app.py`.
5. No secrets, database, Redis, Docker, Node.js or external API key are required.

## Local run (optional)

```bash
python -m venv .venv
# Windows PowerShell:
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
streamlit run app.py
```

## Input data

Supported:
- `.xlsx`, `.xlsm`, `.xls`, `.csv`
- Wide monthly layout (one column per month, e.g. Apr-24, May-24)
- Long transaction layout (date/month column plus quantity column)

The engine detects common column names. For best automatic detection, use clear headers such as Customer, SKU, Product, Plant, Month, Actual Qty.

## Important limitations and safeguards

- Automatic detection is heuristic. Review the detected worksheet summary and data quality output.
- Budget/forecast columns are deliberately excluded from actuals. If your workbook labels actuals unusually, the app may not identify them correctly.
- New product/NPI demand cannot be inferred reliably from no sales history. This version uses a conservative benchmark for limited-history items and does not fabricate nomination or ramp-up quantities. Include a customer nomination/order quantity history in the source workbook for a future NPI-specific extension.
- Plant capacity and confirmed orders are not assumed unless represented in recognised historical demand fields. This version does not yet optimise production scheduling or enforce capacity limits.
- WAPE, RMSE and bias are holdout diagnostics, not a guarantee. Short histories and sparse demand make accuracy estimates unstable.
- Seasonal models and ML are used only when the series has enough history. Croston SBA is used for sparse demand.
- For performance on Streamlit Cloud, this version caps processing at 1,500 combinations per run (highest-volume combinations first). Increase `max_groups` in `app.py` if your deployment resources allow it.
- Forecast quantities are non-negative. Users can edit Forecast Qty and Increment % in the forecast grid. Historical actuals are kept in a separate output sheet.

## Files

- `app.py` — Streamlit application and forecasting engine
- `requirements.txt` — deployment dependencies
