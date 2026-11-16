# Demand Forecasting Program

## Included files
- `demand_forecasting_app.py` — Streamlit application
- `requirements_demand_forecast.txt` — Python dependencies
- `run_demand_forecast.bat` — Windows launcher
- `1.Sales Volume Projections 25-26 to 29-30_12.11.2024_Rev11-Master.xlsx` — supplied source workbook

## Run on Windows
1. Extract this ZIP into a folder.
2. Install Python 3.10 or newer.
3. Double-click `run_demand_forecast.bat`.
4. Open the local Streamlit address shown in the terminal (normally http://localhost:8501).
5. The supplied master workbook loads automatically. You can also upload a revised `.xlsx` workbook in the sidebar.

## Workflow
1. Select the plant, customer or product filters.
2. Open each product and review all extracted source parameters.
3. Review the trailing 12-month mean and trailing 3-month mean.
4. Enter the planned annual growth percentage and confirm the assumptions.
5. Use the Client & Order Register to add a client/order. The entered monthly quantity becomes the new product's starting monthly run-rate. A Delete request attempts to match plant + client + part number/product and excludes matching product(s) from the forecast.
6. Select a 1- or 2-year horizon and forecast start month.
7. Build and download the Excel workbook.

## Calculation method
- Trailing 12-month mean = arithmetic mean of the latest 12 available monthly source values before the forecast start.
- Trailing 3-month mean = arithmetic mean of the latest 3 available monthly source values before the forecast start.
- Blended monthly base = 50% × trailing 12-month mean + 50% × trailing 3-month mean.
- Monthly seasonality is taken from the latest available 12-month pattern and normalized to the blended monthly base.
- Annual growth is compounded by forecast year and applied to each forecast month.

## Important source-data qualification
The Manesar and Gujrat monthly cells in the supplied workbook are labelled `Plan`. The averages therefore represent trailing values in the supplied planning series, not independently verified actual sales. Replace the monthly inputs with actuals if actual-sales averages are required. The program displays the available observation counts and retains all original workbook sheets in the export.

## Export
The generated workbook preserves the original worksheets and appends:
- `Forecast Assumptions`
- `Forecast Monthly`
- `Forecast Manesar`
- `Forecast Gujrat`
- `Client & Order Register`
- `Forecast Read Me`

The original source sheets are not modified or physically deleted.
