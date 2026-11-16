@echo off
python -m pip install -r requirements_demand_forecast.txt
python -m streamlit run demand_forecasting_app.py
pause
