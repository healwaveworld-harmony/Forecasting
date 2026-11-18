import io
import re
import math
import warnings
from datetime import datetime
from copy import copy

import numpy as np
import pandas as pd
import streamlit as st
import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter

APP_VERSION = "2.2.0"
MONTH_START_COL = 36   # AJ
MONTH_END_COL = 95     # CQ, inclusive
MONTH_HEADER_ROW = 6
FIRST_PART_ROW = 11
PART_COL = 6           # F
CUSTOMER_COL = 3       # C
MODEL_COL = 4          # D
VEHICLE_COL = 5        # E
SAP_COL = 8            # H
PRODUCT_COL = 9        # I
SEGMENT_COL = 10       # J
MATERIAL_COL = 12      # L
SPEC_COL = 13          # M
WEIGHT_COL = 18        # R
SALES_PRICE_COL = 109  # DE
PRICE_PC_COL = 121     # DQ

METHODS = [
    "Seasonal Naive",
    "ETS / Holt-Winters",
    "ARIMA",
    "Ridge",
    "Random Forest",
    "Croston SBA",
    "Moving Average",
    "Ensemble",
    "Existing Plan (NPI)",
]

MODEL_DESCRIPTIONS = {
    "Seasonal Naive": ("Repeats the quantity from the same month in the latest available year.",
                       "Useful when the same seasonal pattern is expected to repeat; needs at least 12 months for a full seasonal cycle."),
    "ETS / Holt-Winters": ("Exponential smoothing that learns the recent level, trend and (when enough data exists) monthly seasonality.",
                           "Recent observations receive more weight, so it adapts as demand changes."),
    "ARIMA": ("AutoRegressive Integrated Moving Average; models relationships between past values and changes over time.",
              "Useful for time-series patterns with autocorrelation; this implementation uses a fixed (1,1,1) specification."),
    "Ridge": ("Linear regression with regularisation, using recent demand lags and month-of-year signals.",
              "Regularisation limits unstable coefficients when predictors are related; requires sufficient history."),
    "Random Forest": ("A collection of decision trees that learns non-linear relationships from lagged demand and month-of-year features.",
                      "Can capture more complex patterns but may be less reliable with short histories."),
    "Croston SBA": ("Syntetos–Boylan Approximation for intermittent demand; estimates demand size and the interval between non-zero orders.",
                    "Designed for parts with many zero-demand months; produces a smoothed average demand rate."),
    "Moving Average": ("Uses the average of the latest three observed months as a constant monthly forecast.",
                       "A simple, fast benchmark when trend or seasonality is weak or history is limited."),
    "Ensemble": ("Averages forecasts from the available selected statistical methods.",
                 "Reduces dependence on one model; methods that fail for a part are skipped."),
    "Existing Plan (NPI)": ("Carries forward the uploaded plan for the target months when part history is insufficient or the part is new.",
                            "This is a plan fallback, not a model-derived estimate of actual demand."),
}
METRIC_DESCRIPTIONS = {
    "WAPE %": ("Weighted Absolute Percentage Error", "Total absolute forecast error divided by total actual/baseline volume. Lower is better; meaningful when the comparison total is above zero."),
    "RMSE": ("Root Mean Squared Error", "Square root of the average squared error, in the same units as quantity. Lower is better; large misses receive extra weight."),
    "Bias %": ("Forecast Bias", "Net over-forecast or under-forecast as a percentage of total comparison volume. Positive means over-forecast; negative means under-forecast."),
    "Forecast Bias %": ("Forecast Bias", "Net over-forecast or under-forecast as a percentage of total comparison volume. Positive means over-forecast; negative means under-forecast."),
    "NPI": ("New Product Introduction", "A new part or part with too little usable demand history for a reliable model fit."),
    "ETS": ("Exponential Smoothing", "A family of methods that weights recent observations more heavily and can model level, trend and seasonality."),
    "ARIMA": ("AutoRegressive Integrated Moving Average", "A statistical time-series model using lagged values and differenced data."),
    "SBA": ("Syntetos–Boylan Approximation", "A bias-adjusted Croston method for intermittent demand."),
}
DEMAND_PATTERN_DESCRIPTIONS = {
    "Seasonal": "Demand varies in a repeating month-of-year pattern.",
    "Growing": "Recent history shows an upward trend relative to its average.",
    "Declining": "Recent history shows a downward trend relative to its average.",
    "Sparse": "Demand is intermittent, with many zero or very few non-zero months.",
    "New NPI / Limited history": "Too few observations are available for a confident pattern classification.",
    "New NPI / No history": "No usable part-level monthly history was found.",
    "Stable": "No strong upward, downward or seasonal pattern was detected.",
}

st.set_page_config(page_title="Harmony Manufacturing Forecast Engine", page_icon="🏭", layout="wide")

def safe_text(v):
    if v is None:
        return ""
    if isinstance(v, float) and np.isnan(v):
        return ""
    return str(v).strip()

def numeric(v):
    if v is None or isinstance(v, str) and v.startswith("="):
        return np.nan
    try:
        if isinstance(v, str):
            v = v.replace(",", "").strip()
        return float(v)
    except Exception:
        return np.nan

def fiscal_label(dt):
    start = dt.year if dt.month >= 4 else dt.year - 1
    return f"FY{str(start)[-2:]}-{str(start+1)[-2:]}"

def detect_month_columns(ws_values):
    found = []
    for col in range(MONTH_START_COL, min(MONTH_END_COL, ws_values.max_column) + 1):
        v = ws_values.cell(MONTH_HEADER_ROW, col).value
        if isinstance(v, datetime):
            dt = pd.Timestamp(v).to_period("M").to_timestamp()
            found.append((col, dt, fiscal_label(dt)))
    return found

def detect_plant_sheets(wb_values):
    plants = []
    for ws in wb_values.worksheets:
        if ws.title.strip().lower() in {"combined sales", "summary", "summary (2)", "m-o-m vol vs cap. summary"}:
            continue
        if ws.max_row < FIRST_PART_ROW or ws.max_column < MONTH_START_COL:
            continue
        months = detect_month_columns(ws)
        part_header = safe_text(ws.cell(8, PART_COL).value).lower()
        if len(months) >= 12 and ("part" in part_header or safe_text(ws.cell(8, PART_COL).value) in ["部品番号", "Part No."]):
            plants.append(ws.title)
    return plants

def read_plant_records(wb_formula, wb_values, plant_names):
    records = []
    month_map = {}
    basis_by_year = {}
    for plant in plant_names:
        wsv = wb_values[plant]
        wsf = wb_formula[plant]
        month_cols = detect_month_columns(wsv)
        month_map[plant] = month_cols

        for col, dt, fy in month_cols:
            # The template labels monthly quantities as Plan in row 8.
            status = safe_text(wsv.cell(8, col).value).lower()
            if "actual" in status:
                basis = "Actual"
            elif "budget" in status:
                basis = "Budget"
            elif "plan" in status or "（plan）" in status or "(plan)" in status:
                basis = "Plan"
            else:
                basis = "Unspecified"
            basis_by_year.setdefault(fy, set()).add(basis)

        for r in range(FIRST_PART_ROW, wsv.max_row + 1):
            part = safe_text(wsv.cell(r, PART_COL).value) or safe_text(wsf.cell(r, PART_COL).value)
            customer = safe_text(wsv.cell(r, CUSTOMER_COL).value) or safe_text(wsf.cell(r, CUSTOMER_COL).value)
            if not part or not customer:
                continue
            # Skip section headings and subtotal rows.
            if part.lower() in ["part no.", "part number", "total", "subtotal"] or customer.lower() in ["customer", "total", "subtotal"]:
                continue

            series = {}
            for col, dt, fy in month_cols:
                v = numeric(wsv.cell(r, col).value)
                if np.isfinite(v):
                    series[dt] = max(0.0, float(v))
                else:
                    # If cached formula result is missing, retain a missing value rather than treating formula text as demand.
                    series[dt] = np.nan

            def val(col):
                v = wsv.cell(r, col).value
                if v is None:
                    v = wsf.cell(r, col).value
                return safe_text(v)

            records.append({
                "key": f"{plant}|{r}",
                "source_row": r,
                "plant": plant,
                "customer": customer,
                "model": val(MODEL_COL),
                "vehicle": val(VEHICLE_COL),
                "part_no": part,
                "sap_code": val(SAP_COL),
                "product": val(PRODUCT_COL),
                "segment": val(SEGMENT_COL),
                "material": val(MATERIAL_COL),
                "specification": val(SPEC_COL),
                "weight_kg": numeric(wsv.cell(r, WEIGHT_COL).value),
                "sales_price": numeric(wsv.cell(r, SALES_PRICE_COL).value),
                "price_inr_pc": numeric(wsv.cell(r, PRICE_PC_COL).value),
                "series": series,
            })
    return records, month_map, basis_by_year

def complete_history(series, before_date):
    items = [(d, v) for d, v in series.items() if d < before_date and np.isfinite(v)]
    if not items:
        return pd.Series(dtype=float)
    items.sort(key=lambda x: x[0])
    start, end = items[0][0], items[-1][0]
    idx = pd.date_range(start, end, freq="MS")
    raw = pd.Series({d: v for d, v in items}, dtype=float)
    s = raw.reindex(idx)
    # Internal missing months are not assumed to be zero unless there are numeric observations on both sides.
    s = s.interpolate(limit_area="inside")
    # Leading/trailing blanks remain missing; only trim to observed range.
    return s

def demand_pattern(y):
    y = np.asarray(y, dtype=float)
    y = y[np.isfinite(y)]
    if len(y) == 0:
        return "New NPI / No history"
    nz = y[y > 0]
    if len(nz) <= 2 or np.mean(y == 0) >= 0.45:
        return "Sparse"
    if len(y) < 12 or len(nz) < 6:
        return "New NPI / Limited history"
    avg = max(np.mean(y), 1e-9)
    x = np.arange(len(y))
    slope = np.polyfit(x, y, 1)[0]
    relative = slope * max(len(y)-1, 1) / avg
    if len(y) >= 24:
        month_avg = np.array([np.mean(y[i::12]) for i in range(12)])
        cv = np.std(month_avg) / (np.mean(month_avg) + 1e-9)
        if cv > 0.35:
            return "Seasonal"
    if relative > 0.25:
        return "Growing"
    if relative < -0.25:
        return "Declining"
    return "Stable"

def seasonal_naive(y, h):
    y = np.asarray(y, dtype=float)
    if len(y) >= 12:
        return np.array([y[-12 + (i % 12)] for i in range(h)], dtype=float)
    if len(y):
        return np.repeat(float(np.mean(y[-min(3, len(y)): ])), h)
    return np.zeros(h)

def moving_average(y, h):
    y = np.asarray(y, dtype=float)
    if len(y) == 0:
        return np.zeros(h)
    return np.repeat(max(0.0, float(np.mean(y[-min(3, len(y)):]))), h)

def croston_sba(y, h, alpha=0.1):
    y = np.asarray(y, dtype=float)
    nz = np.where(y > 0)[0]
    if len(nz) == 0:
        return np.zeros(h)
    first = int(nz[0])
    z = float(y[first])
    p = float(first + 1)
    last = first
    for i in nz[1:]:
        gap = int(i - last)
        z = alpha * float(y[i]) + (1-alpha) * z
        p = alpha * gap + (1-alpha) * p
        last = int(i)
    rate = (1 - alpha/2) * z / max(p, 1e-9)
    return np.repeat(max(0.0, rate), h)

def ets_hw(y, h):
    from statsmodels.tsa.holtwinters import ExponentialSmoothing
    arr = np.asarray(y, dtype=float)
    arr = np.maximum(0, arr)
    if len(arr) == 0 or np.nanstd(arr) < 1e-9:
        return moving_average(arr, h)
    if len(arr) >= 24:
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            fit = ExponentialSmoothing(arr, trend="add", seasonal="add", seasonal_periods=12, initialization_method="estimated").fit(optimized=True)
    elif len(arr) >= 8:
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            fit = ExponentialSmoothing(arr, trend="add", initialization_method="estimated").fit(optimized=True)
    else:
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            fit = ExponentialSmoothing(arr, initialization_method="estimated").fit(optimized=True)
    return np.maximum(0, np.asarray(fit.forecast(h), dtype=float))

def arima_forecast(y, h):
    from statsmodels.tsa.arima.model import ARIMA
    arr = np.asarray(y, dtype=float)
    if len(arr) < 12 or np.nanstd(arr) < 1e-9:
        return moving_average(arr, h)
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        fit = ARIMA(arr, order=(1,1,1), enforce_stationarity=False, enforce_invertibility=False).fit(method_kwargs={"maxiter":40})
    return np.maximum(0, np.asarray(fit.forecast(h), dtype=float))

def ml_forecast(y, h, method):
    from sklearn.linear_model import Ridge
    from sklearn.ensemble import RandomForestRegressor
    arr = np.asarray(y, dtype=float)
    lags = [1, 2, 3, 6, 12]
    def feat(hist, t):
        vals = [hist[t-l] if t-l >= 0 else np.nan for l in lags]
        vals += [np.sin(2*np.pi*(t % 12)/12), np.cos(2*np.pi*(t % 12)/12), t]
        return vals
    X, target = [], []
    for t in range(max(lags), len(arr)):
        X.append(feat(arr, t))
        target.append(arr[t])
    if len(X) < 6:
        raise ValueError("At least 18 monthly observations are needed for this ML model.")
    X = np.asarray(X, dtype=float)
    X = np.where(np.isnan(X), np.nanmean(arr), X)
    if method == "Ridge":
        model = Ridge(alpha=2.0)
    else:
        model = RandomForestRegressor(n_estimators=60, max_depth=5, min_samples_leaf=2, random_state=17, n_jobs=1)
    model.fit(X, np.asarray(target, dtype=float))
    hist = list(arr.astype(float))
    preds = []
    for step in range(h):
        t = len(hist)
        x = np.asarray(feat(np.asarray(hist), t), dtype=float)
        x = np.where(np.isnan(x), np.nanmean(arr), x).reshape(1, -1)
        p = max(0.0, float(model.predict(x)[0]))
        preds.append(p)
        hist.append(p)
    return np.asarray(preds)

def predict_method(y, h, method, candidates=None):
    arr = np.asarray(y, dtype=float)
    arr = np.nan_to_num(arr, nan=0.0)
    if method == "Seasonal Naive":
        return np.maximum(0, seasonal_naive(arr, h))
    if method == "Moving Average":
        return moving_average(arr, h)
    if method == "Croston SBA":
        return croston_sba(arr, h)
    if method == "ETS / Holt-Winters":
        return ets_hw(arr, h)
    if method == "ARIMA":
        return arima_forecast(arr, h)
    if method in ["Ridge", "Random Forest"]:
        return ml_forecast(arr, h, method)
    if method == "Ensemble":
        use = candidates or ["Seasonal Naive", "ETS / Holt-Winters", "ARIMA"]
        preds = []
        for m in use:
            if m == "Ensemble":
                continue
            try:
                preds.append(predict_method(arr, h, m))
            except Exception:
                pass
        return np.maximum(0, np.mean(np.vstack(preds), axis=0)) if preds else moving_average(arr, h)
    return moving_average(arr, h)

def calc_metrics(actual, pred):
    actual = np.asarray(actual, dtype=float)
    pred = np.asarray(pred, dtype=float)
    denom = float(np.sum(np.abs(actual)))
    wape = float(np.sum(np.abs(actual-pred)) / denom * 100) if denom > 0 else np.nan
    rmse = float(np.sqrt(np.mean((actual-pred)**2))) if len(actual) else np.nan
    bias = float(np.sum(pred-actual) / denom * 100) if denom > 0 else np.nan
    return wape, rmse, bias

def compare_methods(y, candidates):
    y = np.asarray(y, dtype=float)
    results = []
    if len(y) < 6:
        return pd.DataFrame([{"Method":"Moving Average","WAPE %":np.nan,"RMSE":np.nan,"Bias %":np.nan,"Validation":"Insufficient history"}])
    if np.nanstd(y) < 1e-9:
        return pd.DataFrame([{"Method":"Moving Average","WAPE %":0.0,"RMSE":0.0,"Bias %":0.0,"Validation":"Constant series; no competition needed"}])
    holdout = min(3, max(2, len(y)//5))
    if len(y) - holdout < 4:
        return pd.DataFrame([{"Method":"Moving Average","WAPE %":np.nan,"RMSE":np.nan,"Bias %":np.nan,"Validation":"Insufficient history"}])
    train, test = y[:-holdout], y[-holdout:]
    # Sparse series are compared with intermittent-demand and simple benchmarks;
    # ARIMA/ML are not meaningful when most periods are zero.
    effective = list(candidates)
    if np.mean(y == 0) >= 0.45:
        effective = [m for m in effective if m in ["Croston SBA","Seasonal Naive","Moving Average"]]
    if len(y) < 18:
        effective = [m for m in effective if m not in ["ARIMA","Ridge","Random Forest"]]
    for method in effective:
        try:
            pred = predict_method(train, holdout, method)
            w, r, b = calc_metrics(test, pred)
            results.append({"Method":method,"WAPE %":w,"RMSE":r,"Bias %":b,"Validation":f"Chronological holdout ({holdout} months)"})
        except Exception as e:
            results.append({"Method":method,"WAPE %":np.nan,"RMSE":np.nan,"Bias %":np.nan,"Validation":f"Not run: {str(e)[:90]}"})
    if not results:
        results.append({"Method":"Moving Average","WAPE %":np.nan,"RMSE":np.nan,"Bias %":np.nan,"Validation":"Fallback"})
    return pd.DataFrame(results).sort_values(["WAPE %","RMSE"], na_position="last")

def choose_best(metrics, pattern):
    if pattern == "Sparse":
        croston = metrics[metrics["Method"] == "Croston SBA"]
        if not croston.empty:
            return "Croston SBA"
    valid = metrics[metrics["WAPE %"].notna()]
    if not valid.empty:
        return str(valid.iloc[0]["Method"])
    return "Moving Average"

def month_window(series, target_start):
    hist = complete_history(series, target_start)
    vals = hist.dropna().values.astype(float)
    last = float(vals[-1]) if len(vals) else np.nan
    last3 = float(np.mean(vals[-3:])) if len(vals) else np.nan
    last12 = float(np.mean(vals[-12:])) if len(vals) else np.nan
    return hist, last, last3, last12

def make_forecast_cache(records, month_map, target_fys, candidates):
    cache = {}
    for rec in records:
        # One chronological model competition per part. The selected method is then
        # applied to both forecast years, using each year's available preceding series.
        first_months = [(col, dt, label) for col, dt, label in month_map[rec["plant"]] if label == target_fys[0]]
        if not first_months:
            continue
        first_start = min(dt for _, dt, _ in first_months)
        first_hist, _, _, _ = month_window(rec["series"], first_start)
        first_y = first_hist.dropna().values.astype(float)
        first_pattern = demand_pattern(first_y)
        first_metrics = compare_methods(first_y, candidates)
        default_method = choose_best(first_metrics, first_pattern)
        if len(first_y) < 6:
            first_plan = [rec["series"].get(dt, np.nan) for _, dt, _ in first_months]
            if any(np.isfinite(v) for v in first_plan):
                default_method = "Existing Plan (NPI)"
        for fy in target_fys:
            months = [(col, dt, label) for col, dt, label in month_map[rec["plant"]] if label == fy]
            if not months:
                continue
            target_start = min(dt for _, dt, _ in months)
            hist, last, mean3, mean12 = month_window(rec["series"], target_start)
            y = hist.dropna().values.astype(float)
            pattern = demand_pattern(y)
            target_plan = np.array([rec["series"].get(dt, np.nan) for _, dt, _ in months], dtype=float)
            # NPI / insufficient-history fallback: preserve the existing year plan rather
            # than inventing a zero forecast when there is no part-level history.
            if len(y) < 6 and np.isfinite(target_plan).any():
                forecast = np.nan_to_num(target_plan, nan=0.0)
                applied_default = "Existing Plan (NPI)"
            else:
                applied_default = default_method
                if len(y) == 0:
                    forecast = np.zeros(12)
                else:
                    try:
                        forecast = predict_method(y, 12, applied_default, candidates)
                    except Exception:
                        applied_default = "Moving Average"
                        forecast = moving_average(y, 12)
            cache[(rec["key"], fy)] = {
                "history": hist, "last":last, "mean3":mean3, "mean12":mean12,
                "pattern":pattern, "metrics":first_metrics.copy(),
                "default_method":applied_default, "forecasts":{applied_default:np.maximum(0, forecast)},
                "target_plan":target_plan,
                "target_months":[dt for _, dt, _ in months],
            }
    return cache

def ensure_forecast(info, method, candidates):
    history = info["history"].dropna().values.astype(float)
    # No SKU history: retain the uploaded plan for that NPI instead of inventing demand.
    if len(history) < 6 and np.isfinite(info["target_plan"]).any():
        info["forecasts"]["Existing Plan (NPI)"] = np.maximum(0, np.nan_to_num(info["target_plan"], nan=0.0))
        return info["forecasts"]["Existing Plan (NPI)"], "Existing Plan (NPI)"
    if method == "Existing Plan (NPI)":
        if np.isfinite(info["target_plan"]).any():
            info["forecasts"][method] = np.maximum(0, np.nan_to_num(info["target_plan"], nan=0.0))
        else:
            info["forecasts"][method] = np.maximum(0, moving_average(history, 12))
    elif method not in info["forecasts"]:
        try:
            info["forecasts"][method] = np.maximum(0, predict_method(history, 12, method, candidates))
        except Exception:
            info["forecasts"][method] = np.maximum(0, moving_average(history, 12))
    return info["forecasts"][method], method

def metric_for_method(info, method, candidates):
    metrics = info["metrics"]
    match = metrics[metrics["Method"] == method]
    if not match.empty:
        return match.iloc[0]
    history = info["history"].dropna().values.astype(float)
    if method == "Existing Plan (NPI)" or len(history) < 6:
        return pd.Series({"Method":method,"WAPE %":np.nan,"RMSE":np.nan,"Bias %":np.nan,"Validation":"NPI plan fallback / insufficient part history"})
    try:
        extra = compare_methods(history, [method])
        if not extra.empty:
            return extra.iloc[0]
    except Exception:
        pass
    return pd.Series({"Method":method,"WAPE %":np.nan,"RMSE":np.nan,"Bias %":np.nan,"Validation":"Method not validated"})

def style_new_sheet(ws, start_col, headers):
    thin = Side(style="thin", color="D9E2F3")
    for j, header in enumerate(headers):
        c = start_col + j
        for row, value in [(7, "Forecast Diagnostics"), (8, header), (9, ""), (10, "")]:
            cell = ws.cell(row, c)
            cell.value = value
            cell._style = copy(ws.cell(row, 132)._style)
            cell.alignment = copy(ws.cell(row, 132).alignment)
            cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)
            if row == 8:
                cell.font = Font(name="Arial", size=9, bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor="176B87")
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(c)].width = max(14, min(24, len(header)+3))

def add_table_sheet(wb, title, df):
    if title in wb.sheetnames:
        del wb[title]
    ws = wb.create_sheet(title)
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A2"
    ws.append(list(df.columns))
    for row in df.itertuples(index=False, name=None):
        ws.append([None if (isinstance(v, float) and np.isnan(v)) else v for v in row])
    header_fill = PatternFill("solid", fgColor="176B87")
    header_font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    thin = Side(style="thin", color="D9E2F3")
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = Border(bottom=thin)
    for col in ws.columns:
        letter = get_column_letter(col[0].column)
        maxlen = max((len(str(c.value)) if c.value is not None else 0 for c in col[:100]), default=12)
        ws.column_dimensions[letter].width = min(max(maxlen+2, 12), 30)
    ws.auto_filter.ref = ws.dimensions
    return ws


VARIABLE_DEFINITIONS = [
    ("vehicle_market", "Car / vehicle market outlook", "Expected change in vehicle sales or production that may influence demand for the selected parts.", "all"),
    ("working_days", "Working days / weekday availability", "Net demand effect from fewer or more operating weekdays and working days in the forecast year.", "all"),
    ("customer_orders", "Customer nominations / confirmed orders", "Change supported by customer schedules, nominations, blanket orders or confirmed order book.", "all"),
    ("holiday_effect", "Holiday / festival season effect", "Demand or dispatch impact during selected holiday and festival months; enter a negative value if closures reduce demand.", "months"),
    ("plant_shutdown", "Plant shutdown / maintenance", "Expected production or dispatch reduction during selected plant shutdown or maintenance months.", "months"),
    ("campaign_effect", "Promotions / customer campaigns", "Temporary demand change caused by launches, promotions, price actions or customer campaigns in selected months.", "months"),
]

MONTH_NAMES = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]

def variable_multiplier(fy, month_dt, assumptions):
    cfg = assumptions.get(fy, {})
    multiplier = 1.0
    for key, title, desc, scope in VARIABLE_DEFINITIONS:
        enabled = bool(cfg.get(key + "_enabled", False))
        pct = float(cfg.get(key + "_pct", 0.0) or 0.0)
        if not enabled or abs(pct) < 1e-12:
            continue
        if scope == "months":
            months = cfg.get(key + "_months", [])
            if MONTH_NAMES[month_dt.month - 1] not in months:
                continue
        multiplier *= max(0.0, 1.0 + pct / 100.0)
    return multiplier

def variable_effect_pct(fy, month_dt, assumptions):
    return (variable_multiplier(fy, month_dt, assumptions) - 1.0) * 100.0

def assumptions_to_rows(assumptions, target_fys):
    rows = []
    for fy in target_fys:
        cfg = assumptions.get(fy, {})
        for key, title, desc, scope in VARIABLE_DEFINITIONS:
            rows.append({
                "Forecast FY": fy,
                "Variable": title,
                "Enabled": bool(cfg.get(key + "_enabled", False)),
                "Expected Change %": float(cfg.get(key + "_pct", 0.0) or 0.0),
                "Applicable Months": ", ".join(cfg.get(key + "_months", [])) if scope == "months" else "All 12 months",
                "Description": desc,
            })
    return pd.DataFrame(rows)

def build_glossary_df():
    rows = []
    for method, (meaning, use) in MODEL_DESCRIPTIONS.items():
        rows.append({"Category":"Forecast model", "Term":method, "Meaning / full name":meaning, "How to interpret / use":use})
    for term, (full, meaning) in METRIC_DESCRIPTIONS.items():
        rows.append({"Category":"Metric / technical term", "Term":term, "Meaning / full name":full, "How to interpret / use":meaning})
    for term, meaning in DEMAND_PATTERN_DESCRIPTIONS.items():
        rows.append({"Category":"Demand pattern", "Term":term, "Meaning / full name":term, "How to interpret / use":meaning})
    rows.extend([
        {"Category":"Forecast terminology","Term":"Part Change %","Meaning / full name":"User-entered part-level adjustment","How to interpret / use":"A percentage applied to every forecast month for that part and forecast year."},
        {"Category":"Forecast terminology","Term":"External Variables Effect %","Meaning / full name":"Combined business-assumption adjustment","How to interpret / use":"Compounded impact of enabled market, working-day, order, holiday, shutdown and campaign assumptions."},
        {"Category":"Forecast terminology","Term":"Base Forecast Qty","Meaning / full name":"Model output before manual adjustments","How to interpret / use":"Forecast quantity before part-level percentage and external-variable adjustments."},
        {"Category":"Forecast terminology","Term":"Adjusted Forecast Qty","Meaning / full name":"Final monthly forecast quantity","How to interpret / use":"Base forecast multiplied by the part change and applicable external-variable factors."},
        {"Category":"Validation","Term":"Chronological holdout","Meaning / full name":"Time-ordered backtest","How to interpret / use":"The latest months are held out; earlier months train the model. It is not a random split."},
        {"Category":"Validation","Term":"Plan baseline","Meaning / full name":"Uploaded monthly plan values","How to interpret / use":"Where the source workbook labels monthly values as Plan, the comparison is plan-baseline validation, not actual-demand accuracy."},
    ])
    return pd.DataFrame(rows)


def build_output_workbook(raw_bytes, filename, target_fys, records, month_map, cache, edited, candidates, global_method, assumptions, selected_part_key=None):
    keep_vba = filename.lower().endswith(".xlsm")
    wb = openpyxl.load_workbook(io.BytesIO(raw_bytes), data_only=False, keep_vba=keep_vba)
    # Data-only workbook supplies cached formula results for calculations.
    wb_values = openpyxl.load_workbook(io.BytesIO(raw_bytes), data_only=True, keep_vba=keep_vba)
    by_key = {r["key"]:r for r in records}
    edited_by_key = {str(row["Part Key"]):row for _, row in edited.iterrows()}

    summary_rows, comparison_rows = [], []
    forecast_values = {}
    for rec in records:
        row_edit = edited_by_key.get(rec["key"])
        selected_method = str(row_edit["Selected Method"]) if row_edit is not None else cache[(rec["key"], target_fys[0])]["default_method"]
        changes = [
            float(row_edit["FY1 Change %"]) if row_edit is not None else 0.0,
            float(row_edit["FY2 Change %"]) if row_edit is not None else 0.0,
        ]
        for yi, fy in enumerate(target_fys):
            info = cache.get((rec["key"], fy))
            if not info:
                continue
            pred, method_applied = ensure_forecast(info, selected_method, candidates)
            change = changes[yi]
            month_factors = np.array([variable_multiplier(fy, dt, assumptions) for dt in info["target_months"]], dtype=float)
            adjusted = np.maximum(0, np.asarray(pred, dtype=float) * (1 + change/100.0) * month_factors)
            forecast_values[(rec["key"], fy)] = adjusted
            avg_variable_effect = float(np.mean(month_factors) - 1.0) * 100.0
            metrics = info["metrics"]
            metric_row = metric_for_method(info, method_applied, candidates)
            annual_base = float(np.sum(pred))
            annual_adjusted = float(np.sum(adjusted))
            price = rec["price_inr_pc"] if np.isfinite(rec["price_inr_pc"]) else rec["sales_price"]
            business_value = annual_adjusted * price if np.isfinite(price) else np.nan
            summary_rows.append({
                "Forecast FY":fy,"Plant":rec["plant"],"Customer":rec["customer"],
                "Model":rec["model"],"Vehicle":rec["vehicle"],"Part No.":rec["part_no"],
                "SAP Code":rec["sap_code"],"Product":rec["product"],"Segment":rec["segment"],
                "Material":rec["material"],"Specification":rec["specification"],"Weight kg/pc":rec["weight_kg"],
                "Sales Price INR/pc":rec["sales_price"],"Price INR/pc":rec["price_inr_pc"],
                "Last Month Qty":info["last"],"Last 3M Mean":info["mean3"],"Last 12M Mean":info["mean12"],
                "Demand Pattern":info["pattern"],"Selected Forecast Method":method_applied,
                "WAPE %":metric_row.get("WAPE %",np.nan),"RMSE":metric_row.get("RMSE",np.nan),
                "Forecast Bias %":metric_row.get("Bias %",np.nan),"Part Change %":change,
                "Average External Variables Effect %":avg_variable_effect,
                "Base Annual Forecast Qty":annual_base,"Adjusted Annual Forecast Qty":annual_adjusted,
                "Forecast Business Value INR":business_value,"Data Basis":"Existing monthly plan baseline; not actual demand"
            })
            for _, mrow in metrics.iterrows():
                comparison_rows.append({
                    "Forecast FY":fy,"Plant":rec["plant"],"Customer":rec["customer"],
                    "Part No.":rec["part_no"],"Product":rec["product"],
                    "Demand Pattern":info["pattern"],"Candidate Method":mrow["Method"],
                    "WAPE %":mrow["WAPE %"],"RMSE":mrow["RMSE"],"Bias %":mrow["Bias %"],
                    "Validation":mrow["Validation"],"Selected Method":method_applied
                })
            if method_applied not in set(metrics["Method"].astype(str)):
                comparison_rows.append({
                    "Forecast FY":fy,"Plant":rec["plant"],"Customer":rec["customer"],
                    "Part No.":rec["part_no"],"Product":rec["product"],
                    "Demand Pattern":info["pattern"],"Candidate Method":method_applied,
                    "WAPE %":metric_row.get("WAPE %",np.nan),"RMSE":metric_row.get("RMSE",np.nan),
                    "Bias %":metric_row.get("Bias %",np.nan),"Validation":metric_row.get("Validation",""),
                    "Selected Method":method_applied
                })

    # Add cloned plant sheets. Original tabs remain untouched.
    for fy in target_fys:
        for plant in month_map:
            source_ws = wb[plant]
            new_title = f"FCST {fy.replace('FY','FY-')} {plant}"
            new_title = new_title[:31]
            if new_title in wb.sheetnames:
                del wb[new_title]
            out_ws = wb.copy_worksheet(source_ws)
            out_ws.title = new_title
            out_ws.sheet_properties.tabColor = "176B87"
            out_ws["C2"] = f"MODEL-SELECTED FORECAST — {fy} — {plant}"
            year_months = [(col, dt) for col, dt, label in month_map[plant] if label == fy]
            headers = ["Forecast Method","Demand Pattern","Last Month Qty","Last 3M Mean","Last 12M Mean",
                       "WAPE %","RMSE","Bias %","Change %","Base Annual Qty","Adjusted Annual Qty",
                       "Price INR/pc","Forecast Business Value INR","Average External Variables Effect %","Data Basis","Forecast FY"]
            extra_start = max(out_ws.max_column + 1, 133)
            style_new_sheet(out_ws, extra_start, headers)
            for rec in records:
                if rec["plant"] != plant:
                    continue
                info = cache.get((rec["key"], fy))
                if not info:
                    continue
                row_edit = edited_by_key.get(rec["key"])
                selected_method = str(row_edit["Selected Method"]) if row_edit is not None else info["default_method"]
                yi = target_fys.index(fy)
                change = float(row_edit["FY1 Change %"] if yi == 0 else row_edit["FY2 Change %"]) if row_edit is not None else 0.0
                pred, method_applied = ensure_forecast(info, selected_method, candidates)
                month_factors = np.array([variable_multiplier(fy, dt, assumptions) for dt in info["target_months"]], dtype=float)
                adjusted = np.maximum(0, np.asarray(pred, dtype=float) * (1 + change/100.0) * month_factors)
                forecast_values[(rec["key"], fy)] = adjusted
                # Replace only this year's monthly cells on the cloned worksheet.
                for (col, dt), qty in zip(year_months, adjusted):
                    out_ws.cell(rec["source_row"], col).value = float(qty)
                metric = metric_for_method(info, method_applied, candidates)
                annual_base = float(np.sum(pred))
                annual_adjusted = float(np.sum(adjusted))
                price = rec["price_inr_pc"] if np.isfinite(rec["price_inr_pc"]) else rec["sales_price"]
                value = annual_adjusted * price if np.isfinite(price) else np.nan
                avg_var_effect = float(np.mean(month_factors) - 1.0) * 100.0
                vals = [method_applied,info["pattern"],info["last"],info["mean3"],info["mean12"],
                        metric.get("WAPE %",np.nan),metric.get("RMSE",np.nan),metric.get("Bias %",np.nan),
                        change,annual_base,annual_adjusted,price,value,avg_var_effect,
                        "Existing monthly plan baseline; not actual demand",fy]
                for j, v in enumerate(vals):
                    cell = out_ws.cell(rec["source_row"], extra_start+j)
                    cell.value = None if isinstance(v, float) and not np.isfinite(v) else v
                    cell._style = copy(out_ws.cell(rec["source_row"], 132)._style)
                    cell.alignment = copy(out_ws.cell(rec["source_row"], 132).alignment)
                    if j in [2,3,4,5,7,8]:
                        cell.number_format = "0.0"
                    if j in [9,10,11,12]:
                        cell.number_format = '#,##0.00'
            out_ws.freeze_panes = "AJ11"
            out_ws.sheet_view.showGridLines = False

    summary_df = pd.DataFrame(summary_rows)
    comparison_df = pd.DataFrame(comparison_rows)
    add_table_sheet(wb, "Forecast Summary", summary_df)
    add_table_sheet(wb, "Model Comparison", comparison_df)

    # Consolidated comparison: one row per part, historical months plus the selected 24 forecast months.
    target_months_all = sorted({dt for rec in records for fy in target_fys
                                for dt in (cache.get((rec["key"], fy), {}).get("target_months", []))})
    first_target_month = min(target_months_all) if target_months_all else None
    historical_months = sorted({dt for rec in records for dt in rec["series"].keys()
                                if first_target_month is not None and dt < first_target_month})
    forecast_months = sorted(set(target_months_all))
    part_rows = []
    for rec in records:
        row_edit = edited_by_key.get(rec["key"])
        selected_method = str(row_edit["Selected Method"]) if row_edit is not None else cache[(rec["key"], target_fys[0])]["default_method"]
        part_row = {
            "Plant":rec["plant"], "Customer":rec["customer"], "Model":rec["model"], "Vehicle":rec["vehicle"],
            "Part No.":rec["part_no"], "SAP Code":rec["sap_code"], "Product":rec["product"],
            "Segment":rec["segment"], "Material":rec["material"], "Specification":rec["specification"],
            "Weight kg/pc":rec["weight_kg"], "Price INR/pc":rec["price_inr_pc"],
            "Source Data Basis":"Uploaded monthly plan / historical baseline; not actual demand",
        }
        for dt in historical_months:
            val = rec["series"].get(dt, np.nan)
            part_row[dt.strftime("%b-%y")] = float(val) if np.isfinite(val) else None
        variable_labels = []
        applied_models = []
        for yi, fy in enumerate(target_fys):
            info = cache.get((rec["key"], fy))
            if info is None:
                continue
            pred, method_applied = ensure_forecast(info, selected_method, candidates)
            change = float(row_edit["FY1 Change %"] if yi == 0 else row_edit["FY2 Change %"]) if row_edit is not None else 0.0
            for dt, qty in zip(info["target_months"], forecast_values.get((rec["key"], fy), pred)):
                part_row[dt.strftime("%b-%y")] = float(qty)
            if method_applied not in applied_models:
                applied_models.append(method_applied)
            cfg = assumptions.get(fy, {})
            for key, title, desc, scope in VARIABLE_DEFINITIONS:
                if cfg.get(key + "_enabled") and abs(float(cfg.get(key + "_pct", 0.0) or 0.0)) > 1e-12:
                    months = cfg.get(key + "_months", MONTH_NAMES) if scope == "months" else MONTH_NAMES
                    month_text = f" ({','.join(months)})" if scope == "months" else ""
                    variable_labels.append(f"{fy} {title}: {float(cfg.get(key + '_pct',0.0)):+g}%{month_text}")
            if abs(change) > 1e-12:
                variable_labels.append(f"{fy} Part change: {change:+g}%")
        # Keep these as the final columns as requested.
        part_row["Variables Applied"] = "; ".join(dict.fromkeys(variable_labels)) if variable_labels else "None"
        part_row["Model Used"] = " / ".join(applied_models) if applied_models else selected_method
        part_rows.append(part_row)
    partwise_df = pd.DataFrame(part_rows)
    partwise_ws = add_table_sheet(wb, "Part-wise Monthly Forecast", partwise_df)
    partwise_ws.freeze_panes = "M2"
    for col_idx, col_name in enumerate(partwise_df.columns, start=1):
        if col_name in {dt.strftime("%b-%y") for dt in historical_months + forecast_months}:
            partwise_ws.column_dimensions[get_column_letter(col_idx)].width = 12
            for cells in partwise_ws.iter_cols(min_col=col_idx, max_col=col_idx, min_row=2):
                for cell in cells:
                    cell.number_format = '#,##0.00'
    partwise_ws.auto_filter.ref = partwise_ws.dimensions

    # One selected-part visual in Excel; the same selection is graphed in the app.
    selected_rec = next((r for r in records if r["key"] == selected_part_key), None)
    if selected_rec is None and records:
        selected_rec = records[0]
    if selected_rec is not None:
        chart_ws = wb["Selected Part Trend"] if "Selected Part Trend" in wb.sheetnames else wb.create_sheet("Selected Part Trend")
        chart_ws.delete_rows(1, chart_ws.max_row)
        chart_ws.sheet_view.showGridLines = False
        chart_ws.append(["Month", "Series", "Quantity"])
        selected_info = []
        for dt in historical_months:
            val = selected_rec["series"].get(dt, np.nan)
            if np.isfinite(val):
                selected_info.append((dt, "Historical", float(val)))
        for fy in target_fys:
            info = cache.get((selected_rec["key"], fy))
            if info:
                yi = target_fys.index(fy)
                row_edit = edited_by_key.get(selected_rec["key"])
                method = str(row_edit["Selected Method"]) if row_edit is not None else info["default_method"]
                pred, _ = ensure_forecast(info, method, candidates)
                change = float(row_edit["FY1 Change %"] if yi == 0 else row_edit["FY2 Change %"]) if row_edit is not None else 0.0
                factors = [variable_multiplier(fy, dt, assumptions) for dt in info["target_months"]]
                adjusted = np.maximum(0, np.asarray(pred) * (1 + change/100.0) * np.asarray(factors))
                selected_info.extend((dt, "Forecast", float(qty)) for dt, qty in zip(info["target_months"], adjusted))
        for dt, series_type, qty in sorted(selected_info, key=lambda x: x[0]):
            chart_ws.append([dt.to_pydatetime(), series_type, qty])
        for cell in chart_ws[1]:
            cell.fill = PatternFill("solid", fgColor="176B87")
            cell.font = Font(bold=True, color="FFFFFF")
        chart_ws.column_dimensions["A"].width = 16
        chart_ws.column_dimensions["B"].width = 16
        chart_ws.column_dimensions["C"].width = 18
        chart_ws.freeze_panes = "A2"
        from openpyxl.chart import LineChart, Reference
        line = LineChart()
        line.title = f"Monthly demand — {selected_rec['part_no']}"
        line.y_axis.title = "Quantity"
        line.x_axis.title = "Month"
        line.style = 13
        line.height = 9
        line.width = 20
        data_ref = Reference(chart_ws, min_col=3, min_row=1, max_row=chart_ws.max_row)
        cats_ref = Reference(chart_ws, min_col=1, min_row=2, max_row=chart_ws.max_row)
        line.add_data(data_ref, titles_from_data=True)
        line.set_categories(cats_ref)
        chart_ws.add_chart(line, "E2")
        chart_ws["E20"] = f"Part: {selected_rec['part_no']} | Plant: {selected_rec['plant']}"
        chart_ws["E21"] = "Historical values are from the uploaded workbook; forecast values include part changes and enabled external variables."

    add_table_sheet(wb, "Forecast Glossary", build_glossary_df())

    controls = [
        ("Application Version", APP_VERSION),
        ("Source Workbook", filename),
        ("Forecast Years", ", ".join(target_fys)),
        ("Default UI Method", global_method),
        ("Candidate Methods Compared", ", ".join(candidates)),
        ("Forecast Horizon", "12 months per selected fiscal year"),
        ("Source Plant Sheets", ", ".join(month_map.keys())),
        ("Data Basis", "The provided plant sheets label AJ:CQ quantities as Plan. They are not treated as actual sales."),
        ("Historical Actual Coverage", "Sheet1 (2) contains FY24-25 customer-level actuals only; no part-level monthly actual series was detected."),
        ("Backtest Interpretation", "Backtest is performed against the prior monthly series from this workbook. Because the plant-level input is labelled Plan, this is plan-baseline validation, not actual-demand accuracy."),
        ("Workbook Preservation", "All original worksheets are retained. New forecast plant sheets are copies; only selected forecast-year monthly cells in the copies are replaced."),
        ("Run Timestamp", datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
    ]
    control_df = pd.DataFrame(controls, columns=["Setting","Value"])
    add_table_sheet(wb, "Forecast Controls", control_df)
    add_table_sheet(wb, "Forecast Assumptions", assumptions_to_rows(assumptions, target_fys))
    try:
        wb.calculation.fullCalcOnLoad = True
        wb.calculation.forceFullCalc = True
        wb.calculation.calcMode = "auto"
    except Exception:
        pass
    # Put the consolidated part-wise monthly sheet at the end for quick access.
    # Use the worksheet object created above instead of looking it up by title;
    # this avoids a KeyError if a workbook title is altered or normalised.
    if partwise_ws in wb._sheets:
        wb._sheets.remove(partwise_ws)
        wb._sheets.append(partwise_ws)
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output, summary_df, comparison_df

def main():
    st.title("Harmony Manufacturing Forecast Engine")
    st.caption("Manufacturing workbook-first forecasting • original layout preserved • part / plant / customer level • editable two-year plan")
    uploaded = st.file_uploader("Upload your manufacturing sales / planning workbook", type=["xlsx","xlsm"], help="Use the workbook with Manesar and Gujrat tabs, part numbers, monthly columns and pricing.")
    if uploaded is None:
        st.info("Upload the manufacturing workbook to begin. The app will detect the plant tabs, part numbers, monthly periods, pricing and existing plan columns automatically.")
        st.markdown("**What happens automatically**")
        st.write("Detect plant worksheets → read part number / customer / product / pricing → calculate last month, 3-month mean and 12-month mean → compare forecast methods → apply per-part percentage changes → export a copy of your workbook with forecast sheets.")
        st.caption("No manual column mapping. Original worksheets are retained.")
        return

    raw = uploaded.getvalue()
    try:
        keep_vba = uploaded.name.lower().endswith(".xlsm")
        wb_formula = openpyxl.load_workbook(io.BytesIO(raw), data_only=False, keep_vba=keep_vba)
        wb_values = openpyxl.load_workbook(io.BytesIO(raw), data_only=True, keep_vba=keep_vba)
        plants = detect_plant_sheets(wb_values)
        if not plants:
            st.error("This workbook does not match the expected manufacturing layout. I could not find plant tabs with Part No. in column F and monthly dates in row 6.")
            st.stop()
        records, month_map, basis_by_year = read_plant_records(wb_formula, wb_values, plants)
        all_fys = sorted({fy for months in month_map.values() for _,_,fy in months})
        if len(records) == 0 or len(all_fys) < 2:
            st.error("No part-level monthly data or fewer than two fiscal years were detected.")
            st.stop()
    except Exception as e:
        st.error(f"Could not read the workbook: {e}")
        st.stop()

    current_fy_start = datetime.now().year if datetime.now().month >= 4 else datetime.now().year - 1
    future_fys = [fy for fy in all_fys if int("20"+fy[2:4]) >= current_fy_start+1]
    defaults = future_fys[:2] if len(future_fys) >= 2 else all_fys[-2:]
    if len(defaults) < 2:
        defaults = all_fys[:2]

    st.success(f"Detected {len(plants)} plant sheets and {len(records):,} part / plant records across {len(all_fys)} fiscal years.")
    st.caption(f"Plant sheets detected: {', '.join(plants)}")

    st.markdown("### 1. Select the two forecast years")
    c1,c2,c3 = st.columns([1,1,2])
    with c1:
        fy1 = st.selectbox("Forecast Year 1", all_fys, index=all_fys.index(defaults[0]), key="fy1")
    with c2:
        year2_options = [x for x in all_fys if x != fy1]
        def2 = defaults[1] if defaults[1] in year2_options else (year2_options[-1] if year2_options else fy1)
        fy2 = st.selectbox("Forecast Year 2", year2_options, index=year2_options.index(def2), key="fy2")
    with c3:
        global_method = st.selectbox("Forecast method selection", ["Auto - Best Backtest"] + METHODS, index=0, help="Auto selects the best candidate by chronological holdout WAPE for each part. You can override the selected method part-by-part below.")
    if global_method == "Auto - Best Backtest":
        st.info("**Auto - Best Backtest** — Tests the eligible candidate methods on the latest chronological holdout months for each part, then selects the best validated method. Sparse-demand parts are prioritised for Croston SBA when available.")
    else:
        st.info(f"**{global_method}** — {MODEL_DESCRIPTIONS[global_method][0]} {MODEL_DESCRIPTIONS[global_method][1]}")
    with st.expander("All forecast model names and meanings", expanded=True):
        st.dataframe(pd.DataFrame([{"Model":m, "What it does":MODEL_DESCRIPTIONS[m][0], "When it is useful":MODEL_DESCRIPTIONS[m][1]} for m in METHODS]), use_container_width=True, hide_index=True)
    candidates = st.multiselect("Methods to compare in automatic selection", METHODS, default=["Seasonal Naive","ETS / Holt-Winters","Croston SBA"], help="These methods are compared on the latest available holdout months before each forecast year.")
    if not candidates:
        st.warning("Select at least one method for automatic comparison.")
        st.stop()
    target_fys = [fy1, fy2]

    st.markdown("### 2. External variables & business assumptions")
    st.caption("Enable only the factors relevant to your business. Disabled factors have zero impact. Month-specific factors apply only to the months you select.")
    assumptions = {}
    for fy in target_fys:
        with st.expander(f"Business assumptions — {fy}", expanded=True):
            cfg = {}
            for key, title, desc, scope in VARIABLE_DEFINITIONS:
                st.markdown(f"**{title}**")
                st.caption(desc)
                c_enable, c_pct, c_months = st.columns([1, 1, 2])
                enabled = c_enable.checkbox("Apply", value=False, key=f"{fy}_{key}_enabled")
                pct = c_pct.number_input("Expected change %", min_value=-100.0, max_value=500.0, value=0.0, step=1.0, key=f"{fy}_{key}_pct", disabled=not enabled)
                cfg[key + "_enabled"] = enabled
                cfg[key + "_pct"] = pct
                if scope == "months":
                    months = c_months.multiselect("Apply in months", MONTH_NAMES, default=[], key=f"{fy}_{key}_months", disabled=not enabled)
                    cfg[key + "_months"] = months
                else:
                    cfg[key + "_months"] = MONTH_NAMES
            assumptions[fy] = cfg
    assumptions_df = assumptions_to_rows(assumptions, target_fys)
    active_assumptions = assumptions_df[assumptions_df["Enabled"]]
    if active_assumptions.empty:
        st.info("No external variables enabled. Forecasts use the selected model and per-part percentage changes only.")
    else:
        st.dataframe(active_assumptions, use_container_width=True, hide_index=True)

    # The workbook's plant sheets mark the monthly data as Plan, not actuals.
    source_basis = "Plan" if any("Plan" in str(x) for vals in basis_by_year.values() for x in vals) else "Unspecified / mixed"
    if source_basis == "Plan":
        st.warning("Important: the plant-level monthly values in this workbook are labelled Plan. They are used as the forecasting baseline, not represented as historical actual demand. FY24-25 actuals in Sheet1 (2) are customer-level only, not part-level.")
    with st.expander("Detected data and source interpretation"):
        st.write("Part-level identifiers: Customer, Model, Vehicle, Part No., SAP Code, Product, Segment, Material, Specification, Weight and pricing fields.")
        st.write("Monthly input range: AJ:CQ, grouped by fiscal year from the dates in row 6.")
        st.write("The original workbook sheets will remain unchanged. Forecast values are written into copied plant sheets for the two selected years.")

    with st.spinner("Comparing methods and preparing part-level forecasts..."):
        cache = make_forecast_cache(records, month_map, target_fys, candidates)

    # One row per plant + source row; method default comes from first selected forecast year.
    ui_rows = []
    for rec in records:
        info1 = cache.get((rec["key"], fy1))
        info2 = cache.get((rec["key"], fy2))
        if not info1 or not info2:
            continue
        method = global_method if global_method != "Auto - Best Backtest" else info1["default_method"]
        pred1, method1 = ensure_forecast(info1, method, candidates)
        pred2, method2 = ensure_forecast(info2, method, candidates)
        if method1 == "Existing Plan (NPI)" or method2 == "Existing Plan (NPI)":
            method = "Existing Plan (NPI)"
        ui_rows.append({
            "Part Key":rec["key"],"Plant":rec["plant"],"Customer":rec["customer"],"Part No.":rec["part_no"],
            "Model":rec["model"],"Vehicle":rec["vehicle"],"SAP Code":rec["sap_code"],"Product":rec["product"],
            "Segment":rec["segment"],"Weight kg/pc":rec["weight_kg"],"Sales Price INR/pc":rec["sales_price"],
            "Price INR/pc":rec["price_inr_pc"],"Last Month Qty":info1["last"],"Last 3M Mean":info1["mean3"],
            "Last 12M Mean":info1["mean12"],"Demand Pattern":info1["pattern"],"Selected Method":method,
            "FY1 Change %":0.0,"FY2 Change %":0.0,
            f"{fy1} Base Qty":float(np.sum(pred1)),f"{fy2} Base Qty":float(np.sum(pred2)),
        })
    ui_df = pd.DataFrame(ui_rows).set_index("Part Key", drop=False)
    if ui_df.empty:
        st.error("No records had usable data for both selected forecast years.")
        st.stop()

    st.markdown("### 3. Part-level forecast controls")
    st.caption("Select a model and enter a part-level change for each forecast year. Last 3M Mean is the average monthly quantity across the latest 3 source months; Last 12M Mean is the average monthly quantity across the latest 12 source months. These are reference indicators, not the forecast. Changes are applied to every forecast month.")
    visible_cols = ["Plant","Customer","Part No.","Model","Vehicle","SAP Code","Product","Segment","Weight kg/pc",
                    "Sales Price INR/pc","Price INR/pc","Last 3M Mean","Last 12M Mean",
                    "Demand Pattern","Selected Method","FY1 Change %","FY2 Change %",f"{fy1} Base Qty",f"{fy2} Base Qty"]
    edited = st.data_editor(
        ui_df[visible_cols],
        use_container_width=True,
        hide_index=True,
        num_rows="fixed",
        column_config={
            "Selected Method":st.column_config.SelectboxColumn("Selected Method",options=METHODS,required=True,
                help="Choose the forecasting approach for this part. Full model meanings are shown above in the model guide."),
            "FY1 Change %":st.column_config.NumberColumn(f"{fy1} Change %",min_value=-100.0,max_value=500.0,step=1.0,format="%.1f"),
            "FY2 Change %":st.column_config.NumberColumn(f"{fy2} Change %",min_value=-100.0,max_value=500.0,step=1.0,format="%.1f"),
            "Last 3M Mean":st.column_config.NumberColumn(format="%.2f",disabled=True),
            "Last 12M Mean":st.column_config.NumberColumn(format="%.2f",disabled=True),
            f"{fy1} Base Qty":st.column_config.NumberColumn(format="%.2f",disabled=True),
            f"{fy2} Base Qty":st.column_config.NumberColumn(format="%.2f",disabled=True),
            "Demand Pattern":st.column_config.TextColumn(disabled=True),
        },
        key="part_forecast_editor"
    )
    edited = edited.copy()
    edited["Part Key"] = ui_df.loc[edited.index, "Part Key"].values

    # Live adjusted summary: annual forecast versus comparable source baseline.
    live_rows = []
    monthly_rows = []
    for _, row in edited.iterrows():
        key = row["Part Key"]
        rec = next(r for r in records if r["key"] == key)
        method = str(row["Selected Method"])
        for yi, fy in enumerate(target_fys):
            info = cache[(key, fy)]
            pred, method_applied = ensure_forecast(info, method, candidates)
            change = float(row["FY1 Change %"] if yi == 0 else row["FY2 Change %"])
            month_factors = np.array([variable_multiplier(fy, dt, assumptions) for dt in info["target_months"]], dtype=float)
            adjusted = np.maximum(0, np.asarray(pred) * (1 + change/100.0) * month_factors)
            avg_variable_effect = float(np.mean(month_factors) - 1.0) * 100.0
            metric = metric_for_method(info, method_applied, candidates)
            price = rec["price_inr_pc"] if np.isfinite(rec["price_inr_pc"]) else rec["sales_price"]

            baseline = np.asarray(info["target_plan"], dtype=float)
            valid = np.isfinite(baseline)
            baseline_total = float(np.sum(baseline[valid])) if valid.any() else np.nan
            forecast_comparable = float(np.sum(adjusted[valid])) if valid.any() else np.nan
            variance_qty = forecast_comparable - baseline_total if valid.any() else np.nan
            variance_pct = variance_qty / baseline_total * 100.0 if valid.any() and abs(baseline_total) > 1e-12 else np.nan
            if not valid.any():
                variance_text = "No comparable source values are available for these forecast months."
            elif variance_pct > 0.05:
                variance_text = f"Forecast is {variance_pct:.1f}% ({variance_qty:,.0f} units) above the uploaded {source_basis.lower()} baseline."
            elif variance_pct < -0.05:
                variance_text = f"Forecast is {abs(variance_pct):.1f}% ({abs(variance_qty):,.0f} units) below the uploaded {source_basis.lower()} baseline."
            elif not np.isfinite(variance_pct):
                variance_text = f"Comparable {source_basis.lower()} total is zero; percentage variance is not meaningful."
            else:
                variance_text = f"Forecast is broadly in line with the uploaded {source_basis.lower()} baseline."

            driver_text = (
                f"Model: {method_applied}; part adjustment: {change:+.1f}%; "
                f"average enabled business assumptions: {avg_variable_effect:+.1f}%."
            )
            variance_text = variance_text + " " + driver_text

            live_rows.append({
                "Forecast FY":fy,"Plant":rec["plant"],"Customer":rec["customer"],"Part No.":rec["part_no"],
                "Product":rec["product"],"Forecast Method":method_applied,
                "Last 3M Mean (monthly)":info["mean3"],"Last 12M Mean (monthly)":info["mean12"],
                f"{source_basis} Baseline Qty":baseline_total,
                "Forecast Qty (annual)":float(np.sum(adjusted)),
                "Variance Qty":variance_qty,"Variance %":variance_pct,"Variance explanation":variance_text,
                "WAPE % (backtest)":metric.get("WAPE %",np.nan),
                "RMSE (units)":metric.get("RMSE",np.nan),"Bias % (backtest)":metric.get("Bias %",np.nan),
                "Part Change %":change,"External Variables Effect % (avg)":avg_variable_effect,
                "Price INR/pc":price,
                "Forecast Business Value INR":float(np.sum(adjusted))*price if np.isfinite(price) else np.nan,
            })
            for dt, base_qty, adj_qty, var_factor, baseline_qty in zip(info["target_months"], pred, adjusted, month_factors, baseline):
                has_baseline = np.isfinite(baseline_qty)
                month_var_qty = float(adj_qty - baseline_qty) if has_baseline else np.nan
                month_var_pct = month_var_qty / float(baseline_qty) * 100.0 if has_baseline and abs(float(baseline_qty)) > 1e-12 else np.nan
                monthly_rows.append({
                    "Forecast FY":fy,"Month":dt.strftime("%b-%Y"),"Plant":rec["plant"],
                    "Customer":rec["customer"],"Part No.":rec["part_no"],"Product":rec["product"],
                    "Forecast Method":method_applied,
                    f"{source_basis} Baseline Qty":float(baseline_qty) if has_baseline else np.nan,
                    "Forecast Qty":float(adj_qty),
                    "Variance Qty":month_var_qty,"Variance %":month_var_pct,
                    "Part Change %":change,"External Variables Effect %":float((var_factor-1)*100),
                    "Base Model Qty":float(base_qty),
                })

    live_df = pd.DataFrame(live_rows)
    total_forecast = live_df.groupby("Forecast FY")["Forecast Qty (annual)"].sum()
    m1,m2,m3,m4 = st.columns(4)
    m1.metric(f"{fy1} forecast volume", f"{total_forecast.get(fy1,0):,.0f}")
    m2.metric(f"{fy2} forecast volume", f"{total_forecast.get(fy2,0):,.0f}")
    m3.metric("Part / plant combinations", f"{len(ui_df):,}")
    m4.metric("Forecast methods used", f"{len(set(edited['Selected Method']))}")

    st.markdown("### 4. Forecast preview — annual variance and monthly quantities")
    st.caption("Forecast Qty is the model's 12-month total after your part-level change and enabled business assumptions. The month-by-month quantities are shown below.")
    st.markdown("**How to read the variance**")
    st.write(
        f"Variance Qty = Forecast Qty − {source_basis} Baseline Qty. "
        f"Variance % = Variance Qty ÷ {source_basis} Baseline Qty × 100. "
        "Positive means the forecast is higher; negative means it is lower. "
        "The comparison uses only matching forecast months where source values exist."
    )
    st.markdown("**Forecast adjustment formula**")
    st.code(
        "Adjusted monthly forecast = Base model forecast × (1 + Part Change % / 100) "
        "× applicable external-variable multipliers\n"
        "Annual forecast = Sum of the 12 adjusted monthly forecasts",
        language="text"
    )
    st.markdown("**Backtest metric formulas — historical validation, not future variance**")
    st.write(
        "WAPE % = Σ|Baseline − Backtest forecast| ÷ Σ|Baseline| × 100  •  "
        "Bias % = Σ(Backtest forecast − Baseline) ÷ Σ|Baseline| × 100  •  "
        "RMSE = √[Mean((Baseline − Backtest forecast)²)]."
    )
    st.caption(
        "WAPE shows the size of past holdout errors (lower is better). "
        "Bias shows direction: positive = over-forecast; negative = under-forecast. "
        "RMSE is the error scale in quantity units and penalises large misses. "
        "These are historical backtest measures, not a guarantee of future accuracy."
    )
    if source_basis == "Plan":
        st.warning("Source note: the plant-level monthly series in this workbook is labelled Plan, not Actual. Variance is therefore against the uploaded Plan baseline, not actual sales. Part-level actuals are needed for a true actual-versus-forecast comparison.")
    st.dataframe(live_df, use_container_width=True, hide_index=True)

    monthly_df = pd.DataFrame(monthly_rows)
    st.markdown("**Monthly forecast quantities and variance**")
    st.caption("Each row is one part, plant and calendar month. Forecast Qty is the calculated monthly output, not an annual figure.")
    part_filter_options = ["All parts"] + sorted(monthly_df["Part No."].dropna().astype(str).unique().tolist())
    selected_monthly_part = st.selectbox("Filter monthly detail by part number", part_filter_options, key="monthly_part_filter")
    monthly_view = monthly_df if selected_monthly_part == "All parts" else monthly_df[monthly_df["Part No."].astype(str) == selected_monthly_part]
    st.dataframe(monthly_view, use_container_width=True, hide_index=True)
    st.markdown("### 5. Selected-part monthly visual")
    part_options = {f"{r['part_no']} | {r['plant']} | {r['customer']}": r["key"] for r in records}
    selected_part_label = st.selectbox("Choose a part to view its historical and forecast trend", list(part_options.keys()))
    selected_part_key = part_options[selected_part_label]
    selected_rec_ui = next(r for r in records if r["key"] == selected_part_key)
    first_forecast_month_ui = min(dt for fy in target_fys for dt in cache[(selected_part_key, fy)]["target_months"])
    hist_plot = [{"Month":dt, "Type":"Historical / Source", "Quantity":float(v)} for dt,v in sorted(selected_rec_ui["series"].items()) if dt < first_forecast_month_ui and np.isfinite(v)]
    fc_plot = []
    selected_edit_row = edited[edited["Part Key"] == selected_part_key].iloc[0]
    for yi, fy in enumerate(target_fys):
        selected_info = cache[(selected_part_key, fy)]
        selected_method = str(selected_edit_row["Selected Method"])
        selected_pred, _ = ensure_forecast(selected_info, selected_method, candidates)
        selected_change = float(selected_edit_row["FY1 Change %"] if yi == 0 else selected_edit_row["FY2 Change %"])
        selected_factors = [variable_multiplier(fy, dt, assumptions) for dt in selected_info["target_months"]]
        selected_adjusted = np.maximum(0, np.asarray(selected_pred) * (1 + selected_change/100.0) * np.asarray(selected_factors))
        fc_plot.extend({"Month":dt, "Type":"Forecast", "Quantity":float(qty)} for dt,qty in zip(selected_info["target_months"], selected_adjusted))
    plot_df = pd.DataFrame(hist_plot + fc_plot).sort_values("Month")
    if not plot_df.empty:
        st.line_chart(plot_df.pivot_table(index="Month", columns="Type", values="Quantity", aggfunc="first"))
    with st.expander("Technical terms, model names, patterns and accuracy metrics — meanings", expanded=True):
        st.dataframe(build_glossary_df(), use_container_width=True, hide_index=True)
        st.caption("WAPE: lower is better. RMSE: lower is better and is measured in quantity units. Bias: positive means over-forecast; negative means under-forecast. Validation uses a chronological holdout.")
    with st.expander("Model comparison details by part and forecast year"):
        comparison_preview = []
        for rec in records:
            for fy in target_fys:
                info = cache.get((rec["key"], fy))
                if info is None:
                    continue
                for _, mr in info["metrics"].iterrows():
                    comparison_preview.append({"Forecast FY":fy,"Plant":rec["plant"],"Customer":rec["customer"],"Part No.":rec["part_no"],**mr.to_dict()})
        st.dataframe(pd.DataFrame(comparison_preview), use_container_width=True, hide_index=True)

    st.markdown("### 6. Export original workbook with forecast sheets")
    with st.spinner("Creating a formatted Excel copy and adding forecast-year plant sheets..."):
        output, summary_df, comparison_df = build_output_workbook(raw, uploaded.name, target_fys, records, month_map, cache, edited, candidates, global_method, assumptions, selected_part_key)
    st.success("Forecast workbook is ready. Your original worksheets are retained; new forecast copies are added for each selected year and plant.")
    st.download_button(
        "Download completed manufacturing forecast workbook",
        data=output,
        file_name=f"Harmony_Manufacturing_Forecast_{fy1}_{fy2}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True
    )
    st.caption("Excel output includes: original tabs unchanged; copied forecast plant tabs with adjusted monthly quantities; Forecast Summary; Model Comparison; Forecast Controls; Forecast Assumptions, Part-wise Monthly Forecast (historical months plus selected 24 forecast months), Selected Part Trend chart, and Forecast Glossary. The final two columns in the part-wise sheet are Variables Applied and Model Used.")

if __name__ == "__main__":
    main()
