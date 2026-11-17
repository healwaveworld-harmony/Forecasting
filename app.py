
import io
import re
import warnings
from datetime import datetime
import numpy as np
import pandas as pd
import streamlit as st

warnings.filterwarnings("ignore")

APP_TITLE = "Harmony Manufacturing Forecast Engine"
HORIZON = 24
DIMENSION_HINTS = {
    "customer": ["customer", "client", "account", "buyer", "customer name"],
    "sku": ["sku", "item code", "item no", "material code", "product code", "part no", "part number"],
    "product": ["product", "item description", "material description", "description", "product name"],
    "plant": ["plant", "factory", "location", "site", "unit", "manufacturing location"],
}
QTY_HINTS = ["actual qty", "actual quantity", "sales qty", "sales quantity", "quantity", "qty", "volume", "units", "demand", "sales", "consumption", "dispatch"]
DATE_HINTS = ["date", "month", "period", "year month", "month year"]
EXCLUDE_WORDS = ["budget", "forecast", "plan", "target", "projection", "variance", "growth", "increment", "capacity"]

st.set_page_config(page_title=APP_TITLE, page_icon="📈", layout="wide")

def norm(s):
    return re.sub(r"[^a-z0-9]+", " ", str(s).strip().lower()).strip()

def score_name(name, hints):
    n = norm(name)
    return max([len(h) for h in hints if h in n] or [0])

def parse_month_label(value):
    """Parse a monthly column heading into a month-start timestamp."""
    if isinstance(value, (pd.Timestamp, datetime)):
        return pd.Timestamp(value).to_period("M").to_timestamp()
    if isinstance(value, (int, float)) and 190001 <= value <= 210012:
        s = str(int(value))
        try:
            return pd.Timestamp(year=int(s[:4]), month=int(s[4:]), day=1)
        except Exception:
            return None
    s = str(value).strip()
    if not s or any(x in norm(s) for x in EXCLUDE_WORDS):
        return None
    # Common Excel headings: Apr-24, April 2024, 2024-04, 04/2024
    for fmt in ("%Y-%m", "%b-%y", "%B-%y", "%b %Y", "%B %Y", "%m/%Y", "%Y/%m", "%Y_%m"):
        try:
            return pd.to_datetime(s, format=fmt).to_period("M").to_timestamp()
        except Exception:
            pass
    try:
        dt = pd.to_datetime(s, errors="coerce", dayfirst=True)
        if pd.notna(dt) and 1990 <= dt.year <= 2100:
            return dt.to_period("M").to_timestamp()
    except Exception:
        pass
    return None

def find_header_and_data(df):
    best = None
    for r in range(min(8, len(df))):
        vals = [str(x) for x in df.iloc[r].tolist()]
        score = sum(score_name(v, sum(DIMENSION_HINTS.values(), []) + QTY_HINTS + DATE_HINTS) > 0 for v in vals)
        score += sum(parse_month_label(v) is not None for v in vals)
        if best is None or score > best[0]:
            best = (score, r)
    if best is None:
        return None, None
    header_row = best[1]
    headers = []
    for i, v in enumerate(df.iloc[header_row].tolist()):
        v = str(v).strip()
        headers.append(v if v and v.lower() != "nan" else f"Column_{i+1}")
    data = df.iloc[header_row + 1:].copy()
    data.columns = headers
    data = data.dropna(how="all")
    return data, header_row

def identify_columns(df):
    cols = list(df.columns)
    dims = {}
    for dim, hints in DIMENSION_HINTS.items():
        candidates = sorted([(score_name(c, hints), c) for c in cols], reverse=True)
        if candidates and candidates[0][0] > 0:
            dims[dim] = candidates[0][1]
    date_col = None
    qty_col = None
    for c in cols:
        if date_col is None and score_name(c, DATE_HINTS) > 0:
            date_col = c
        if qty_col is None and score_name(c, QTY_HINTS) > 0 and not any(x in norm(c) for x in EXCLUDE_WORDS):
            qty_col = c
    return dims, date_col, qty_col

def workbook_sheets(uploaded, filename):
    name = filename.lower()
    if name.endswith(".csv"):
        uploaded.seek(0)
        return {"CSV Upload": pd.read_csv(uploaded, header=None)}
    uploaded.seek(0)
    xls = pd.ExcelFile(uploaded)
    out = {}
    for sheet in xls.sheet_names:
        try:
            out[sheet] = pd.read_excel(xls, sheet_name=sheet, header=None)
        except Exception:
            continue
    return out

def extract_monthly_data(sheets):
    frames, report = [], []
    for sheet_name, raw in sheets.items():
        if raw is None or raw.empty:
            continue
        data, header_row = find_header_and_data(raw)
        if data is None or data.empty:
            continue
        dims, date_col, qty_col = identify_columns(data)
        # Prefer explicit date + quantity long-form data.
        if date_col and qty_col:
            tmp = data.copy()
            tmp["_date"] = pd.to_datetime(tmp[date_col], errors="coerce", dayfirst=True)
            tmp["_qty"] = pd.to_numeric(tmp[qty_col].astype(str).str.replace(",", "", regex=False), errors="coerce")
            tmp = tmp[tmp["_date"].notna() & tmp["_qty"].notna()]
            if not tmp.empty:
                tmp["_month"] = tmp["_date"].dt.to_period("M").dt.to_timestamp()
                for d in ["customer", "sku", "product", "plant"]:
                    tmp[d] = tmp[dims[d]].astype(str).str.strip() if d in dims else "Not specified"
                frames.append(tmp[["customer", "sku", "product", "plant", "_month", "_qty"]].rename(columns={"_month":"month","_qty":"actual_qty"}))
                report.append({"Worksheet":sheet_name,"Header row":header_row+1,"Format":"Long format","Rows imported":len(tmp),"Detected dimensions":", ".join(dims.keys())})
                continue
        # Wide format: one column per month.
        month_cols = []
        for c in data.columns:
            dt = parse_month_label(c)
            if dt is not None and not any(x in norm(c) for x in EXCLUDE_WORDS):
                month_cols.append((c, dt))
        if month_cols:
            tmp = data.copy()
            for d in ["customer", "sku", "product", "plant"]:
                tmp[d] = tmp[dims[d]].astype(str).str.strip() if d in dims else "Not specified"
            # Skip blank/total rows where all identifiers and monthly values are empty.
            pieces = []
            for c, dt in month_cols:
                q = pd.to_numeric(tmp[c].astype(str).str.replace(",", "", regex=False), errors="coerce")
                part = tmp[["customer","sku","product","plant"]].copy()
                part["month"] = dt
                part["actual_qty"] = q
                pieces.append(part)
            long = pd.concat(pieces, ignore_index=True)
            long = long[long["actual_qty"].notna()]
            # Drop obvious total rows.
            id_text = long[["customer","sku","product","plant"]].astype(str).agg(" ".join, axis=1).str.lower()
            long = long[~id_text.str.contains(r"\b(grand total|subtotal|total)\b", regex=True)]
            if not long.empty:
                frames.append(long)
                report.append({"Worksheet":sheet_name,"Header row":header_row+1,"Format":"Monthly columns","Rows imported":len(long),"Detected dimensions":", ".join(dims.keys())})
    if not frames:
        return pd.DataFrame(), pd.DataFrame(report)
    all_data = pd.concat(frames, ignore_index=True)
    all_data["month"] = pd.to_datetime(all_data["month"]).dt.to_period("M").dt.to_timestamp()
    all_data = all_data.groupby(["customer","sku","product","plant","month"], as_index=False)["actual_qty"].sum()
    all_data = all_data.sort_values(["customer","sku","product","plant","month"])
    return all_data, pd.DataFrame(report)

def classify_demand(y):
    y = np.asarray(y, dtype=float)
    y = np.nan_to_num(y, nan=0.0)
    nz = y[y > 0]
    n = len(y)
    if len(nz) <= 2 or (n and np.mean(y == 0) >= 0.45):
        return "Sparse"
    if len(nz) < 6 or n < 6:
        return "New NPI / Limited history"
    mean = np.mean(y)
    if mean <= 0:
        return "Sparse"
    # Month-of-year variation only when at least 18 observations are present.
    if n >= 18:
        month_means = [np.mean(y[i::12]) for i in range(12) if len(y[i::12])]
        if len(month_means) >= 6 and np.std(month_means) / (np.mean(month_means) + 1e-9) > 0.35:
            return "Seasonal"
    x = np.arange(n)
    slope = np.polyfit(x, y, 1)[0]
    relative = slope * max(n-1, 1) / (mean + 1e-9)
    if relative > 0.25:
        return "Growing"
    if relative < -0.25:
        return "Declining"
    return "Stable"

def wape(actual, pred):
    denom = np.sum(np.abs(actual))
    return float(np.sum(np.abs(actual-pred)) / denom * 100) if denom > 0 else np.nan

def bias_pct(actual, pred):
    denom = np.sum(np.abs(actual))
    return float(np.sum(pred-actual) / denom * 100) if denom > 0 else np.nan

def croston_forecast(y, h, alpha=0.1, sba=True):
    y = np.asarray(y, dtype=float)
    nz_idx = np.where(y > 0)[0]
    if len(nz_idx) == 0:
        return np.zeros(h)
    first = nz_idx[0]
    demand = y[first]
    interval = first + 1
    last = first
    for idx in nz_idx[1:]:
        gap = idx - last
        demand = alpha*y[idx] + (1-alpha)*demand
        interval = alpha*gap + (1-alpha)*interval
        last = idx
    rate = demand / max(interval, 1e-9)
    if sba:
        rate *= (1-alpha/2)
    return np.repeat(max(0, rate), h)

def seasonal_naive(y, h):
    y = np.asarray(y, dtype=float)
    if len(y) >= 12:
        return np.array([y[-12 + (i % 12)] for i in range(h)], dtype=float)
    return np.repeat(float(np.mean(y[-min(3,len(y)):])) if len(y) else 0.0, h)

def ets_forecast(y, h):
    try:
        from statsmodels.tsa.holtwinters import ExponentialSmoothing
        arr = np.asarray(y, dtype=float)
        if len(arr) >= 24:
            model = ExponentialSmoothing(arr, trend="add", seasonal="add", seasonal_periods=12, initialization_method="estimated").fit(optimized=True)
        elif len(arr) >= 6:
            model = ExponentialSmoothing(arr, trend="add", initialization_method="estimated").fit(optimized=True)
        else:
            model = ExponentialSmoothing(arr, initialization_method="estimated").fit(optimized=True)
        return np.maximum(0, np.asarray(model.forecast(h), dtype=float))
    except Exception:
        return np.maximum(0, seasonal_naive(y, h))

def ml_forecast(y, h, model_name):
    from sklearn.linear_model import Ridge
    from sklearn.ensemble import RandomForestRegressor, HistGradientBoostingRegressor
    arr = np.asarray(y, dtype=float)
    lags = [1,2,3,6,12]
    def features(series, t):
        f = [series[t-l] if t-l >= 0 else np.nan for l in lags]
        f += [np.sin(2*np.pi*(t%12)/12), np.cos(2*np.pi*(t%12)/12), t]
        return f
    X, target = [], []
    for t in range(max(lags), len(arr)):
        X.append(features(arr, t)); target.append(arr[t])
    if len(X) < 5:
        raise ValueError("Not enough observations for ML")
    X = np.asarray(X); target = np.asarray(target)
    # Fill unavailable lag features with series mean.
    X = np.where(np.isnan(X), np.mean(arr), X)
    if model_name == "Ridge":
        model = Ridge(alpha=2.0)
    elif model_name == "Random Forest":
        model = RandomForestRegressor(n_estimators=80, max_depth=5, min_samples_leaf=2, random_state=42, n_jobs=1)
    else:
        model = HistGradientBoostingRegressor(max_iter=80, max_leaf_nodes=7, l2_regularization=1.0, random_state=42)
    model.fit(X, target)
    hist = list(arr.astype(float))
    out = []
    for step in range(h):
        t = len(hist)
        f = features(np.asarray(hist), t)
        f = np.where(np.isnan(f), np.mean(arr), f).reshape(1,-1)
        val = max(0.0, float(model.predict(f)[0]))
        out.append(val); hist.append(val)
    return np.asarray(out)

def arima_forecast(y, h):
    from statsmodels.tsa.arima.model import ARIMA
    arr = np.asarray(y, dtype=float)
    # A deliberately compact ARIMA specification keeps cloud runtime bounded.
    fitted = ARIMA(arr, order=(1, 1, 1), enforce_stationarity=False, enforce_invertibility=False).fit()
    return np.maximum(0, np.asarray(fitted.forecast(h), dtype=float))

def ensemble_forecast(y, h):
    forecasts = []
    for candidate in ["Seasonal Naive", "ETS / Holt-Winters", "Ridge"]:
        try:
            forecasts.append(model_predict(y, h, candidate))
        except Exception:
            pass
    if not forecasts:
        return np.maximum(0, seasonal_naive(y, h))
    return np.maximum(0, np.mean(np.vstack(forecasts), axis=0))

def model_predict(y, h, name):
    if name == "Ensemble":
        return ensemble_forecast(y, h)
    if name == "ARIMA":
        return arima_forecast(y, h)
    if name == "Seasonal Naive":
        return np.maximum(0, seasonal_naive(y, h))
    if name == "Moving Average":
        avg = float(np.mean(np.asarray(y)[-min(3,len(y)):])) if len(y) else 0
        return np.repeat(max(0, avg), h)
    if name == "ETS / Holt-Winters":
        return ets_forecast(y, h)
    if name == "Croston SBA":
        return croston_forecast(y, h)
    if name in ["Ridge", "Random Forest", "Gradient Boosting"]:
        return ml_forecast(y, h, name)
    raise ValueError(name)

def compete_models(y):
    y = np.asarray(y, dtype=float)
    models = ["Seasonal Naive", "Moving Average", "ETS / Holt-Winters", "Croston SBA"]
    if len(y) >= 18 and np.count_nonzero(y) >= 8:
        models += ["ARIMA", "Ridge", "Random Forest", "Gradient Boosting", "Ensemble"]
    if len(y) < 4:
        return "Moving Average", pd.DataFrame([{"Model":"Moving Average","WAPE %":np.nan,"RMSE":np.nan,"Bias %":np.nan,"Validation":"Insufficient history"}])
    holdout = min(6, max(2, len(y)//5))
    train, test = y[:-holdout], y[-holdout:]
    results = []
    for name in models:
        try:
            pred = model_predict(train, holdout, name)
            results.append({"Model":name,"WAPE %":wape(test,pred),"RMSE":float(np.sqrt(np.mean((test-pred)**2))),"Bias %":bias_pct(test,pred),"Validation":"Rolling-origin holdout"})
        except Exception:
            continue
    if not results:
        return "Moving Average", pd.DataFrame([{"Model":"Moving Average","WAPE %":np.nan,"RMSE":np.nan,"Bias %":np.nan,"Validation":"Fallback"}])
    res = pd.DataFrame(results)
    valid = res[res["WAPE %"].notna()]
    best = valid.sort_values(["WAPE %","RMSE"], ascending=True).iloc[0]["Model"] if not valid.empty else "Moving Average"
    return best, res.sort_values("WAPE %", na_position="last")

def forecast_series(group, h=HORIZON):
    group = group.sort_values("month")
    full_idx = pd.date_range(group["month"].min(), group["month"].max(), freq="MS")
    s = group.set_index("month")["actual_qty"].reindex(full_idx, fill_value=0).astype(float)
    y = s.values
    pattern = classify_demand(y)
    if pattern == "Sparse":
        chosen = "Croston SBA"
    elif len(y) < 4:
        chosen = "Moving Average"
    else:
        chosen, _ = compete_models(y)
    # New NPI / limited history: do not pretend an ML model has been validated.
    if pattern == "New NPI / Limited history":
        chosen = "Moving Average"
    pred = model_predict(y, h, chosen)
    return full_idx, y, pattern, chosen, pred

def make_excel(original_bytes, original_name, historical, forecasts, validation, quality):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        historical.to_excel(writer, sheet_name="Historical Actuals", index=False)
        forecasts.to_excel(writer, sheet_name="Editable 24M Forecast", index=False)
        validation.to_excel(writer, sheet_name="Model Validation", index=False)
        quality.to_excel(writer, sheet_name="Data Quality", index=False)
        # Include source workbook sheets to preserve the uploaded document.
        try:
            if original_name.lower().endswith((".xlsx", ".xlsm")):
                xls = pd.ExcelFile(io.BytesIO(original_bytes))
                for sn in xls.sheet_names:
                    if sn in ["Historical Actuals","Editable 24M Forecast","Model Validation","Data Quality"]:
                        continue
                    try:
                        src = pd.read_excel(xls, sheet_name=sn)
                        safe = re.sub(r"[\[\]\:\*\?\/\\]", "_", sn)[:31] or "Source"
                        if safe not in ["Historical Actuals","Editable 24M Forecast","Model Validation","Data Quality"]:
                            src.to_excel(writer, sheet_name=safe, index=False)
                    except Exception:
                        pass
        except Exception:
            pass
    output.seek(0)
    return output

def main():
    st.title("Harmony Manufacturing Forecast Engine")
    st.caption("Upload your sales workbook. The engine detects the data structure, analyses demand patterns, compares forecasting methods and creates an editable 24-month forecast.")
    st.markdown("### 1. Upload your sales document")
    uploaded = st.file_uploader("Upload Excel or CSV", type=["xlsx","xlsm","xls","csv"], help="Supports monthly columns or transaction-level date and quantity records.")
    if uploaded is None:
        st.info("Start by uploading your workbook. No column mapping or model selection is required.")
        c1,c2,c3 = st.columns(3)
        c1.metric("Forecast horizon","24 months")
        c2.metric("Model selection","Automatic")
        c3.metric("Output","Editable Excel")
        with st.expander("What the engine detects automatically"):
            st.write("Customer, SKU, product, plant, monthly actual quantities, historical periods and common budget/forecast labels. It classifies demand as seasonal, growing, declining, sparse, new NPI/limited history or stable.")
            st.write("It compares benchmark, ETS/Holt-Winters, Croston SBA and selected machine-learning models when there is enough history. Validation includes WAPE, RMSE and forecast bias.")
        return

    raw_bytes = uploaded.getvalue()
    with st.spinner("Reading workbook and identifying actual sales history..."):
        try:
            sheets = workbook_sheets(io.BytesIO(raw_bytes), uploaded.name)
            hist, ingestion = extract_monthly_data(sheets)
        except Exception as e:
            st.error(f"Could not read this document: {e}")
            st.stop()
    if hist.empty:
        st.error("I couldn't confidently identify monthly actual quantities. Please check that the workbook contains monthly date columns (for example Apr-24) or transaction Date and Quantity columns. Budget/forecast columns are intentionally excluded.")
        if not ingestion.empty:
            st.dataframe(ingestion, use_container_width=True)
        st.stop()

    # Data quality and source summary
    hist = hist.dropna(subset=["month","actual_qty"])
    hist["actual_qty"] = pd.to_numeric(hist["actual_qty"], errors="coerce").fillna(0).clip(lower=0)
    hist = hist.groupby(["customer","sku","product","plant","month"], as_index=False)["actual_qty"].sum()
    keys = ["customer","sku","product","plant"]
    groups = list(hist.groupby(keys, dropna=False))
    max_groups = 1500
    truncated = len(groups) > max_groups
    if truncated:
        groups = sorted(groups, key=lambda kv: kv[1]["actual_qty"].sum(), reverse=True)[:max_groups]
    st.success(f"Workbook processed: {len(hist):,} monthly records across {len(groups):,} customer / SKU / product / plant combinations.")
    if truncated:
        st.warning(f"The workbook contains more than {max_groups:,} combinations. The highest-volume {max_groups:,} combinations were forecast in this run to keep Streamlit Cloud responsive.")
    if not ingestion.empty:
        with st.expander("Automatic workbook detection"):
            st.dataframe(ingestion, use_container_width=True)

    forecast_rows, validation_rows, hist_rows, quality_rows = [], [], [], []
    progress = st.progress(0, text="Analysing demand patterns...")
    for i, (key, g) in enumerate(groups):
        idx, y, pattern, model, pred = forecast_series(g)
        customer, sku, product, plant = key
        group_id = {"customer":customer,"sku":sku,"product":product,"plant":plant}
        for dt, val in zip(idx, y):
            hist_rows.append({**group_id,"month":dt,"actual_qty":float(val)})
        _, valres = compete_models(y)
        for _, vr in valres.iterrows():
            validation_rows.append({**group_id,"Demand Pattern":pattern,**vr.to_dict(),"Selected Model":model})
        future_dates = pd.date_range(idx[-1] + pd.offsets.MonthBegin(1), periods=HORIZON, freq="MS")
        for dt, val in zip(future_dates, pred):
            forecast_rows.append({**group_id,"Demand Pattern":pattern,"Selected Model":model,"Forecast Month":dt,"Forecast Qty":round(float(max(0,val)),2),"Increment %":0.0,"Adjusted Forecast Qty":round(float(max(0,val)),2)})
        quality_rows.append({**group_id,"First Actual Month":idx[0],"Last Actual Month":idx[-1],"History Months":len(y),"Non-zero Months":int(np.count_nonzero(y)),"Zero Demand %":round(float(np.mean(y==0)*100),1),"Demand Pattern":pattern,"Selected Model":model})
        if i % 10 == 0 or i == len(groups)-1:
            progress.progress((i+1)/len(groups), text=f"Analysing combination {i+1:,} of {len(groups):,}...")
    progress.empty()

    forecast_df = pd.DataFrame(forecast_rows)
    validation_df = pd.DataFrame(validation_rows)
    historical_df = pd.DataFrame(hist_rows)
    quality_df = pd.DataFrame(quality_rows)
    if forecast_df.empty:
        st.error("No forecast could be generated from the detected records.")
        st.stop()

    st.markdown("### 2. Automatic demand analysis")
    a,b,c,d = st.columns(4)
    a.metric("Combinations forecast", f"{len(groups):,}")
    b.metric("Historical months", f"{hist['month'].nunique():,}")
    c.metric("Forecast horizon", "24 months")
    d.metric("Demand patterns", f"{quality_df['Demand Pattern'].nunique()}")
    pattern_counts = quality_df["Demand Pattern"].value_counts().rename_axis("Demand Pattern").reset_index(name="Combinations")
    st.dataframe(pattern_counts, use_container_width=True, hide_index=True)
    st.caption("New NPI / limited-history items use a conservative benchmark unless launch, nomination or comparable-product information is present in the uploaded data. The engine does not invent launch demand.")

    st.markdown("### 3. Editable 24-month forecast")
    st.caption("Edit Forecast Qty or Increment %. Increment is applied to the model forecast; your historical actuals remain unchanged. You can edit individual months and SKUs.")
    display_cols = keys + ["Demand Pattern","Selected Model","Forecast Month","Forecast Qty","Increment %","Adjusted Forecast Qty"]
    edited = st.data_editor(
        forecast_df[display_cols],
        use_container_width=True,
        hide_index=True,
        num_rows="fixed",
        column_config={
            "Forecast Month": st.column_config.DateColumn("Forecast Month", disabled=True, format="MMM YYYY"),
            "Demand Pattern": st.column_config.TextColumn(disabled=True),
            "Selected Model": st.column_config.TextColumn(disabled=True),
            "Forecast Qty": st.column_config.NumberColumn("Forecast Qty", min_value=0.0, step=1.0, format="%.2f"),
            "Increment %": st.column_config.NumberColumn("Increment %", min_value=-100.0, max_value=500.0, step=1.0, format="%.1f"),
            "Adjusted Forecast Qty": st.column_config.NumberColumn("Adjusted Forecast Qty", disabled=True, format="%.2f"),
        },
        key="forecast_editor"
    )
    edited = edited.copy()
    # Recalculate adjusted quantity from current editable baseline and increment.
    edited["Adjusted Forecast Qty"] = (pd.to_numeric(edited["Forecast Qty"], errors="coerce").fillna(0) * (1 + pd.to_numeric(edited["Increment %"], errors="coerce").fillna(0)/100)).clip(lower=0).round(2)
    st.dataframe(edited[display_cols], use_container_width=True, hide_index=True)

    st.markdown("### 4. Validation and data quality")
    if not validation_df.empty:
        st.dataframe(validation_df, use_container_width=True, hide_index=True)
    st.caption("WAPE and bias are calculated on a chronological holdout period. They are not guarantees of future accuracy. For short or sparse histories, validation may be unavailable or unstable.")
    with st.expander("Data quality and history by SKU"):
        st.dataframe(quality_df, use_container_width=True, hide_index=True)

    # Make output file with adjusted forecasts
    out = make_excel(raw_bytes, uploaded.name, historical_df, edited, validation_df, quality_df)
    st.markdown("### 5. Download your completed workbook")
    st.download_button(
        "Download Harmony 24-Month Forecast (Excel)",
        data=out,
        file_name=f"Harmony_Manufacturing_Forecast_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True
    )
    st.caption("The export includes Historical Actuals, Editable 24M Forecast, Model Validation, Data Quality and copies of the original workbook worksheets where supported.")

if __name__ == "__main__":
    main()
