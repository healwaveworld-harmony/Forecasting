import streamlit as st
import pandas as pd
import numpy as np
import openpyxl
import requests
import zipfile
import xml.etree.ElementTree as ET
from io import BytesIO
from datetime import datetime
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Demand Forecasting | Product Intelligence", page_icon="📈", layout="wide")
st.markdown("""
<style>
.block-container{padding-top:1.3rem;max-width:1600px}
h1,h2,h3{letter-spacing:-.03em}
[data-testid="stMetric"]{background:#f6f8fb;padding:14px;border-radius:12px}
div[data-testid="stDataFrame"]{border:1px solid #e7eaf0;border-radius:10px}
</style>
""", unsafe_allow_html=True)

st.title("Demand Forecasting")
st.caption("Per-sheet category forecasts • consolidated plan • product assumptions • optional OpenRouter commentary")

DEFAULT_FILE = "1.Sales Volume Projections 25-26 to 29-30_12.11.2024_Rev11-Master.xlsx"
PLANT_SHEETS = ["Manesar", "Gujrat"]
MONTH_START_COL = 36  # AJ; source workbook has monthly columns from AJ onwards
MONTH_END_COL = 95    # CQ inclusive (5 fiscal years × 12 months)
META_COLS = {"Customer":3,"Model":4,"Vehicle / Nameplate":5,"Part No.":6,
             "SOP":7,"SAP Code":8,"Product":9,"Segment":10,"Material":11,
             "Material Description":12,"Specifications":13,"Paint":14,
             "Machining":16,"Machined Description":17,"Casting Weight":18,
             "AfterMC Weight":19,"Yield":20,"FY25 Melting kg/pcs":21,"FY26 Melting kg/pcs":22,"FY27 Melting kg/pcs":23,"FY28 Melting kg/pcs":24,"FY29 Melting kg/pcs":25,"FY25 Yield":26,"FY26 Yield":27,"FY27 Yield":28,"FY28 Yield":29,"FY29 Yield":30,"Melting":31,"Casting":32,"AfterMC":33,"Mold":34,"Credit Period Days":35}

def numeric(v):
    try:
        if v is None or v == "": return np.nan
        return float(v)
    except Exception:
        return np.nan

def load_source(file_bytes):
    wb = openpyxl.load_workbook(BytesIO(file_bytes), data_only=True, read_only=False)
    records, issues = [], []
    price_cols = {
        "Sales Price BP INR": 109,       # DE: BP@ / Sales Price
        "Latest PO Price INR": 121,      # DQ: Latest PO Price
        "Budget Price INR": 123,         # DS: Budget / price
    }
    for sheet in PLANT_SHEETS:
        if sheet not in wb.sheetnames:
            issues.append(f"Missing plant sheet: {sheet}")
            continue
        ws = wb[sheet]
        date_cols = []
        for c in range(MONTH_START_COL, min(ws.max_column, MONTH_END_COL)+1):
            d = ws.cell(6,c).value
            plan_label = str(ws.cell(8,c).value or "").lower()
            if isinstance(d, (datetime, pd.Timestamp)) and "plan" in plan_label:
                date_cols.append((c, pd.Timestamp(d).to_period("M").to_timestamp()))
        if not date_cols:
            issues.append(f"{sheet}: no month headers labelled Plan found in row 8 (AJ:CQ).")
            continue
        for r in range(11, ws.max_row+1):
            customer = ws.cell(r,3).value
            part = ws.cell(r,6).value
            product = ws.cell(r,9).value
            model = ws.cell(r,4).value
            if all(v is None or str(v).strip()=="" for v in [customer,part,product,model]):
                continue
            vals = []
            for c, month in date_cols:
                value = numeric(ws.cell(r,c).value)
                # Empty source cells remain missing; do not convert blanks into zero demand.
                if pd.notna(value):
                    vals.append({"month":month,"quantity":float(value)})
            month_df = pd.DataFrame(vals, columns=["month","quantity"])
            meta = {k:ws.cell(r,c).value for k,c in META_COLS.items()}
            prices = {k:numeric(ws.cell(r,c).value) for k,c in price_cols.items()}
            # Keep source financial context and annual business-value columns for review.
            annual_values = {}
            for col, label in zip(range(128,133), ["FY25 Business Value INR Mn","FY26 Business Value INR Mn",
                                                   "FY27 Business Value INR Mn","FY28 Business Value INR Mn",
                                                   "FY29 Business Value INR Mn"]):
                annual_values[label] = numeric(ws.cell(r,col).value)
            bp = prices["Sales Price BP INR"]
            po = prices["Latest PO Price INR"]
            budget = prices["Budget Price INR"]
            suggested_price = bp if pd.notna(bp) and bp > 0 else (po if pd.notna(po) and po > 0 else (budget if pd.notna(budget) and budget > 0 else np.nan))
            records.append({"plant":sheet,"source_row":r,"status":ws.cell(r,1).value,
                            "serial":ws.cell(r,2).value, **meta, **prices, **annual_values,
                            "suggested_price":suggested_price,
                            "monthly":month_df,
                            "fractional_plan_points":int(((month_df["quantity"] % 1).abs() > 1e-9).sum()) if not month_df.empty else 0})
    wb.close()
    return records, issues, wb.sheetnames

def build_forecast(rec, growth, horizon_years, start_month, manual_base=None, manual_monthly=None):
    hist = rec["monthly"].copy()
    if not hist.empty:
        hist = hist[hist["month"] < start_month].sort_values("month")
    # Use fixed calendar windows, not the last N nonblank cells. This prevents
    # a blank recent month from causing an older month to be pulled into the average.
    last12_start = start_month - pd.DateOffset(months=12)
    last3_start = start_month - pd.DateOffset(months=3)
    last12 = hist[(hist["month"] >= last12_start) & (hist["month"] < start_month)] if not hist.empty else pd.DataFrame(columns=["month","quantity"])
    last3 = hist[(hist["month"] >= last3_start) & (hist["month"] < start_month)] if not hist.empty else pd.DataFrame(columns=["month","quantity"])
    avg12 = float(last12.quantity.mean()) if len(last12) else np.nan
    avg3 = float(last3.quantity.mean()) if len(last3) else (avg12 if pd.notna(avg12) else np.nan)
    blended = (avg12 + avg3) / 2 if pd.notna(avg12) and pd.notna(avg3) else (avg12 if pd.notna(avg12) else 0.0)
    if manual_base is not None and pd.notna(manual_base):
        blended = max(0.0, float(manual_base))
    # Do not infer a positive plan from blanks. A zero-only history is explicitly flagged for review.
    nonzero_count = int((hist["quantity"].fillna(0) != 0).sum()) if not hist.empty else 0
    plan_status = "Source plan available" if nonzero_count else "No non-zero source plan — enter a plan or confirm zero"
    if len(last12):
        seasonal = last12.copy()
        seasonal["month_no"] = seasonal.month.dt.month
        pattern = seasonal.groupby("month_no").quantity.mean().to_dict()
        pattern_mean = float(np.mean(list(pattern.values()))) if pattern else 0
    else:
        pattern, pattern_mean = {}, 0
    dates = pd.date_range(start=start_month, periods=12*horizon_years, freq="MS")
    manual_values = []
    if manual_monthly:
        try:
            manual_values = [float(x) for x in manual_monthly if str(x).strip() != ""]
        except Exception:
            manual_values = []
    rows=[]
    for i,d in enumerate(dates):
        if i < len(manual_values):
            qty = max(0.0, manual_values[i])
        else:
            raw = pattern.get(d.month, blended)
            season_factor = (raw / pattern_mean) if pattern_mean and np.isfinite(raw) else 1.0
            annual_growth_factor = (1 + growth/100) ** ((i//12)+1)
            qty = blended * season_factor * annual_growth_factor
        rows.append({"Month":d,"Forecast Qty":int(max(0,round(qty)))})
    return {"avg12":0.0 if pd.isna(avg12) else avg12,
            "avg3":0.0 if pd.isna(avg3) else avg3,
            "blended":blended,"n12":len(last12),"n3":len(last3),
            "nonzero_count":nonzero_count,"plan_status":plan_status,
            "forecast":pd.DataFrame(rows),"history_end":hist.month.max() if len(hist) else None}

def product_label(r):
    return f'{r["plant"]} | {r.get("Customer") or "No client"} | {r.get("Product") or r.get("Model") or "Unnamed"} | {r.get("Part No.") or "No part"}'

def excel_export(source_bytes, products, params, start_month, years, client_orders, source_name, original_sheetnames):
    # Build forecast tabs separately, then append their XML parts into a copy of the source
    # XLSX package. This preserves every original worksheet XML part byte-for-byte.
    forecast_wb = openpyxl.Workbook()
    forecast_wb.remove(forecast_wb.active)
    added_names = []

    def add_sheet(name, headers, rows):
        base = str(name)[:31]
        final = base
        suffix = 1
        while final in forecast_wb.sheetnames:
            tail = f"_{suffix}"
            final = base[:31-len(tail)] + tail
            suffix += 1
        ws = forecast_wb.create_sheet(final)
        added_names.append(final)
        if headers:
            ws.append([str(h) if isinstance(h, (pd.Timestamp, datetime)) else h for h in headers])
        for row in rows:
            clean = []
            for v in row:
                if isinstance(v, (pd.Timestamp, datetime)):
                    clean.append(v.strftime("%Y-%m"))
                elif pd.isna(v) if not isinstance(v, (str, bytes, list, tuple, dict)) else False:
                    clean.append(None)
                elif isinstance(v, np.generic):
                    clean.append(v.item())
                else:
                    clean.append(v)
            ws.append(clean)
        ws.freeze_panes = "A2"
        if ws.max_row >= 1 and ws.max_column >= 1:
            ws.auto_filter.ref = ws.dimensions
        return ws

    assumption_headers = ["Plant","Source Row","Status","Customer","Model","Vehicle / Nameplate","Part No.",
        "SOP","SAP Code","Product","Segment","Material","Material Description","Specifications","Paint",
        "Machining","Machined Description","Casting Weight","AfterMC Weight","Yield",
        "FY25 Melting kg/pcs","FY26 Melting kg/pcs","FY27 Melting kg/pcs","FY28 Melting kg/pcs","FY29 Melting kg/pcs",
        "FY25 Yield","FY26 Yield","FY27 Yield","FY28 Yield","FY29 Yield","Melting","Casting","AfterMC","Mold","Credit Period Days",
         "Sales Price BP INR","Latest PO Price INR","Budget Price INR",
        "FY25 Business Value INR Mn","FY26 Business Value INR Mn","FY27 Business Value INR Mn","FY28 Business Value INR Mn","FY29 Business Value INR Mn",
        "Confirmed Unit Price INR","Planned Annual Growth %","Trailing 12M Source Mean","Trailing 3M Source Mean",
        "Blended Monthly Base","12M Observations","3M Observations","Reviewed / Confirmed","Price Confirmed","Include"]
    assumption_rows, detail_rows = [], []
    forecast_months = [d.strftime("%Y-%m") for d in pd.date_range(start=start_month, periods=12*years, freq="MS")]

    for rec in products:
        p = params[rec["key"]]
        result = build_forecast(rec, p["growth"], years, start_month, p.get("manual_base"), p.get("manual_monthly"))
        assumption_rows.append([rec["plant"],rec["source_row"],rec.get("status"),rec.get("Customer"),
            rec.get("Model"),rec.get("Vehicle / Nameplate"),rec.get("Part No."),rec.get("SOP"),
            rec.get("SAP Code"),rec.get("Product"),rec.get("Segment"),rec.get("Material"),
            rec.get("Material Description"),rec.get("Specifications"),rec.get("Paint"),rec.get("Machining"),
            rec.get("Machined Description"),rec.get("Casting Weight"),rec.get("AfterMC Weight"),rec.get("Yield"),
            rec.get("FY25 Melting kg/pcs"),rec.get("FY26 Melting kg/pcs"),rec.get("FY27 Melting kg/pcs"),rec.get("FY28 Melting kg/pcs"),rec.get("FY29 Melting kg/pcs"),
            rec.get("FY25 Yield"),rec.get("FY26 Yield"),rec.get("FY27 Yield"),rec.get("FY28 Yield"),rec.get("FY29 Yield"),
            rec.get("Melting"),rec.get("Casting"),rec.get("AfterMC"),rec.get("Mold"),rec.get("Credit Period Days"),
            rec.get("Sales Price BP INR"),rec.get("Latest PO Price INR"),rec.get("Budget Price INR"),
            rec.get("FY25 Business Value INR Mn"),rec.get("FY26 Business Value INR Mn"),rec.get("FY27 Business Value INR Mn"),
            rec.get("FY28 Business Value INR Mn"),rec.get("FY29 Business Value INR Mn"),
            p.get("price"),p["growth"],result["avg12"],result["avg3"],result["blended"],result["n12"],result["n3"],
            "Yes" if p["confirmed"] else "Pending","Yes" if p.get("price_confirmed") else "Pending",
            "Yes" if p["include"] else "No"])
        if p["include"]:
            for _, row in result["forecast"].iterrows():
                detail_rows.append({"Plant":rec["plant"],"Customer":rec.get("Customer"),
                    "Model":rec.get("Model"),"Vehicle / Nameplate":rec.get("Vehicle / Nameplate"),
                    "Product":rec.get("Product"),"Part No.":rec.get("Part No."),
                    "Segment":rec.get("Segment"),"Material":rec.get("Material"),
                     "Month":row["Month"].strftime("%Y-%m"),"Forecast Qty":float(row["Forecast Qty"]),
                    "Unit Price INR":p.get("price"),"Forecast Value INR":float(row["Forecast Qty"])*(float(p.get("price") or 0)),
                    "Forecast Value INR Mn":float(row["Forecast Qty"])*(float(p.get("price") or 0))/1_000_000,
                    "Annual Growth %":p["growth"],
                    "Confirmation":"Yes" if p["confirmed"] else "Pending"})

    add_sheet("Forecast Assumptions", assumption_headers, assumption_rows)
    detail_headers = ["Plant","Customer","Model","Vehicle / Nameplate","Product","Part No.",
                      "Segment","Material","Month","Forecast Qty","Unit Price INR","Forecast Value INR","Forecast Value INR Mn","Annual Growth %","Confirmation"]
    if detail_rows:
        ddf = pd.DataFrame(detail_rows)
        add_sheet("Forecast Monthly", detail_headers, ddf[detail_headers].values.tolist())
    else:
        ddf = pd.DataFrame(columns=detail_headers)
        add_sheet("Forecast Monthly", ["Message"], [["No included products have forecast values."]])

    # Companion tabs: each category view is calculated from the selected source records.
    sheet_category_rules = {
        "manesar": (["Plant","Customer","Model","Product","Part No."], lambda d: d["Plant"].eq("Manesar")),
        "gujrat": (["Plant","Customer","Model","Product","Part No."], lambda d: d["Plant"].eq("Gujrat")),
        "combined sales": (["Plant","Customer","Model","Product","Part No."], lambda d: pd.Series(True,index=d.index)),
        "summary": (["Customer","Segment","Plant"], lambda d: pd.Series(True,index=d.index)),
        "summary (2)": (["Customer","Plant","Segment"], lambda d: pd.Series(True,index=d.index)),
        "sheet1": (["Customer","Model","Product"], lambda d: pd.Series(True,index=d.index)),
        "sheet1 (2)": (["Customer","Segment","Plant"], lambda d: pd.Series(True,index=d.index)),
        "assumptions": (["Customer","Segment","Plant"], lambda d: pd.Series(True,index=d.index)),
        "m-o-m vol vs cap. summary": (["Plant"], lambda d: pd.Series(True,index=d.index)),
        "m&m business": (["Customer","Segment","Plant"], lambda d: pd.Series(False,index=d.index))
    }
    def sheet_safe(s):
        return re.sub(r"[^A-Za-z0-9]+","_",str(s)).strip("_")[:24] or "Sheet"

    for source_sheet in original_sheetnames:
        rule = sheet_category_rules.get(str(source_sheet).strip().lower())
        if rule is None:
            continue
        group_cols, selector = rule
        if ddf.empty:
            add_sheet(("FC_"+sheet_safe(source_sheet))[:31],
                      ["Source Sheet","Status"], [[source_sheet,"No forecastable product records were included."]])
            continue
        subset = ddf.loc[selector(ddf)].copy()
        if subset.empty:
            add_sheet(("FC_"+sheet_safe(source_sheet))[:31],
                      ["Source Sheet","Status"], [[source_sheet,"No matching forecast records."]])
            continue
        grouped = subset.groupby(group_cols+["Month"], dropna=False, as_index=False)["Forecast Qty"].sum()
        pivot = grouped.pivot_table(index=group_cols, columns="Month", values="Forecast Qty",
                                    aggfunc="sum", fill_value=0).reset_index()
        month_columns = sorted([c for c in pivot.columns if isinstance(c, str) and re.match(r"^\d{4}-\d{2}$",c)])
        pivot["Forecast Total"] = pivot[month_columns].sum(axis=1) if month_columns else 0
        for yr in range(years):
            year_months = forecast_months[yr*12:(yr+1)*12]
            pivot[f"Forecast Year {yr+1} Total"] = pivot[year_months].sum(axis=1) if year_months else 0
        ordered = group_cols + month_columns + [f"Forecast Year {yr+1} Total" for yr in range(years)] + ["Forecast Total"]
        rows = [[rr[c] for c in ordered] for _,rr in pivot.iterrows()]
        add_sheet(("FC_"+sheet_safe(source_sheet))[:31], ordered, rows)

    # One consolidated final tab containing every included product/month forecast.
    if detail_rows:
        add_sheet("All Plans Calculated", detail_headers, ddf[detail_headers].values.tolist())
    else:
        add_sheet("All Plans Calculated", ["Status"], [["No included product forecasts to calculate."]])

    order_headers = ["Client","Plant","Order / Project","Model","Product","Part No.","Monthly Qty",
                     "Action","Recorded At"]
    order_rows = [[x.get(h) for h in order_headers] for x in client_orders]
    add_sheet("Client & Order Register", order_headers, order_rows)
    add_sheet("Forecast Read Me", ["Item","Details"], [
        ["Source Workbook",source_name],
        ["Forecast Start",start_month.strftime("%Y-%m")],
        ["Forecast Horizon (years)",years],
        ["Baseline method","50% mean of nonblank Plan cells in the 12 calendar months before forecast start + 50% mean of nonblank Plan cells in the 3 calendar months before forecast start. Blank months are not backfilled from older months."],
        ["Seasonality","Latest available 12-month pattern normalized to the blended monthly base."],
        ["Growth","Planned annual growth is compounded by forecast year and applied month-wise."],
        ["Sheet design","FC_ companion tabs are appended for supported original sheets, grouped by the relevant category."],
        ["Final consolidated output","All Plans Calculated contains included product-level forecasts across plants."],
        ["Source caveat","Manesar and Gujrat source monthly columns are labelled Plan; averages reflect source planning values, not independently verified actual sales."],
        ["Pricing","Unit price defaults to DE (BP sales price), then DQ (latest PO price), then DS (budget price); user can edit and confirm it."],
        ["Value calculation","Forecast value = rounded whole-piece forecast quantity × confirmed unit price; value is shown in INR and INR million."],
        ["Quantity validation","Source monthly plan values are checked for fractional quantities; forecast quantities are rounded to whole pieces."],
        ["AI","OpenRouter is optional and used only for aggregate narrative commentary. It does not change numerical forecast calculations."],
        ["Original sheets","Original workbook package parts and original worksheet XML are retained byte-for-byte."]
    ])

    # Save the new tabs into a temporary workbook; all forecast cells use default style.
    temp = BytesIO()
    forecast_wb.save(temp)
    temp.seek(0)

    MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
    CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
    ET.register_namespace("", MAIN_NS)
    ET.register_namespace("r", REL_NS)

    source_zip = zipfile.ZipFile(BytesIO(source_bytes), "r")
    temp_zip = zipfile.ZipFile(temp, "r")
    workbook_xml = ET.fromstring(source_zip.read("xl/workbook.xml"))
    rels_xml = ET.fromstring(source_zip.read("xl/_rels/workbook.xml.rels"))
    content_xml = ET.fromstring(source_zip.read("[Content_Types].xml"))

    sheets_node = workbook_xml.find(f"{{{MAIN_NS}}}sheets")
    existing_sheet_ids = [int(x.attrib.get("sheetId","0")) for x in sheets_node]
    next_sheet_id = max(existing_sheet_ids or [0]) + 1
    rel_ids = []
    for rel in rels_xml:
        m = re.match(r"rId(\d+)$", rel.attrib.get("Id",""))
        if m: rel_ids.append(int(m.group(1)))
    next_rel_id = max(rel_ids or [0]) + 1
    existing_parts = source_zip.namelist()
    part_nums = []
    for part in existing_parts:
        m = re.match(r"xl/worksheets/sheet(\d+)\.xml$",part)
        if m: part_nums.append(int(m.group(1)))
    next_part_num = max(part_nums or [0]) + 1

    new_parts = {}
    for i, sheet_name in enumerate(added_names, start=1):
        new_part = f"xl/worksheets/sheet{next_part_num+i-1}.xml"
        temp_part = f"xl/worksheets/sheet{i}.xml"
        new_parts[new_part] = temp_zip.read(temp_part)
        sheet_id = next_sheet_id + i - 1
        rid = f"rId{next_rel_id+i-1}"
        ET.SubElement(sheets_node, f"{{{MAIN_NS}}}sheet",
                      {"name":sheet_name,"sheetId":str(sheet_id),f"{{{REL_NS}}}id":rid})
        ET.SubElement(rels_xml, f"{{{PKG_REL_NS}}}Relationship",
                      {"Id":rid,"Type":f"{REL_NS}/worksheet",
                       "Target":f"worksheets/sheet{next_part_num+i-1}.xml"})
        ET.SubElement(content_xml, f"{{{CT_NS}}}Override",
                      {"PartName":f"/{new_part}",
                       "ContentType":"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"})

    workbook_bytes = ET.tostring(workbook_xml,encoding="utf-8",xml_declaration=True)
    rels_bytes = ET.tostring(rels_xml,encoding="utf-8",xml_declaration=True)
    content_bytes = ET.tostring(content_xml,encoding="utf-8",xml_declaration=True)

    final = BytesIO()
    with zipfile.ZipFile(final,"w",zipfile.ZIP_DEFLATED) as outzip:
        for item in source_zip.infolist():
            data = source_zip.read(item.filename)
            if item.filename == "xl/workbook.xml":
                data = workbook_bytes
            elif item.filename == "xl/_rels/workbook.xml.rels":
                data = rels_bytes
            elif item.filename == "[Content_Types].xml":
                data = content_bytes
            outzip.writestr(item,data)
        for name,data in new_parts.items():
            outzip.writestr(name,data,compress_type=zipfile.ZIP_DEFLATED)
    source_zip.close()
    temp_zip.close()
    final.seek(0)
    return final.getvalue()

uploaded = st.sidebar.file_uploader("Upload the master sales workbook (.xlsx)", type=["xlsx"])
if uploaded:
    file_bytes=uploaded.getvalue(); source_name=uploaded.name
else:
    candidate=DEFAULT_FILE
    try:
        with open(candidate,"rb") as f: file_bytes=f.read()
        source_name=candidate
        st.sidebar.success("Loaded the attached master workbook.")
    except Exception:
        file_bytes=None; source_name=""

if not file_bytes:
    st.info("Upload the sales projection workbook to begin. The program reads the Manesar and Gujrat plant sheets and keeps source fields available for review.")
    st.stop()

try:
    records, issues, sheet_names=load_source(file_bytes)
except Exception as e:
    st.error(f"Could not read workbook: {e}")
    st.stop()

if not records:
    st.error("No product rows were detected. Check that Manesar/Gujrat product data starts at row 11 and monthly dates are in row 6.")
    st.stop()

if "custom_products" not in st.session_state:
    st.session_state.custom_products = []
records.extend(st.session_state.custom_products)
for i,r in enumerate(records):
    r["key"]=f'{r["plant"]}_{r["source_row"]}_{i}'

if "params" not in st.session_state:
    st.session_state.params={}
for r in records:
    st.session_state.params.setdefault(r["key"],{"growth":0.0,"confirmed":False,"include":True,
        "manual_base":None,"manual_monthly":None,"price":(float(r["suggested_price"]) if pd.notna(r.get("suggested_price",np.nan)) else None),
        "price_confirmed":False})
    st.session_state.params[r["key"]].setdefault("manual_base",None)
    st.session_state.params[r["key"]].setdefault("manual_monthly",None)
    st.session_state.params[r["key"]].setdefault("price",(float(r["suggested_price"]) if pd.notna(r.get("suggested_price",np.nan)) else None))
    st.session_state.params[r["key"]].setdefault("price_confirmed",False)
if "client_orders" not in st.session_state:
    st.session_state.client_orders=[]

with st.sidebar:
    st.header("Forecast controls")
    horizon=st.selectbox("Forecast horizon", [1,2], index=1, format_func=lambda x:f"{x} year" if x==1 else f"{x} years")
    start_date=st.date_input("Forecast start month", value=pd.Timestamp.today().replace(day=1).date())
    start_month=pd.Timestamp(start_date).to_period("M").to_timestamp()
    plant_filter=st.multiselect("Plant sheets", PLANT_SHEETS, default=PLANT_SHEETS)
    customer_options=sorted({str(r.get("Customer")) for r in records if r.get("Customer") not in [None,""]})
    customer_filter=st.multiselect("Clients", customer_options)
    search=st.text_input("Find product / part / model")
    st.divider()
    st.caption("Optional AI commentary (forecast maths never depends on AI)")
    openrouter_key=st.text_input("OpenRouter API key (optional)",type="password")
    openrouter_model=st.text_input("Small model ID",value="openai/gpt-4o-mini")

visible=[r for r in records if r["plant"] in plant_filter
         and (not customer_filter or str(r.get("Customer")) in customer_filter)
         and (not search or search.lower() in " ".join(str(r.get(k) or "") for k in ["Customer","Model","Product","Part No.","SAP Code"]).lower())]

st.caption(f"Source: {source_name} · {len(records):,} product rows detected across plant sheets")
st.warning("Source-data note: the Manesar and Gujrat monthly columns are labelled “Plan” in the uploaded workbook. The displayed trailing averages therefore reflect the available source-plan values; they should not be described as actual sales unless you replace them with actuals.")
if issues:
    for issue in issues: st.warning(issue)
tabs=st.tabs(["1 · Product assumptions","2 · Client & order register","3 · Forecast review","4 · Export"])

with tabs[0]:
    st.subheader("Bulk product planning & price confirmation")
    st.write("Edit growth, inclusion, baseline and price for many products in one grid. Expand a product only when you need to enter a month-wise manual plan.")
    if not visible:
        st.warning("No products match the current filters.")
    else:
        # Bulk growth update for all currently filtered products.
        st.markdown("#### Apply one growth change to all filtered products")
        b1,b2,b3=st.columns([2,1,1])
        growth_options=["No change","-20%","-10%","-5%","0%","+5%","+10%","+15%","+20%","+25%","Custom"]
        bulk_choice=b1.selectbox("Growth setting",growth_options,index=0,key="bulk_growth_choice")
        custom_growth=b2.number_input("Custom annual growth (%)",value=0.0,step=0.5,key="bulk_custom_growth")
        apply_growth=b3.button("Apply to filtered products",type="secondary")
        if apply_growth:
            if bulk_choice != "No change":
                new_growth=custom_growth if bulk_choice=="Custom" else float(bulk_choice.replace("%","").replace("+",""))
                for rec in visible:
                    st.session_state.params[rec["key"]]["growth"]=new_growth
                st.session_state.pop("bulk_product_editor",None)
            st.success("Bulk growth settings applied. Review the editable table below and save.")

        grid_rows=[]
        for rec in visible:
            p=st.session_state.params[rec["key"]]
            pr=build_forecast(rec,p["growth"],horizon,start_month,p.get("manual_base"),p.get("manual_monthly"))
            grid_rows.append({"Key":rec["key"],"Plant":rec["plant"],"Customer":rec.get("Customer"),
                "Product":rec.get("Product") or rec.get("Model"),"Part No.":rec.get("Part No."),
                 "Plan status":pr["plan_status"],"Plan Months":len(rec["monthly"]),
                "Fractional Plan Values":rec.get("fractional_plan_points",0),
                "Last 12M Avg":pr["avg12"],"Last 3M Avg":pr["avg3"],
                "Blended Base":pr["blended"],"Annual Growth %":p["growth"],"Include":p["include"],
                "Manual Monthly Base":p.get("manual_base"),"Unit Price INR":p.get("price"),
                "Price Confirmed":p.get("price_confirmed",False),"Demand Confirmed":p.get("confirmed",False)})
        grid=pd.DataFrame(grid_rows)
        st.caption("Editable columns: annual growth (positive or negative), include, manual base, unit price and confirmations. Blank monthly plan rows are not treated as demand.")
        edited=st.data_editor(grid,use_container_width=True,hide_index=True,height=520,
            disabled=["Key","Plant","Customer","Product","Part No.","Plan status","Plan Months","Fractional Plan Values","Last 12M Avg","Last 3M Avg","Blended Base"],
            column_config={
                "Annual Growth %":st.column_config.NumberColumn("Annual Growth %",min_value=-100,max_value=500,step=0.5,format="%.1f%%"),
                "Include":st.column_config.CheckboxColumn("Include"),
                "Manual Monthly Base":st.column_config.NumberColumn("Manual Monthly Base",min_value=0,step=1,help="Use when source plan is blank or you want to override the blended average."),
                "Unit Price INR":st.column_config.NumberColumn("Unit Price (₹)",min_value=0,step=0.01,format="₹%.2f"),
                "Price Confirmed":st.column_config.CheckboxColumn("Price confirmed"),
                "Demand Confirmed":st.column_config.CheckboxColumn("Demand confirmed")
            },key="bulk_product_editor")
        if st.button("Save all product and price changes",type="primary",key="save_bulk_product_edits"):
            for _,row in edited.iterrows():
                k=row["Key"]
                if k in st.session_state.params:
                    old=st.session_state.params[k]
                    old.update({"growth":float(row["Annual Growth %"]),"include":bool(row["Include"]),
                        "manual_base":None if pd.isna(row["Manual Monthly Base"]) else float(row["Manual Monthly Base"]),
                        "price":None if pd.isna(row["Unit Price INR"]) else float(row["Unit Price INR"]),
                        "price_confirmed":bool(row["Price Confirmed"]),"confirmed":bool(row["Demand Confirmed"])})
            st.session_state.pop("bulk_product_editor",None)
            st.success("All visible product assumptions and prices saved.")
            st.rerun()

        st.divider()
        st.markdown("#### Manual month-wise plan and detailed product review")
        for rec in visible:
            p=st.session_state.params[rec["key"]]
            preview=build_forecast(rec,p["growth"],horizon,start_month,p.get("manual_base"),p.get("manual_monthly"))
            label=product_label(rec)
            with st.expander(label,expanded=False):
                st.caption(f'{preview["plan_status"]} · {preview["n12"]} trailing 12-month observations · {preview["n3"]} trailing 3-month observations')
                if rec["monthly"].empty or preview["nonzero_count"]==0:
                    st.warning("No non-zero monthly Plan quantities were found before the forecast start date. Enter a manual monthly plan or a manual monthly base below. A blank is not converted into a demand quantity.")
                price_context=pd.DataFrame([{
                    "BP Sales Price (₹)":rec.get("Sales Price BP INR"),
                    "Latest PO Price (₹)":rec.get("Latest PO Price INR"),
                    "Budget Price (₹)":rec.get("Budget Price INR"),
                    "FY25 Business Value (₹ Mn)":rec.get("FY25 Business Value INR Mn"),
                    "FY26 Business Value (₹ Mn)":rec.get("FY26 Business Value INR Mn"),
                    "FY27 Business Value (₹ Mn)":rec.get("FY27 Business Value INR Mn"),
                    "FY28 Business Value (₹ Mn)":rec.get("FY28 Business Value INR Mn"),
                    "FY29 Business Value (₹ Mn)":rec.get("FY29 Business Value INR Mn")}])
                st.markdown("**Source pricing and business-value context**")
                st.dataframe(price_context,use_container_width=True,hide_index=True)
                st.markdown("**Source monthly Plan values (AJ:CQ)**")
                src=rec["monthly"].copy()
                if not src.empty:
                    st.dataframe(src.assign(Month=src["month"].dt.strftime("%b %Y"))[["Month","quantity"]].rename(columns={"quantity":"Plan Qty"}),
                                 use_container_width=True,hide_index=True,height=220)
                else:
                    st.info("No numeric monthly Plan values were present for this product.")
                with st.form(f'manual_plan_{rec["key"]}'):
                    m1,m2=st.columns(2)
                    manual_base=m1.number_input("Override blended monthly base (optional)",min_value=0.0,
                        value=float(p["manual_base"] or 0.0),step=1.0,key=f'mbase_{rec["key"]}')
                    price=m2.number_input("Confirmed unit selling price (₹)",min_value=0.0,
                        value=float(p["price"] or 0.0),step=0.01,key=f'price_{rec["key"]}')
                    monthly_text=st.text_area("Manual monthly plan quantities (comma-separated, in forecast month order)",
                        value=",".join(str(x) for x in p.get("manual_monthly", [])) if isinstance(p.get("manual_monthly"), list) else (p.get("manual_monthly") or ""),height=90,
                        placeholder="Example: 1200, 1250, 1300, 1300 ... enter 12 or 24 values as needed",
                        key=f'mplan_{rec["key"]}')
                    demand_confirm=st.checkbox("I confirm the demand assumptions",value=bool(p["confirmed"]),key=f'dconfirm_{rec["key"]}')
                    price_confirm=st.checkbox("I confirm this unit price",value=bool(p.get("price_confirmed",False)),key=f'pconfirm_{rec["key"]}')
                    save_detail=st.form_submit_button("Save detailed plan and price")
                    if save_detail:
                        vals=None
                        if monthly_text.strip():
                            try:
                                vals=[float(x.strip()) for x in re.split(r"[,;\n]+",monthly_text) if x.strip()]
                                if len(vals) not in [12,24]:
                                    st.error("Enter exactly 12 or 24 monthly quantities, separated by commas.")
                                    vals=None
                            except Exception:
                                st.error("Monthly plan must contain numeric quantities separated by commas.")
                                vals=None
                        if not monthly_text.strip() or vals is not None:
                            p["manual_base"]=manual_base if manual_base>0 else None
                            p["manual_monthly"]=vals
                            p["price"]=price if price>0 else None
                            p["confirmed"]=demand_confirm
                            p["price_confirmed"]=price_confirm
                            st.success("Detailed product plan and price saved.")
                st.markdown(f"**Forecast preview — {12*horizon} months**")
                st.dataframe(preview["forecast"].assign(Month=lambda x:x.Month.dt.strftime("%b %Y")),
                             use_container_width=True,hide_index=True,height=360)

with tabs[1]:
    st.subheader("Client & order register")
    st.write("Add a new client/order as a forecast product using its entered monthly quantity as the starting run-rate, or mark a matching existing product for exclusion. Original source sheets are never edited.")
    with st.form("add_order"):
        c1,c2,c3=st.columns(3)
        client=c1.text_input("Client / customer")
        plant=c2.selectbox("Plant",PLANT_SHEETS)
        order=c3.text_input("Order / project reference")
        c4,c5,c6=st.columns(3)
        model=c4.text_input("Model")
        product=c5.text_input("Product")
        part=c6.text_input("Part number")
        qty=st.number_input("Monthly quantity (optional)",min_value=0.0,value=0.0,step=1.0)
        new_price=st.number_input("Unit selling price (₹, optional)",min_value=0.0,value=0.0,step=0.01)
        status=st.selectbox("Action",["Add","Delete"])
        add=st.form_submit_button("Add register entry")
        if add:
            entry={"Client":client,"Plant":plant,"Order / Project":order,
                 "Model":model,"Product":product,"Part No.":part,"Monthly Qty":qty,"Unit Price INR":new_price or None,"Action":status,
                "Recorded At":datetime.now().strftime("%Y-%m-%d %H:%M")}
            st.session_state.client_orders.append(entry)
            if status == "Add":
                months = pd.date_range(end=start_month-pd.offsets.MonthBegin(1), periods=12, freq="MS")
                custom = {"plant":plant,"source_row":0,"status":"New","serial":None,
                    "Customer":client,"Model":model,"Vehicle / Nameplate":model,"Part No.":part,
                    "SOP":None,"SAP Code":None,"Product":product,"Segment":None,"Material":None,
                    "Material Description":None,"Specifications":None,"Paint":None,"Machining":None,
                    "Machined Description":None,"Casting Weight":None,"AfterMC Weight":None,"Yield":None,
                    "FY25 Melting kg/pcs":None,"FY26 Melting kg/pcs":None,"FY27 Melting kg/pcs":None,
                    "FY28 Melting kg/pcs":None,"FY29 Melting kg/pcs":None,"FY25 Yield":None,
                    "FY26 Yield":None,"FY27 Yield":None,"FY28 Yield":None,"FY29 Yield":None,
                    "Melting":None,"Casting":None,"AfterMC":None,"Mold":None,"Credit Period Days":None,
                    "Sales Price BP INR":None,"Latest PO Price INR":None,"Budget Price INR":None,
                    "FY25 Business Value INR Mn":None,"FY26 Business Value INR Mn":None,"FY27 Business Value INR Mn":None,
                    "FY28 Business Value INR Mn":None,"FY29 Business Value INR Mn":None,
                    "suggested_price":(new_price if new_price>0 else None),
                    "fractional_plan_points":0,
                    "monthly":pd.DataFrame({"month":months,"quantity":[qty]*12})}
                st.session_state.custom_products.append(custom)
            else:
                matched=[]
                for existing in records:
                    same_client = str(existing.get("Customer") or "").strip().lower() == str(client or "").strip().lower()
                    same_part = bool(part) and str(existing.get("Part No.") or "").strip().lower() == str(part).strip().lower()
                    same_product = bool(product) and str(existing.get("Product") or "").strip().lower() == str(product).strip().lower()
                    if existing["plant"] == plant and same_client and (same_part or same_product):
                        st.session_state.params[existing["key"]]["include"] = False
                        matched.append(product_label(existing))
                if matched:
                    st.success(f"Deletion recorded. {len(matched)} matching product(s) excluded from the forecast.")
                else:
                    st.warning("Deletion request recorded, but no exact plant/client/product match was found. Check the client, product and part number.")
            st.success("Register entry added.")
    if st.session_state.client_orders:
        reg=pd.DataFrame(st.session_state.client_orders)
        st.dataframe(reg,use_container_width=True,hide_index=True)
        idx=st.selectbox("Select register row to remove",options=list(range(len(reg))),
                         format_func=lambda x:f'{x+1}. {reg.iloc[x]["Client"]} — {reg.iloc[x]["Action"]}')
        if st.button("Remove selected register entry"):
            st.session_state.client_orders.pop(idx); st.rerun()
    else:
        st.info("No client/order changes recorded yet.")

with tabs[2]:
    st.subheader("Forecast review")
    st.write("The forecast uses a blended monthly baseline (50% trailing 12-month source mean + 50% trailing 3-month source mean), retains the latest 12-month seasonal profile, and applies the confirmed annual growth rate.")
    rows=[]
    forecasts=[]
    unconfirmed=[]
    for rec in visible:
        p=st.session_state.params[rec["key"]]
        if not p["include"]: continue
        if not p["confirmed"]: unconfirmed.append(product_label(rec))
        result=build_forecast(rec,p["growth"],horizon,start_month,p.get("manual_base"),p.get("manual_monthly"))
        rows.append({"Plant":rec["plant"],"Customer":rec.get("Customer"),"Product":rec.get("Product"),
            "Part No.":rec.get("Part No."),"Growth %":p["growth"],"Last 12M Avg":result["avg12"],
             "Last 3M Avg":result["avg3"],"Blended Base":result["blended"],
            "Unit Price INR":p.get("price"),
            "Demand Value INR":float(result["forecast"]["Forecast Qty"].sum())*float(p.get("price") or 0),
            "Demand Value INR Mn":float(result["forecast"]["Forecast Qty"].sum())*float(p.get("price") or 0)/1_000_000,
            "Plan Status":result["plan_status"],
            "Price Confirmed":"Yes" if p.get("price_confirmed") else "Pending",
            "Confirmed":"Yes" if p["confirmed"] else "Pending"})
        f=result["forecast"].copy()
        f["Plant"]=rec["plant"]; f["Customer"]=rec.get("Customer"); f["Product"]=rec.get("Product")
        f["Part No."]=rec.get("Part No.")
        pprice=float(p.get("price") or 0)
        f["Unit Price INR"]=pprice
        f["Forecast Value INR"]=f["Forecast Qty"]*pprice
        f["Forecast Value INR Mn"]=f["Forecast Value INR"]/1_000_000
        forecasts.append(f)
    if unconfirmed:
        st.warning(f"{len(unconfirmed)} included products still need confirmation. Forecast values are provisional until confirmed.")
    if rows:
        summary=pd.DataFrame(rows)
        st.dataframe(summary,use_container_width=True,hide_index=True)
        allf=pd.concat(forecasts,ignore_index=True)
        st.markdown("#### Month-wise forecast")
        st.dataframe(allf.assign(Month=allf.Month.dt.strftime("%b %Y"))[
            ["Plant","Customer","Product","Part No.","Month","Forecast Qty","Unit Price INR","Forecast Value INR","Forecast Value INR Mn"]],
            use_container_width=True,hide_index=True,height=480)
        st.metric("Total forecast quantity",f'{allf["Forecast Qty"].sum():,.2f}')
    else:
        st.info("No products are currently included.")

with tabs[3]:
    st.subheader("Excel export")
    st.write("The export preserves every original tab and appends category forecast tabs for supported sheets, a consolidated All Plans Calculated tab, assumptions and the client/order register.")
    included=[r for r in records if st.session_state.params[r["key"]]["include"]]
    confirmed_count=sum(st.session_state.params[r["key"]]["confirmed"] for r in included)
    st.metric("Confirmed included products",f"{confirmed_count:,} / {len(included):,}")
    if included and confirmed_count < len(included):
        st.warning("You can export now, but unconfirmed assumptions will be clearly marked as Pending in the review sheet.")
    if openrouter_key:
        st.caption("The optional model receives only aggregate forecast totals, not customer/product-level source rows.")
        if st.button("Generate AI forecast commentary"):
            included_for_ai=[r for r in records if st.session_state.params[r["key"]]["include"]]
            total_qty=0.0
            plant_totals={}
            for r in included_for_ai:
                pp=st.session_state.params[r["key"]]
                ff=build_forecast(r,pp["growth"],horizon,start_month,pp.get("manual_base"),pp.get("manual_monthly"))["forecast"]
                total_qty += float(ff["Forecast Qty"].sum())
                plant_totals[r["plant"]]=plant_totals.get(r["plant"],0.0)+float(ff["Forecast Qty"].sum())
            prompt=("Write a concise operations-planning commentary for a demand forecast. "
                    "Do not invent facts. Explain this is a modelled plan, not verified actual demand. "
                    f"Horizon={horizon} years; start={start_month.date()}; total forecast quantity={total_qty:.2f}; "
                    f"plant totals={plant_totals}. Give risks and review checks.")
            try:
                response=requests.post("https://openrouter.ai/api/v1/chat/completions",
                    headers={"Authorization":f"Bearer {openrouter_key}","Content-Type":"application/json",
                             "HTTP-Referer":"https://streamlit.io","X-Title":"Demand Forecasting"},
                    json={"model":openrouter_model,"messages":[{"role":"user","content":prompt}],
                          "temperature":0.2,"max_tokens":500},timeout=45)
                response.raise_for_status()
                commentary=response.json()["choices"][0]["message"]["content"]
                st.session_state["ai_forecast_commentary"]=commentary
            except Exception as e:
                st.error(f"OpenRouter request failed: {e}")
    if st.session_state.get("ai_forecast_commentary"):
        st.markdown("#### AI planning commentary")
        st.write(st.session_state["ai_forecast_commentary"])
    if st.button("Build Excel forecast workbook",type="primary"):
        export_params={r["key"]:st.session_state.params[r["key"]] for r in records}
        data=excel_export(file_bytes,records,export_params,start_month,horizon,st.session_state.client_orders,source_name,sheet_names)
        st.download_button("Download forecast workbook (.xlsx)",data=data,
            file_name=f"Demand_Forecast_{start_month:%Y_%m}_{horizon}Y.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    st.markdown("**Workbook structure**")
    st.write("Original worksheets are retained. Forecast companion tabs are appended with category roll-ups; the final tab is `All Plans Calculated`.")
    st.caption("The source workbook is preserved. New products and deletion decisions affect the forecast output and are recorded in the separate downloadable workbook; original source rows are not physically erased.")
