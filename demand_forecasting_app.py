import streamlit as st
import pandas as pd
import numpy as np
import openpyxl
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
st.caption("Product-level planning • plant-wise review • client and order control • Excel export")

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
    for sheet in PLANT_SHEETS:
        if sheet not in wb.sheetnames:
            issues.append(f"Missing plant sheet: {sheet}")
            continue
        ws = wb[sheet]
        date_cols = []
        for c in range(MONTH_START_COL, min(ws.max_column, MONTH_END_COL)+1):
            d = ws.cell(6,c).value
            if isinstance(d, (datetime, pd.Timestamp)):
                date_cols.append((c, pd.Timestamp(d).to_period("M").to_timestamp()))
        if not date_cols:
            issues.append(f"{sheet}: no monthly date headers found in row 6 (AJ:CQ).")
            continue
        for r in range(11, ws.max_row+1):
            customer = ws.cell(r,3).value
            part = ws.cell(r,6).value
            product = ws.cell(r,9).value
            model = ws.cell(r,4).value
            if all(v is None or str(v).strip()=="" for v in [customer,part,product,model]):
                continue
            vals = [numeric(ws.cell(r,c).value) for c,_ in date_cols]
            month_df = pd.DataFrame({"month":[d for _,d in date_cols],"quantity":vals})
            month_df = month_df.dropna(subset=["quantity"])
            # Preserve zero months as valid observations; remove only blank/non-numeric cells.
            meta = {k:ws.cell(r,c).value for k,c in META_COLS.items()}
            records.append({"plant":sheet,"source_row":r,"status":ws.cell(r,1).value,
                            "serial":ws.cell(r,2).value, **meta, "monthly":month_df})
    return records, issues, wb.sheetnames

def build_forecast(rec, growth, horizon_years, start_month, confirm=True):
    hist = rec["monthly"].copy()
    hist = hist[hist["month"] < start_month].sort_values("month")
    if hist.empty:
        hist = rec["monthly"].sort_values("month").copy()
    last12 = hist.tail(12)
    last3 = hist.tail(3)
    avg12 = float(last12.quantity.mean()) if len(last12) else 0.0
    avg3 = float(last3.quantity.mean()) if len(last3) else avg12
    blended = (avg12 + avg3) / 2
    if len(last12):
        seasonal = last12.copy()
        seasonal["month_no"] = seasonal.month.dt.month
        pattern = seasonal.groupby("month_no").quantity.mean().to_dict()
        pattern_mean = float(np.mean(list(pattern.values()))) if pattern else 0
    else:
        pattern, pattern_mean = {}, 0
    dates = pd.date_range(start=start_month, periods=12*horizon_years, freq="MS")
    rows=[]
    for i,d in enumerate(dates):
        # Retain the latest 12-month seasonal profile; normalize it to the blended run-rate.
        raw = pattern.get(d.month, blended)
        season_factor = (raw / pattern_mean) if pattern_mean and np.isfinite(raw) else 1.0
        annual_growth_factor = (1 + growth/100) ** ((i//12)+1)
        qty = blended * season_factor * annual_growth_factor
        rows.append({"Month":d,"Forecast Qty":max(0,qty)})
    return {"avg12":avg12,"avg3":avg3,"blended":blended,"n12":len(last12),"n3":len(last3),
            "forecast":pd.DataFrame(rows),"history_end":hist.month.max() if len(hist) else None}

def product_label(r):
    return f'{r["plant"]} | {r.get("Customer") or "No client"} | {r.get("Product") or r.get("Model") or "Unnamed"} | {r.get("Part No.") or "No part"}'

def excel_export(source_bytes, products, params, start_month, years, client_orders, source_name):
    # Preserve every original worksheet, formula, and workbook layout; append forecast outputs.
    wb = openpyxl.load_workbook(BytesIO(source_bytes), data_only=False)
    def add_sheet(name, headers, rows):
        if name in wb.sheetnames:
            del wb[name]
        ws = wb.create_sheet(name)
        ws.append(headers)
        for row in rows:
            ws.append([v.to_pydatetime() if isinstance(v, pd.Timestamp) else v for v in row])
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        for cell in ws[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="172B4D")
            cell.alignment = Alignment(wrap_text=True, vertical="center")
        ws.row_dimensions[1].height = 32
        for col in ws.columns:
            letter = get_column_letter(col[0].column)
            lengths = [len(str(c.value)) for c in list(col)[:250] if c.value is not None]
            ws.column_dimensions[letter].width = min(max(max(lengths, default=10)+2, 12), 34)
        return ws

    assumption_headers = ["Plant","Source Row","Status","Customer","Model","Vehicle / Nameplate","Part No.",
        "SOP","SAP Code","Product","Segment","Material","Material Description","Specifications","Paint",
        "Machining","Machined Description","Casting Weight","AfterMC Weight","Yield",
        "FY25 Melting kg/pcs","FY26 Melting kg/pcs","FY27 Melting kg/pcs","FY28 Melting kg/pcs","FY29 Melting kg/pcs",
        "FY25 Yield","FY26 Yield","FY27 Yield","FY28 Yield","FY29 Yield","Melting","Casting","AfterMC","Mold","Credit Period Days",
        "Planned Annual Growth %","Trailing 12M Source Mean","Trailing 3M Source Mean",
        "Blended Monthly Base","12M Observations","3M Observations","Reviewed / Confirmed","Include"]
    assumption_rows, monthly_rows = [], []
    plant_pivots = {"Manesar": [], "Gujrat": []}
    for rec in products:
        p = params[rec["key"]]
        result = build_forecast(rec, p["growth"], years, start_month)
        assumption_rows.append([rec["plant"],rec["source_row"],rec.get("status"),rec.get("Customer"),
            rec.get("Model"),rec.get("Vehicle / Nameplate"),rec.get("Part No."),rec.get("SOP"),
            rec.get("SAP Code"),rec.get("Product"),rec.get("Segment"),rec.get("Material"),
            rec.get("Material Description"),rec.get("Specifications"),rec.get("Paint"),rec.get("Machining"),
            rec.get("Machined Description"),rec.get("Casting Weight"),rec.get("AfterMC Weight"),rec.get("Yield"),
            rec.get("FY25 Melting kg/pcs"),rec.get("FY26 Melting kg/pcs"),rec.get("FY27 Melting kg/pcs"),rec.get("FY28 Melting kg/pcs"),rec.get("FY29 Melting kg/pcs"),
            rec.get("FY25 Yield"),rec.get("FY26 Yield"),rec.get("FY27 Yield"),rec.get("FY28 Yield"),rec.get("FY29 Yield"),
            rec.get("Melting"),rec.get("Casting"),rec.get("AfterMC"),rec.get("Mold"),rec.get("Credit Period Days"),
            p["growth"],result["avg12"],result["avg3"],result["blended"],result["n12"],result["n3"],
            "Yes" if p["confirmed"] else "Pending","Yes" if p["include"] else "No"])
        if p["include"]:
            month_map = {}
            for _, row in result["forecast"].iterrows():
                month_map[row["Month"]] = row["Forecast Qty"]
                monthly_rows.append([rec["plant"],rec.get("Customer"),rec.get("Model"),rec.get("Product"),
                    rec.get("Part No."),row["Month"].to_pydatetime(),row["Forecast Qty"],p["growth"],
                    "Yes" if p["confirmed"] else "Pending"])
            plant_pivots[rec["plant"]].append((rec, month_map))

    add_sheet("Forecast Assumptions", assumption_headers, assumption_rows)
    add_sheet("Forecast Monthly", ["Plant","Customer","Model","Product","Part No.","Month",
        "Forecast Qty","Annual Growth %","Confirmation"], monthly_rows)
    for plant, items in plant_pivots.items():
        dates = sorted({d for _, mp in items for d in mp})
        headers = ["Customer","Model","Vehicle / Nameplate","Product","Part No."] + dates
        rows = []
        for rec, mp in items:
            rows.append([rec.get("Customer"),rec.get("Model"),rec.get("Vehicle / Nameplate"),
                         rec.get("Product"),rec.get("Part No.")] + [mp.get(d,0) for d in dates])
        add_sheet(("Forecast "+plant)[:31], headers, rows)
    order_headers = ["Client","Plant","Order / Project","Model","Product","Part No.","Monthly Qty",
                     "Action","Recorded At"]
    order_rows = [[x.get(h) for h in order_headers] for x in client_orders]
    add_sheet("Client & Order Register", order_headers, order_rows)
    add_sheet("Forecast Read Me", ["Item","Details"], [
        ["Source Workbook",source_name],
        ["Forecast Start",start_month.to_pydatetime()],
        ["Forecast Horizon (years)",years],
        ["Baseline method","50% trailing 12-month source mean + 50% trailing 3-month source mean."],
        ["Seasonality","Latest available 12-month source pattern is normalized to the blended monthly base."],
        ["Growth","Planned annual growth is compounded by forecast year and applied month-wise."],
        ["Source caveat","The uploaded plant sheets label the monthly data as Plan. These are source-plan averages, not independently verified actual sales."],
        ["Review","Check data counts and confirm each product assumption before treating the forecast as approved."],
        ["Original sheets","All original workbook sheets are retained; forecast sheets are appended."]
    ])
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()

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
    st.session_state.params.setdefault(r["key"],{"growth":0.0,"confirmed":False,"include":True})
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

visible=[r for r in records if r["plant"] in plant_filter
         and (not customer_filter or str(r.get("Customer")) in customer_filter)
         and (not search or search.lower() in " ".join(str(r.get(k) or "") for k in ["Customer","Model","Product","Part No.","SAP Code"]).lower())]

st.caption(f"Source: {source_name} · {len(records):,} product rows detected across plant sheets")
st.warning("Source-data note: the Manesar and Gujrat monthly columns are labelled “Plan” in the uploaded workbook. The displayed trailing averages therefore reflect the available source-plan values; they should not be described as actual sales unless you replace them with actuals.")
if issues:
    for issue in issues: st.warning(issue)
tabs=st.tabs(["1 · Product assumptions","2 · Client & order register","3 · Forecast review","4 · Export"])

with tabs[0]:
    st.subheader("Product-level assumptions")
    st.write("For each product, review the source-derived averages before confirming its growth assumption. Growth is entered as an annual percentage.")
    if not visible:
        st.warning("No products match the current filters.")
    for rec in visible:
        p=st.session_state.params[rec["key"]]
        preview=build_forecast(rec,p["growth"],horizon,start_month)
        label=product_label(rec)
        with st.expander(label,expanded=False):
            st.markdown("**Source product parameters — review before confirmation**")
            source_fields = {k:v for k,v in rec.items() if k not in ["monthly","key"]}
            st.dataframe(pd.DataFrame([source_fields]), use_container_width=True, hide_index=True)
            a,b,c,d=st.columns(4)
            a.metric("Trailing 12-month source mean",f'{preview["avg12"]:,.2f}')
            b.metric("Trailing 3-month source mean",f'{preview["avg3"]:,.2f}')
            c.metric("Blended monthly base",f'{preview["blended"]:,.2f}')
            d.metric("Data points",f'{preview["n12"]} / {preview["n3"]} months')
            st.caption("Averages are calculated from the latest available monthly values before the forecast start date. If source months are incomplete, the available count is shown.")
            with st.form(f'form_{rec["key"]}'):
                growth=st.number_input("Planned annual growth (%)",min_value=-100.0,max_value=500.0,
                    value=float(p["growth"]),step=0.5,key=f'g_{rec["key"]}')
                include=st.checkbox("Include product in forecast",value=bool(p["include"]),key=f'i_{rec["key"]}')
                confirmed=st.checkbox("I have reviewed the averages and confirm this product assumption",
                    value=bool(p["confirmed"]),key=f'c_{rec["key"]}')
                submitted=st.form_submit_button("Save product assumption")
                if submitted:
                    st.session_state.params[rec["key"]]={"growth":growth,"include":include,"confirmed":confirmed}
                    st.success("Assumption saved.")
            st.dataframe(preview["forecast"].assign(Month=lambda x:x.Month.dt.strftime("%b %Y")).head(12),
                         use_container_width=True,hide_index=True)

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
        status=st.selectbox("Action",["Add","Delete"])
        add=st.form_submit_button("Add register entry")
        if add:
            entry={"Client":client,"Plant":plant,"Order / Project":order,
                "Model":model,"Product":product,"Part No.":part,"Monthly Qty":qty,"Action":status,
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
        result=build_forecast(rec,p["growth"],horizon,start_month)
        rows.append({"Plant":rec["plant"],"Customer":rec.get("Customer"),"Product":rec.get("Product"),
            "Part No.":rec.get("Part No."),"Growth %":p["growth"],"Last 12M Avg":result["avg12"],
            "Last 3M Avg":result["avg3"],"Blended Base":result["blended"],
            "Confirmed":"Yes" if p["confirmed"] else "Pending"})
        f=result["forecast"].copy()
        f["Plant"]=rec["plant"]; f["Customer"]=rec.get("Customer"); f["Product"]=rec.get("Product")
        f["Part No."]=rec.get("Part No.")
        forecasts.append(f)
    if unconfirmed:
        st.warning(f"{len(unconfirmed)} included products still need confirmation. Forecast values are provisional until confirmed.")
    if rows:
        summary=pd.DataFrame(rows)
        st.dataframe(summary,use_container_width=True,hide_index=True)
        allf=pd.concat(forecasts,ignore_index=True)
        st.markdown("#### Month-wise forecast")
        st.dataframe(allf.assign(Month=allf.Month.dt.strftime("%b %Y"))[
            ["Plant","Customer","Product","Part No.","Month","Forecast Qty"]],
            use_container_width=True,hide_index=True,height=480)
        st.metric("Total forecast quantity",f'{allf["Forecast Qty"].sum():,.2f}')
    else:
        st.info("No products are currently included.")

with tabs[3]:
    st.subheader("Excel export")
    st.write("The export contains assumptions and average calculations, month-wise forecast rows, plant-wise forecast tabs, and the client/order register.")
    included=[r for r in records if st.session_state.params[r["key"]]["include"]]
    confirmed_count=sum(st.session_state.params[r["key"]]["confirmed"] for r in included)
    st.metric("Confirmed included products",f"{confirmed_count:,} / {len(included):,}")
    if included and confirmed_count < len(included):
        st.warning("You can export now, but unconfirmed assumptions will be clearly marked as Pending in the review sheet.")
    if st.button("Build Excel forecast workbook",type="primary"):
        export_params={r["key"]:st.session_state.params[r["key"]] for r in records}
        data=excel_export(file_bytes,records,export_params,start_month,horizon,st.session_state.client_orders,source_name)
        st.download_button("Download forecast workbook (.xlsx)",data=data,
            file_name=f"Demand_Forecast_{start_month:%Y_%m}_{horizon}Y.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    st.caption("The source workbook is preserved. New products and deletion decisions affect the forecast output and are recorded in the separate downloadable workbook; original source rows are not physically erased.")
