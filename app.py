
import streamlit as st
import pandas as pd
import numpy as np
import requests
from pyxlsb import open_workbook
from datetime import datetime, timedelta

st.set_page_config(
    page_title="Predictive Fault Weather Engine v3",
    page_icon="⛈️",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
.block-container {padding-top: 1rem; padding-bottom: 1rem; max-width: 1280px;}
.main-title {font-size: 2rem; font-weight: 900; margin-bottom: 0;}
.subtitle {color:#6b7280; margin-bottom: 1rem;}
.card {
    border-radius: 20px;
    padding: 18px;
    background: #f9fafb;
    border: 1px solid #e5e7eb;
    margin-bottom: 14px;
    box-shadow: 0 6px 18px rgba(0,0,0,.04);
}
.risk-low {background:#ecfdf5;border-left:9px solid #10b981;}
.risk-medium {background:#fffbeb;border-left:9px solid #f59e0b;}
.risk-high {background:#fff1f2;border-left:9px solid #f43f5e;}
.risk-critical {background:#fee2e2;border-left:9px solid #991b1b;}
.big {font-size: 1.3rem; font-weight: 900;}
.small {font-size:.92rem;color:#6b7280;}
.value {font-weight:800;}
</style>
""", unsafe_allow_html=True)

WEATHER_POINTS = {
    "Αττική / Αθήνα κέντρο": (37.9842, 23.7281),
    "Αχαρνές": (38.0833, 23.7333),
    "Μαρούσι": (38.0500, 23.8000),
    "Χαλάνδρι": (38.0237, 23.8007),
    "Ηλιούπολη": (37.9315, 23.7606),
    "Πειραιάς": (37.9420, 23.6469),
    "Περιστέρι": (38.0154, 23.6919),
    "Κηφισιά": (38.0744, 23.8111),
    "Γλυφάδα": (37.8629, 23.7544),
    "Λαύριο": (37.7145, 24.0565),
}

def excel_serial_to_date(x):
    try:
        return (datetime(1899, 12, 30) + timedelta(days=int(float(x)))).date()
    except Exception:
        return None

def normalize_text(x):
    if x is None:
        return ""
    return str(x).strip()

def safe_float(x):
    try:
        if x is None or x == "":
            return np.nan
        return float(x)
    except Exception:
        return np.nan

def extract_total_daily_series(file_path):
    rows = []
    with open_workbook(file_path) as wb:
        if "Table" not in wb.sheets:
            return pd.DataFrame()
        sh = wb.get_sheet("Table")
        for row in sh.rows():
            vals = [c.v for c in row]
            # Expected Table H:L => date, completed, stock, inflow, not_completed
            if len(vals) >= 12 and isinstance(vals[7], (int, float)):
                d = excel_serial_to_date(vals[7])
                if d is None:
                    continue
                completed = safe_float(vals[8])
                stock = safe_float(vals[9])
                inflow = safe_float(vals[10])
                not_completed = safe_float(vals[11])
                if d.year >= 2024 and not np.isnan(inflow):
                    rows.append([d, completed, stock, inflow, not_completed])

    df = pd.DataFrame(rows, columns=["date", "completed", "stock", "inflow", "not_completed"])
    if df.empty:
        return df
    df = df.drop_duplicates(subset=["date"]).sort_values("date")
    df["date"] = pd.to_datetime(df["date"])
    df["weekday"] = df["date"].dt.day_name()
    return df

def extract_weight_table(file_path):
    """
    Reads daily weighted split from sheet 'Table', columns CL:CP-like:
    date, GroupA, GroupB, Stock W, Inflow W.
    This is used to isolate only GroupA='Βλάβη'.
    """
    rows = []
    with open_workbook(file_path) as wb:
        if "Table" not in wb.sheets:
            return pd.DataFrame()
        sh = wb.get_sheet("Table")
        for row in sh.rows():
            vals = [c.v for c in row]
            if len(vals) >= 94 and isinstance(vals[89], (int, float)):
                d = excel_serial_to_date(vals[89])
                if d is None:
                    continue
                group_a = normalize_text(vals[90])
                group_b = normalize_text(vals[91])
                stock_w = safe_float(vals[92])
                inflow_w = safe_float(vals[93])
                if d.year >= 2024 and group_a and group_b and not group_a.startswith("Group"):
                    rows.append([d, group_a, group_b, stock_w, inflow_w])

    df = pd.DataFrame(rows, columns=["date", "group_a", "group_b", "stock_weight", "inflow_weight"])
    if df.empty:
        return df
    df["date"] = pd.to_datetime(df["date"])
    return df.dropna(subset=["inflow_weight"], how="all")

def extract_clean_tasks(file_path):
    rows = []
    with open_workbook(file_path) as wb:
        if "Clean Tasks" not in wb.sheets:
            return pd.DataFrame()
        sh = wb.get_sheet("Clean Tasks")
        for i, row in enumerate(sh.rows()):
            vals = [c.v for c in row]
            if i == 0:
                continue
            if len(vals) >= 5 and vals[1] is not None:
                rows.append([
                    normalize_text(vals[0]),
                    normalize_text(vals[1]),
                    normalize_text(vals[2]),
                    normalize_text(vals[3]),
                    safe_float(vals[4]),
                ])
    return pd.DataFrame(rows, columns=["category", "task_type", "group_a", "group_b", "count"])

def extract_area_stock_table(file_path):
    """
    Extracts daily stock by broad technical area/department from the Table sheet.
    In the user's workbook this appears around columns AX:BH:
    Date + columns like ΤΠΒΕ - ΑΝΑΤΟΛΙΚΟ, ΤΠΝΕ - ΑΝΑΤ. ΑΤΤΙΚΗΣ, etc.
    """
    rows = []
    headers = None

    with open_workbook(file_path) as wb:
        if "Table" not in wb.sheets:
            return pd.DataFrame()
        sh = wb.get_sheet("Table")
        for i, row in enumerate(sh.rows()):
            vals = [c.v for c in row]

            # Header row: col 49 = " Stock", col 50+ = technical areas
            if len(vals) > 60 and normalize_text(vals[49]) == "Stock":
                h = []
                for c in range(50, min(64, len(vals))):
                    v = normalize_text(vals[c])
                    if v and v != "Γενικό Άθροισμα":
                        h.append((c, v))
                if h:
                    headers = h

            if headers and len(vals) > 60 and isinstance(vals[49], (int, float)):
                d = excel_serial_to_date(vals[49])
                if d is None:
                    continue
                for c, area in headers:
                    val = safe_float(vals[c]) if c < len(vals) else np.nan
                    if not np.isnan(val):
                        rows.append([d, area, val])

    df = pd.DataFrame(rows, columns=["date", "technical_area", "stock"])
    if df.empty:
        return df
    df["date"] = pd.to_datetime(df["date"])
    return df.drop_duplicates(subset=["date", "technical_area"]).sort_values(["technical_area", "date"])

def apply_scope(total_df, weights_df, mode, group_a_choice=None, group_b_choice=None):
    df = total_df.copy()
    df["scope_note"] = "Όλα όπως διαβάστηκαν από το pivot"

    if mode == "Όλα":
        return df

    if weights_df.empty:
        df["scope_note"] = "Δεν βρέθηκε weighted table — χρησιμοποιούνται όλα"
        return df

    w = weights_df.copy()

    if mode == "Μόνο Βλάβες":
        w = w[w["group_a"] == "Βλάβη"]
        note = "Μόνο GroupA = Βλάβη"
    elif mode == "Επιλογή GroupA":
        w = w[w["group_a"] == group_a_choice]
        note = f"GroupA = {group_a_choice}"
    elif mode == "Επιλογή GroupB":
        w = w[w["group_b"] == group_b_choice]
        note = f"GroupB = {group_b_choice}"
    else:
        note = "Custom filter"

    if w.empty:
        df["scope_note"] = f"Δεν βρέθηκαν βάρη για {note} — χρησιμοποιούνται όλα"
        return df

    agg = w.groupby("date", as_index=False).agg(
        inflow_weight=("inflow_weight", "sum"),
        stock_weight=("stock_weight", "sum")
    )

    df = df.merge(agg, on="date", how="left")
    df["inflow_weight"] = df["inflow_weight"].fillna(0)
    df["stock_weight"] = df["stock_weight"].fillna(df["inflow_weight"])

    df["inflow_total"] = df["inflow"]
    df["stock_total"] = df["stock"]
    df["completed_total"] = df["completed"]

    df["inflow"] = df["inflow"] * df["inflow_weight"]
    df["stock"] = df["stock"] * df["stock_weight"]
    df["completed"] = df["completed"] * df["inflow_weight"]
    df["not_completed"] = df["not_completed"] * df["inflow_weight"]
    df["scope_note"] = note
    return df

@st.cache_data(show_spinner=False)
def fetch_historical_weather(lat, lon, start_date, end_date):
    url = "https://archive-api.open-meteo.com/v1/archive"
    params = {
        "latitude": lat,
        "longitude": lon,
        "start_date": str(start_date),
        "end_date": str(end_date),
        "daily": "precipitation_sum,rain_sum,weather_code,wind_speed_10m_max",
        "timezone": "Europe/Athens",
    }
    r = requests.get(url, params=params, timeout=30)
    r.raise_for_status()
    data = r.json()["daily"]
    df = pd.DataFrame(data)
    df["date"] = pd.to_datetime(df["time"])
    return df.drop(columns=["time"])

@st.cache_data(show_spinner=False)
def fetch_7day_forecast(lat, lon):
    url = "https://api.open-meteo.com/v1/forecast"
    params = {
        "latitude": lat,
        "longitude": lon,
        "daily": "precipitation_sum,rain_sum,weather_code,wind_speed_10m_max,precipitation_probability_max",
        "forecast_days": 7,
        "timezone": "Europe/Athens",
    }
    r = requests.get(url, params=params, timeout=30)
    r.raise_for_status()
    data = r.json()["daily"]
    df = pd.DataFrame(data)
    df["date"] = pd.to_datetime(df["time"])
    return df.drop(columns=["time"])

def classify_weather(row, rain_medium, rain_high, wind_high):
    rain = float(row.get("precipitation_sum", 0) or 0)
    wind = float(row.get("wind_speed_10m_max", 0) or 0)
    prob = float(row.get("precipitation_probability_max", 0) or 0)
    code = int(row.get("weather_code", 0) or 0)

    score = 0
    reasons = []
    if rain >= 5:
        score += 10; reasons.append("βροχή")
    if rain >= rain_medium:
        score += 15; reasons.append("έντονη βροχή")
    if rain >= rain_high:
        score += 25; reasons.append("πολύ υψηλή βροχή")
    if wind >= wind_high:
        score += 15; reasons.append("ισχυρός άνεμος")
    if prob >= 70:
        score += 10; reasons.append("υψηλή πιθανότητα υετού")
    if code in [95, 96, 99]:
        score += 35; reasons.append("καταιγίδα / πιθανή ηλεκτρική δραστηριότητα")

    if score >= 65:
        return "CRITICAL", score, ", ".join(sorted(set(reasons)))
    if score >= 40:
        return "HIGH", score, ", ".join(sorted(set(reasons)))
    if score >= 20:
        return "MEDIUM", score, ", ".join(sorted(set(reasons)))
    return "LOW", score, "χωρίς έντονο φαινόμενο"

def risk_class_name(risk):
    return {
        "LOW": "risk-low",
        "MEDIUM": "risk-medium",
        "HIGH": "risk-high",
        "CRITICAL": "risk-critical",
    }.get(risk, "risk-low")

def analyze_history(df, rain_medium, wind_high, stock_recovery_pct):
    df = df.copy()
    df["baseline_inflow"] = df.groupby("weekday")["inflow"].transform("median")
    df["inflow_delta"] = df["inflow"] - df["baseline_inflow"]
    df["inflow_delta_pct"] = np.where(df["baseline_inflow"] > 0, df["inflow_delta"] / df["baseline_inflow"], 0)
    df["stock_prev"] = df["stock"].shift(1)
    df["stock_delta"] = df["stock"] - df["stock_prev"]

    df["bad_weather"] = (
        (df["precipitation_sum"].fillna(0) >= rain_medium) |
        (df["wind_speed_10m_max"].fillna(0) >= wind_high) |
        (df["weather_code"].isin([95, 96, 99]))
    )

    median_inflow = df["inflow"].median()
    median_stock = df["stock"].median()
    normal_inflow_threshold = median_inflow * 1.10
    normal_stock_threshold = median_stock * (1 + stock_recovery_pct / 100)

    bad_dates = sorted(set(df.loc[df["bad_weather"], "date"].dt.date))
    events = []

    for bad_day in bad_dates:
        window = df[(df["date"].dt.date >= bad_day) & (df["date"].dt.date <= bad_day + timedelta(days=7))].copy()
        after = window[window["date"].dt.date > bad_day]
        if window.empty:
            continue

        bad_row = df[df["date"].dt.date == bad_day].iloc[0]
        elevated_stock_days = int((after["stock"] > normal_stock_threshold).sum())
        recovery_rows = after[
            (after["inflow"] <= normal_inflow_threshold) &
            (after["stock"] <= normal_stock_threshold)
        ]

        recovery_days = None
        if not recovery_rows.empty:
            recovery_days = int((recovery_rows.iloc[0]["date"].date() - bad_day).days)

        event_risk = "LOW"
        if bad_row.get("precipitation_sum", 0) >= rain_medium or bad_row.get("wind_speed_10m_max", 0) >= wind_high:
            event_risk = "MEDIUM"
        if bad_row.get("precipitation_sum", 0) >= rain_medium * 1.8:
            event_risk = "HIGH"
        if bad_row.get("weather_code", 0) in [95, 96, 99]:
            event_risk = "HIGH"

        events.append({
            "Ημερομηνία κακοκαιρίας": bad_day,
            "Risk": event_risk,
            "Βροχή mm": round(float(bad_row.get("precipitation_sum", 0) or 0), 1),
            "Άνεμος km/h": round(float(bad_row.get("wind_speed_10m_max", 0) or 0), 1),
            "Weather Code": int(bad_row.get("weather_code", 0) or 0),
            "Εισροή ημέρας": round(float(bad_row["inflow"]), 0),
            "Συνήθης εισροή": round(float(bad_row["baseline_inflow"]), 0),
            "Αύξηση εισροής %": round(float(bad_row["inflow_delta_pct"] * 100), 1),
            "Μέγιστη εισροή επόμενων 7 ημερών": round(float(window["inflow"].max()), 0),
            "Μέγιστο stock επόμενων 7 ημερών": round(float(window["stock"].max()), 0),
            "Μέρες αυξημένου stock μετά": elevated_stock_days,
            "Μέρες μέχρι επιστροφή": recovery_days,
        })

    return df, pd.DataFrame(events)

def build_forecast_prediction(forecast, history, rain_medium, rain_high, wind_high):
    hist = history.copy()
    bad = hist[hist["bad_weather"]]
    normal = hist[~hist["bad_weather"]]

    normal_inflow = normal["inflow"].median() if not normal.empty else hist["inflow"].median()
    median_stock = hist["stock"].median()
    completion_capacity = hist["completed"].median()

    high_uplift = max(0.0, bad["inflow_delta_pct"].median()) if not bad.empty else 0.20

    rows = []
    projected_stock = float(hist.sort_values("date").iloc[-1]["stock"])

    for _, row in forecast.iterrows():
        risk, score, reasons = classify_weather(row, rain_medium, rain_high, wind_high)

        if risk == "CRITICAL":
            uplift = max(high_uplift, 0.35)
        elif risk == "HIGH":
            uplift = max(high_uplift, 0.22)
        elif risk == "MEDIUM":
            uplift = 0.10
        else:
            uplift = 0.00

        expected_inflow = normal_inflow * (1 + uplift)
        projected_stock = max(0, projected_stock + expected_inflow - completion_capacity)

        if projected_stock <= median_stock:
            recovery_days = 0
        else:
            daily_recovery_power = max(completion_capacity - normal_inflow, 1)
            recovery_days = int(np.ceil((projected_stock - median_stock) / daily_recovery_power))

        rows.append({
            "Ημερομηνία": row["date"].date(),
            "Risk": risk,
            "Risk Score": score,
            "Αιτία": reasons,
            "Βροχή mm": round(float(row.get("precipitation_sum", 0) or 0), 1),
            "Πιθανότητα %": round(float(row.get("precipitation_probability_max", 0) or 0), 0),
            "Άνεμος km/h": round(float(row.get("wind_speed_10m_max", 0) or 0), 1),
            "Εκτίμηση εισροής": round(expected_inflow, 0),
            "Αύξηση vs normal %": round(uplift * 100, 1),
            "Projected Stock": round(projected_stock, 0),
            "Εκτίμηση απορρόφησης ημέρες": recovery_days,
        })
    return pd.DataFrame(rows)

def build_area_risk_table(area_stock, forecast_risk, forecast_score):
    if area_stock.empty:
        return pd.DataFrame()

    rows = []
    latest_date = area_stock["date"].max()
    recent = area_stock[area_stock["date"] == latest_date].copy()

    for area in sorted(area_stock["technical_area"].unique()):
        s = area_stock[area_stock["technical_area"] == area].sort_values("date")
        latest = recent.loc[recent["technical_area"] == area, "stock"]
        if latest.empty:
            continue
        latest_stock = float(latest.iloc[0])
        median_stock = float(s["stock"].median())
        p75 = float(s["stock"].quantile(0.75))
        p90 = float(s["stock"].quantile(0.90))

        pressure_pct = ((latest_stock / median_stock) - 1) * 100 if median_stock else 0
        score = 0
        if latest_stock >= p75:
            score += 15
        if latest_stock >= p90:
            score += 20
        if pressure_pct >= 10:
            score += 15
        if pressure_pct >= 20:
            score += 20
        score += min(int(forecast_score * 0.6), 45)

        if score >= 65:
            risk = "CRITICAL"
        elif score >= 40:
            risk = "HIGH"
        elif score >= 20:
            risk = "MEDIUM"
        else:
            risk = "LOW"

        rows.append({
            "Τεχνική Περιοχή / Τμήμα": area,
            "Τελευταίο stock": round(latest_stock, 0),
            "Συνήθες stock": round(median_stock, 0),
            "Πίεση vs normal %": round(pressure_pct, 1),
            "Forecast Risk": forecast_risk,
            "Area Risk": risk,
            "Area Score": score,
            "Σχόλιο": (
                "Υψηλή προτεραιότητα παρακολούθησης" if risk in ["HIGH", "CRITICAL"]
                else "Κανονική παρακολούθηση"
            )
        })

    out = pd.DataFrame(rows)
    if out.empty:
        return out

    order = {"CRITICAL": 4, "HIGH": 3, "MEDIUM": 2, "LOW": 1}
    out["_order"] = out["Area Risk"].map(order)
    return out.sort_values(["_order", "Area Score"], ascending=False).drop(columns=["_order"])

st.markdown('<p class="main-title">⛈️ Predictive Fault Weather Engine v3</p>', unsafe_allow_html=True)
st.markdown('<p class="subtitle">Ιστορικές κακοκαιρίες / Μόνο βλάβες / Περιοχές-Τεχνικά Τμήματα / Πρόγνωση 7 ημερών</p>', unsafe_allow_html=True)

uploaded = st.file_uploader("Ανέβασε το Chart.v4.xlsb", type=["xlsb"])

with st.sidebar:
    st.header("⚙️ Επιλογές")
    weather_area = st.selectbox("Περιοχή καιρού", list(WEATHER_POINTS.keys()), index=0)
    lat, lon = WEATHER_POINTS[weather_area]

    st.markdown("---")
    rain_medium = st.slider("Όριο έντονης βροχής mm", 5, 30, 10)
    rain_high = st.slider("Όριο υψηλής βροχής mm", 10, 60, 20)
    wind_high = st.slider("Όριο ισχυρού ανέμου km/h", 25, 80, 45)
    stock_recovery_pct = st.slider("Όριο φυσιολογικού stock +%", 0, 20, 5)

if not uploaded:
    st.info("Ανέβασε το Chart.v4.xlsb για να ξεκινήσει η ανάλυση.")
    st.stop()

temp_path = "uploaded_chart.xlsb"
with open(temp_path, "wb") as f:
    f.write(uploaded.getbuffer())

try:
    total = extract_total_daily_series(temp_path)
    weights = extract_weight_table(temp_path)
    clean_tasks = extract_clean_tasks(temp_path)
    area_stock = extract_area_stock_table(temp_path)

    if total.empty:
        st.error("Δεν βρέθηκε daily operational series στο sheet Table.")
        st.stop()

    group_as = sorted(weights["group_a"].dropna().unique().tolist()) if not weights.empty else []
    group_bs = sorted(weights["group_b"].dropna().unique().tolist()) if not weights.empty else []

    st.sidebar.markdown("---")
    st.sidebar.header("🎯 Φίλτρο εργασιών")
    mode = st.sidebar.selectbox(
        "Scope δεδομένων",
        ["Μόνο Βλάβες", "Όλα", "Επιλογή GroupA", "Επιλογή GroupB"],
        index=0
    )

    group_a_choice = None
    group_b_choice = None
    if mode == "Επιλογή GroupA":
        group_a_choice = st.sidebar.selectbox("GroupA", group_as if group_as else ["Βλάβη"])
    if mode == "Επιλογή GroupB":
        group_b_choice = st.sidebar.selectbox("GroupB", group_bs if group_bs else ["LLU", "Cooper", "FTTH", "FWA"])

    scoped = apply_scope(total, weights, mode, group_a_choice, group_b_choice)
    scope_note = scoped["scope_note"].iloc[0] if "scope_note" in scoped.columns else mode

    start_d = scoped["date"].min().date()
    end_d = scoped["date"].max().date()

    with st.spinner(f"Διαβάζω από {start_d} έως {end_d}, τραβάω καιρό και χτίζω dashboard..."):
        weather = fetch_historical_weather(lat, lon, start_d, end_d)
        forecast = fetch_7day_forecast(lat, lon)

    merged = scoped.merge(weather, on="date", how="left")
    analyzed, events = analyze_history(merged, rain_medium, wind_high, stock_recovery_pct)
    prediction = build_forecast_prediction(forecast, analyzed, rain_medium, rain_high, wind_high)

    top = prediction.iloc[0]
    area_risk = build_area_risk_table(area_stock, top["Risk"], top["Risk Score"])

    tab1, tab2, tab3, tab4 = st.tabs([
        "🏠 Dashboard",
        "⛈️ Ιστορικές Κακοκαιρίες",
        "📍 Περιοχές / Τεχνικά Τμήματα",
        "🔎 Data Check"
    ])

    with tab1:
        st.markdown(f"""
        <div class="card {risk_class_name(top['Risk'])}">
            <div class="big">Τρέχουσα σήμανση: {top['Risk']}</div>
            <div>Scope: <span class="value">{scope_note}</span> | Περιοχή καιρού: <span class="value">{weather_area}</span></div>
            <div>Ιστορικό από: <span class="value">{start_d}</span> έως <span class="value">{end_d}</span></div>
            <br>
            <div>Αιτία: <span class="value">{top['Αιτία']}</span></div>
            <div>Εκτίμηση εισροής: <span class="value">{top['Εκτίμηση εισροής']:.0f}</span>
            ({top['Αύξηση vs normal %']:.1f}% vs normal)</div>
            <div>Projected stock: <span class="value">{top['Projected Stock']:.0f}</span></div>
            <div>Εκτίμηση απορρόφησης: <span class="value">{top['Εκτίμηση απορρόφησης ημέρες']} ημέρες</span></div>
        </div>
        """, unsafe_allow_html=True)

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Ημέρες ιστορικού", len(analyzed))
        c2.metric("Συνήθης εισροή", f"{analyzed['inflow'].median():.0f}")
        c3.metric("Συνήθες stock", f"{analyzed['stock'].median():.0f}")
        c4.metric("Ημέρες κακοκαιρίας", len(events))

        st.subheader("🔔 Πρόγνωση 7 ημερών")
        st.dataframe(prediction, use_container_width=True)

        st.subheader("📈 Ιστορική εικόνα εισροής / ολοκληρώσεων / stock")
        st.line_chart(analyzed.set_index("date")[["inflow", "completed", "stock"]])

    with tab2:
        st.subheader("⛈️ Αναδρομή σε ημέρες κακοκαιρίας")
        if events.empty:
            st.warning("Δεν βρέθηκαν ημέρες κακοκαιρίας με τα thresholds που έχεις επιλέξει.")
        else:
            event_labels = [
                f"{r['Ημερομηνία κακοκαιρίας']} | {r['Risk']} | {r['Βροχή mm']}mm | +{r['Αύξηση εισροής %']}%"
                for _, r in events.sort_values("Ημερομηνία κακοκαιρίας", ascending=False).iterrows()
            ]

            selected_label = st.selectbox("Επίλεξε ημέρα κακοκαιρίας", event_labels)
            selected_date = selected_label.split(" | ")[0]
            ev = events[events["Ημερομηνία κακοκαιρίας"].astype(str) == selected_date].iloc[0]
            d0 = pd.to_datetime(selected_date)

            cols = st.columns(5)
            cols[0].metric("Βροχή", f"{ev['Βροχή mm']} mm")
            cols[1].metric("Άνεμος", f"{ev['Άνεμος km/h']} km/h")
            cols[2].metric("Αύξηση εισροής", f"{ev['Αύξηση εισροής %']}%")
            cols[3].metric("Αυξημένο stock μετά", f"{ev['Μέρες αυξημένου stock μετά']} μέρες")
            cols[4].metric("Επιστροφή", f"{ev['Μέρες μέχρι επιστροφή']} μέρες" if pd.notna(ev["Μέρες μέχρι επιστροφή"]) else "Δεν βρέθηκε")

            st.markdown(f"""
            <div class="card">
                <div class="big">Παράδειγμα παρουσίασης</div>
                <p>
                Στις <b>{selected_date}</b> καταγράφηκε κακοκαιρία με <b>{ev['Βροχή mm']}mm</b> βροχής.
                Η εισροή ήταν <b>{ev['Εισροή ημέρας']:.0f}</b> έναντι συνήθους εισροής <b>{ev['Συνήθης εισροή']:.0f}</b>,
                δηλαδή μεταβολή <b>{ev['Αύξηση εισροής %']}%</b>.
                Το μέγιστο stock των επόμενων 7 ημερών έφτασε <b>{ev['Μέγιστο stock επόμενων 7 ημερών']:.0f}</b>.
                </p>
            </div>
            """, unsafe_allow_html=True)

            window = analyzed[(analyzed["date"] >= d0) & (analyzed["date"] <= d0 + pd.Timedelta(days=7))].copy()
            show_cols = ["date", "precipitation_sum", "wind_speed_10m_max", "inflow", "baseline_inflow", "stock", "completed", "stock_delta"]
            st.subheader("Παράθυρο 7 ημερών μετά την κακοκαιρία")
            st.dataframe(window[show_cols], use_container_width=True)
            st.line_chart(window.set_index("date")[["inflow", "baseline_inflow", "stock"]])

            st.subheader("Όλες οι ημέρες κακοκαιρίας")
            st.dataframe(events.sort_values("Ημερομηνία κακοκαιρίας", ascending=False), use_container_width=True)

    with tab3:
        st.subheader("📍 Περιοχές / Τεχνικά Τμήματα που χρειάζονται προσοχή")
        if area_risk.empty:
            st.warning("Δεν βρέθηκε daily stock ανά τεχνική περιοχή στο αρχείο. Για ΤΤΛΠ-level πρόβλεψη θα χρειαστεί raw export ανά Τεχνικό Τμήμα.")
        else:
            c1, c2 = st.columns([1, 2])
            with c1:
                selected_area = st.selectbox("Επιλογή περιοχής/τεχνικού τμήματος", area_risk["Τεχνική Περιοχή / Τμήμα"].tolist())
            with c2:
                area_row = area_risk[area_risk["Τεχνική Περιοχή / Τμήμα"] == selected_area].iloc[0]
                st.markdown(f"""
                <div class="card {risk_class_name(area_row['Area Risk'])}">
                    <div class="big">{selected_area}: {area_row['Area Risk']}</div>
                    <div>Τελευταίο stock: <span class="value">{area_row['Τελευταίο stock']:.0f}</span></div>
                    <div>Συνήθες stock: <span class="value">{area_row['Συνήθες stock']:.0f}</span></div>
                    <div>Πίεση vs normal: <span class="value">{area_row['Πίεση vs normal %']}%</span></div>
                    <div>Σχόλιο: <span class="value">{area_row['Σχόλιο']}</span></div>
                </div>
                """, unsafe_allow_html=True)

            st.dataframe(area_risk, use_container_width=True)

            selected_series = area_stock[area_stock["technical_area"] == selected_area].sort_values("date")
            if not selected_series.empty:
                st.line_chart(selected_series.set_index("date")[["stock"]])

            st.caption("Σημείωση: εδώ χρησιμοποιείται το daily stock ανά ευρύτερη τεχνική περιοχή που υπάρχει στο workbook. Για πραγματικό ΤΤΛΠ/ΑΚ prediction χρειάζεται raw export με στήλες ΤΤΛΠ και Α/Κ.")

    with tab4:
        st.subheader("🔎 Έλεγχος δεδομένων από αρχείο")
        st.write("Scope:", scope_note)
        st.write("Παλαιότερη ημερομηνία που διαβάστηκε:", start_d)
        st.write("Τελευταία ημερομηνία που διαβάστηκε:", end_d)

        st.write("GroupA που βρέθηκαν:", group_as)
        st.write("GroupB που βρέθηκαν:", group_bs)

        if not clean_tasks.empty:
            st.subheader("Clean Tasks / ταξινόμηση ειδών εργασίας")
            st.dataframe(clean_tasks, use_container_width=True)

        if not area_stock.empty:
            st.subheader("Περιοχές που βρέθηκαν στο stock table")
            st.dataframe(area_stock.head(100), use_container_width=True)

        st.download_button(
            "Κατέβασε πλήρη merged analysis CSV",
            analyzed.to_csv(index=False).encode("utf-8-sig"),
            file_name="fault_weather_analysis_v3.csv",
            mime="text/csv"
        )

        if not events.empty:
            st.download_button(
                "Κατέβασε bad weather events CSV",
                events.to_csv(index=False).encode("utf-8-sig"),
                file_name="bad_weather_events_v3.csv",
                mime="text/csv"
            )

except Exception as e:
    st.error(f"Σφάλμα: {e}")
    st.caption("Έλεγξε ότι το αρχείο είναι το σωστό Chart.v4.xlsb και ότι υπάρχει πρόσβαση internet για Open-Meteo.")
