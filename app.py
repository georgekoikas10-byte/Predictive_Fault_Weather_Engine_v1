
import streamlit as st
import pandas as pd
import numpy as np
import requests
from pyxlsb import open_workbook
from datetime import datetime, timedelta

st.set_page_config(
    page_title="Predictive Fault Weather Engine",
    page_icon="⛈️",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
.block-container {padding-top: 1rem; padding-bottom: 1rem; max-width: 1200px;}
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
.label {color:#6b7280;font-size:.9rem;}
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

def extract_total_daily_series(file_path):
    """
    Reads main daily total series from sheet 'Table', columns H:L:
    date, completed, stock, inflow, not_completed.
    """
    rows = []
    with open_workbook(file_path) as wb:
        if "Table" not in wb.sheets:
            return pd.DataFrame()
        sh = wb.get_sheet("Table")
        for row in sh.rows():
            vals = [c.v for c in row]
            if len(vals) >= 12 and isinstance(vals[7], (int, float)):
                d = excel_serial_to_date(vals[7])
                if d is None:
                    continue
                try:
                    completed = float(vals[8]) if vals[8] is not None else np.nan
                    stock = float(vals[9]) if vals[9] is not None else np.nan
                    inflow = float(vals[10]) if vals[10] is not None else np.nan
                    not_completed = float(vals[11]) if vals[11] is not None else np.nan
                    if d.year >= 2024 and not np.isnan(inflow):
                        rows.append([d, completed, stock, inflow, not_completed])
                except Exception:
                    pass

    df = pd.DataFrame(rows, columns=["date", "completed", "stock", "inflow", "not_completed"])
    if df.empty:
        return df
    df = df.drop_duplicates(subset=["date"]).sort_values("date")
    df["date"] = pd.to_datetime(df["date"])
    df["weekday"] = df["date"].dt.day_name()
    return df

def extract_weight_table(file_path):
    """
    Reads daily weighted split from sheet 'Table', columns CL:CP approximately:
    date, GroupA, GroupB, Stock W, Sum of Today W.
    In the user's file this is where GroupA / Group B dropdown-like structure exists.
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
                try:
                    stock_w = float(vals[92]) if vals[92] is not None else np.nan
                    inflow_w = float(vals[93]) if vals[93] is not None else np.nan
                except Exception:
                    stock_w, inflow_w = np.nan, np.nan

                if d.year >= 2024 and group_a and group_b and not group_a.startswith("Group"):
                    rows.append([d, group_a, group_b, stock_w, inflow_w])

    df = pd.DataFrame(rows, columns=["date", "group_a", "group_b", "stock_weight", "inflow_weight"])
    if df.empty:
        return df
    df["date"] = pd.to_datetime(df["date"])
    df = df.dropna(subset=["inflow_weight"], how="all")
    return df

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
                    vals[4],
                ])
    return pd.DataFrame(rows, columns=["category", "task_type", "group_a", "group_b", "count"])

def apply_scope(total_df, weights_df, mode, group_a_choice=None, group_b_choice=None):
    """
    Uses daily percentages to convert the main daily series into:
    - Only faults
    - Selected GroupA
    - Selected GroupB
    Total mode uses unscaled totals.
    """
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

    # Convert totals to selected scope.
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
        score += 10
        reasons.append("βροχή")
    if rain >= rain_medium:
        score += 15
        reasons.append("έντονη βροχή")
    if rain >= rain_high:
        score += 25
        reasons.append("πολύ υψηλή βροχή")
    if wind >= wind_high:
        score += 15
        reasons.append("ισχυρός άνεμος")
    if prob >= 70:
        score += 10
        reasons.append("υψηλή πιθανότητα υετού")
    if code in [95, 96, 99]:
        score += 35
        reasons.append("καταιγίδα / πιθανή ηλεκτρική δραστηριότητα")

    if score >= 65:
        return "CRITICAL", score, ", ".join(sorted(set(reasons)))
    if score >= 40:
        return "HIGH", score, ", ".join(sorted(set(reasons)))
    if score >= 20:
        return "MEDIUM", score, ", ".join(sorted(set(reasons)))
    return "LOW", score, "χωρίς έντονο φαινόμενο"

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

        events.append({
            "Ημερομηνία κακοκαιρίας": bad_day,
            "Βροχή mm": round(float(bad_row.get("precipitation_sum", 0) or 0), 1),
            "Άνεμος km/h": round(float(bad_row.get("wind_speed_10m_max", 0) or 0), 1),
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

    if not bad.empty:
        high_uplift = max(0.0, bad["inflow_delta_pct"].median())
    else:
        high_uplift = 0.20

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

st.markdown('<p class="main-title">⛈️ Predictive Fault Weather Engine v2</p>', unsafe_allow_html=True)
st.markdown('<p class="subtitle">Μόνο βλάβες / dropdown επιλογές / πρόγνωση 7 ημερών / εισροή & stock recovery</p>', unsafe_allow_html=True)

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

    with st.spinner("Τραβάω ιστορικό καιρό / πρόγνωση 7 ημερών και υπολογίζω..."):
        weather = fetch_historical_weather(lat, lon, start_d, end_d)
        forecast = fetch_7day_forecast(lat, lon)

    merged = scoped.merge(weather, on="date", how="left")
    analyzed, events = analyze_history(merged, rain_medium, wind_high, stock_recovery_pct)
    prediction = build_forecast_prediction(forecast, analyzed, rain_medium, rain_high, wind_high)

    top = prediction.iloc[0]
    risk_class = {
        "LOW": "risk-low",
        "MEDIUM": "risk-medium",
        "HIGH": "risk-high",
        "CRITICAL": "risk-critical",
    }[top["Risk"]]

    st.markdown(f"""
    <div class="card {risk_class}">
        <div class="big">Τρέχουσα σήμανση: {top['Risk']}</div>
        <div>Scope: <span class="value">{scope_note}</span> | Περιοχή καιρού: <span class="value">{weather_area}</span></div>
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
    c4.metric("Μέση ολοκλήρωση", f"{analyzed['completed'].median():.0f}")

    st.subheader("🔔 Πρόγνωση 7 ημερών")
    st.dataframe(prediction, use_container_width=True)

    st.subheader("📊 Ιστορική σύγκριση κακοκαιρίας → εισροής / stock")
    if events.empty:
        st.warning("Δεν βρέθηκαν ημέρες κακοκαιρίας με τα thresholds που έχεις επιλέξει.")
    else:
        st.dataframe(events, use_container_width=True)

    st.subheader("📈 Γράφημα scoped εισροής / ολοκληρώσεων / stock")
    st.line_chart(analyzed.set_index("date")[["inflow", "completed", "stock"]])

    with st.expander("🔎 Έλεγχος φίλτρων από το αρχείο"):
        st.write("GroupA που βρέθηκαν:", group_as)
        st.write("GroupB που βρέθηκαν:", group_bs)
        if not clean_tasks.empty:
            st.dataframe(clean_tasks, use_container_width=True)

    st.download_button(
        "Κατέβασε πλήρη ανάλυση CSV",
        analyzed.to_csv(index=False).encode("utf-8-sig"),
        file_name="fault_weather_analysis_v2.csv",
        mime="text/csv"
    )

except Exception as e:
    st.error(f"Σφάλμα: {e}")
    st.caption("Έλεγξε ότι το αρχείο είναι το σωστό Chart.v4.xlsb και ότι υπάρχει πρόσβαση internet για Open-Meteo.")
