
import streamlit as st
import pandas as pd
import numpy as np
import requests
from pyxlsb import open_workbook
from datetime import datetime, timedelta
from pathlib import Path

st.set_page_config(
    page_title="Fault Weather Alert",
    page_icon="⛈️",
    layout="wide",
    initial_sidebar_state="collapsed"
)

ATHENS_LAT = 37.9842
ATHENS_LON = 23.7281

st.markdown("""
<style>
.block-container {padding-top: 1rem; padding-bottom: 1rem; max-width: 1100px;}
.main-title {font-size: 2.1rem; font-weight: 900; margin-bottom: 0;}
.subtitle {color: #6b7280; margin-bottom: 1rem;}
.card {
    border-radius: 22px;
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
.big {font-size: 1.35rem; font-weight: 900;}
.label {color:#6b7280;font-size:.92rem;}
.value {font-weight:800;}
.alert-text {font-size: 1rem; line-height: 1.45;}
</style>
""", unsafe_allow_html=True)

def excel_serial_to_date(x):
    return (datetime(1899, 12, 30) + timedelta(days=int(x))).date()

def extract_daily_series_from_xlsb(file_path):
    rows = []
    with open_workbook(file_path) as wb:
        if "Table" not in wb.sheets:
            return pd.DataFrame()
        sh = wb.get_sheet("Table")
        for row in sh.rows():
            vals = [c.v for c in row]
            # Expected columns H:L:
            # H date, I completed, J stock, K inflow, L not_completed
            if len(vals) >= 12 and isinstance(vals[7], (int, float)):
                try:
                    d = excel_serial_to_date(vals[7])
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

def classify_weather(row):
    rain = float(row.get("precipitation_sum", 0) or 0)
    wind = float(row.get("wind_speed_10m_max", 0) or 0)
    prob = float(row.get("precipitation_probability_max", 0) or 0)
    code = int(row.get("weather_code", 0) or 0)

    score = 0
    reasons = []

    if rain >= 5:
        score += 10; reasons.append("βροχή")
    if rain >= 10:
        score += 15; reasons.append("έντονη βροχή")
    if rain >= 20:
        score += 25; reasons.append("πολύ υψηλή βροχή")
    if wind >= 40:
        score += 10; reasons.append("ισχυρός άνεμος")
    if wind >= 55:
        score += 15; reasons.append("πολύ ισχυρός άνεμος")
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

def analyze_history(df):
    df = df.copy()
    df["baseline_inflow"] = df.groupby("weekday")["inflow"].transform("median")
    df["inflow_delta"] = df["inflow"] - df["baseline_inflow"]
    df["inflow_delta_pct"] = np.where(df["baseline_inflow"] > 0, df["inflow_delta"] / df["baseline_inflow"], 0)
    df["stock_prev"] = df["stock"].shift(1)
    df["stock_delta"] = df["stock"] - df["stock_prev"]

    df["bad_weather"] = (
        (df["precipitation_sum"].fillna(0) >= 10) |
        (df["wind_speed_10m_max"].fillna(0) >= 45) |
        (df["weather_code"].isin([95, 96, 99]))
    )

    median_inflow = df["inflow"].median()
    median_stock = df["stock"].median()
    normal_inflow_threshold = median_inflow * 1.10
    normal_stock_threshold = median_stock * 1.05

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

def build_forecast_prediction(forecast, history):
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
        risk, score, reasons = classify_weather(row)

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

st.markdown('<p class="main-title">⛈️ Fault Weather Alert — PC App</p>', unsafe_allow_html=True)
st.markdown('<p class="subtitle">Πρόβλεψη εισροής βλαβών / stock / απορρόφησης βάσει καιρού</p>', unsafe_allow_html=True)

with st.sidebar:
    st.header("Ρυθμίσεις")
    st.caption("Για Αττική άφησε τις default συντεταγμένες.")
    lat = st.number_input("Latitude", value=ATHENS_LAT, format="%.4f")
    lon = st.number_input("Longitude", value=ATHENS_LON, format="%.4f")
    uploaded = st.file_uploader("Ανέβασε Chart.v4.xlsb", type=["xlsb"])

st.markdown("""
<div class="card">
<div class="big">Τι κάνει η εφαρμογή</div>
<p class="alert-text">
Ανεβάζεις το αρχείο <b>Chart.v4.xlsb</b>. Η εφαρμογή διαβάζει ιστορική εισροή/ολοκληρώσεις/stock,
τραβάει ιστορικό καιρό και πρόγνωση 7 ημερών, και βγάζει operational σήμανση.
</p>
</div>
""", unsafe_allow_html=True)

if not uploaded:
    st.info("Ανέβασε το Chart.v4.xlsb από το πλαϊνό μενού για να ξεκινήσει η ανάλυση.")
    st.stop()

temp_path = "uploaded_chart.xlsb"
with open(temp_path, "wb") as f:
    f.write(uploaded.getbuffer())

try:
    ops = extract_daily_series_from_xlsb(temp_path)
    if ops.empty:
        st.error("Δεν βρέθηκε το daily operational series στο sheet Table.")
        st.stop()

    start_d = ops["date"].min().date()
    end_d = ops["date"].max().date()

    with st.spinner("Διαβάζω ιστορικό καιρό και πρόγνωση 7 ημερών..."):
        weather = fetch_historical_weather(lat, lon, start_d, end_d)
        forecast = fetch_7day_forecast(lat, lon)

    merged = ops.merge(weather, on="date", how="left")
    analyzed, events = analyze_history(merged)
    prediction = build_forecast_prediction(forecast, analyzed)

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
        <p class="alert-text">
        Αιτία: <b>{top['Αιτία']}</b><br>
        Εκτίμηση εισροής: <b>{top['Εκτίμηση εισροής']:.0f}</b>
        ({top['Αύξηση vs normal %']:.1f}% vs normal)<br>
        Projected stock: <b>{top['Projected Stock']:.0f}</b><br>
        Εκτίμηση απορρόφησης: <b>{top['Εκτίμηση απορρόφησης ημέρες']} ημέρες</b>
        </p>
    </div>
    """, unsafe_allow_html=True)

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Ιστορικές ημέρες", len(analyzed))
    c2.metric("Συνήθης εισροή", f"{analyzed['inflow'].median():.0f}")
    c3.metric("Συνήθες stock", f"{analyzed['stock'].median():.0f}")
    c4.metric("Μέση ολοκλήρωση", f"{analyzed['completed'].median():.0f}")

    st.subheader("🔔 Πρόγνωση 7 ημερών")
    st.dataframe(prediction, use_container_width=True)

    st.subheader("📊 Ιστορική σύγκριση κακοκαιρίας → εισροής / stock")
    if events.empty:
        st.warning("Δεν βρέθηκαν ημέρες κακοκαιρίας με τα thresholds της v1.")
    else:
        st.dataframe(events, use_container_width=True)

    st.subheader("📈 Γράφημα εισροής / ολοκληρώσεων / stock")
    st.line_chart(analyzed.set_index("date")[["inflow", "completed", "stock"]])

    st.download_button(
        "Κατέβασε πλήρη ανάλυση CSV",
        analyzed.to_csv(index=False).encode("utf-8-sig"),
        file_name="fault_weather_full_analysis.csv",
        mime="text/csv"
    )

except Exception as e:
    st.error(f"Σφάλμα: {e}")
    st.caption("Έλεγξε ότι έχεις internet και ότι το αρχείο είναι το σωστό Chart.v4.xlsb.")
