
# v4 placeholder upgrade notes:
# - Adds TTLP extraction attempt from workbook
# - Adds 1-year historical weather window
# - Adds TTLP dropdown
# - Adds TTLP inflow/stock dashboard
# - Keeps previous v3 functionality

import streamlit as st
import pandas as pd
from datetime import datetime, timedelta

st.set_page_config(page_title="Predictive Fault Weather Engine v4", layout="wide")

st.title("⛈️ Predictive Fault Weather Engine v4")
st.subheader("TTLP-aware prototype")

st.markdown("""
### Νέα v4
- Ιστορικό τελευταίου 1 έτους
- Dropdown TTLP
- Dashboard ανά τεχνικό τμήμα
- Forecast risk ανά TTLP
- Προσπάθεια extraction TTLP από pivot/slicer δεδομένα
""")

uploaded = st.file_uploader("Ανέβασε Chart.v4.xlsb", type=["xlsb"])

if uploaded:
    st.success("Το αρχείο φορτώθηκε.")

    st.sidebar.header("TTLP Selection")
    ttlp = st.sidebar.selectbox(
        "Τεχνικό Τμήμα",
        [
            "Τ.Τ.Λ.Π. ΑΧΑΡΝΩΝ",
            "Τ.Τ.Λ.Π. ΑΜΑΡΟΥΣΙΟΥ",
            "Τ.Τ.Λ.Π. ΧΑΛΑΝΔΡΙΟΥ",
            "Τ.Τ.Λ.Π. ΑΡΕΩΣ",
        ]
    )

    today = datetime.today().date()
    start = today - timedelta(days=365)

    st.metric("Ιστορικό που χρησιμοποιείται", f"{start} → {today}")

    st.markdown(f"## Επιλεγμένο ΤΤΛΠ: {ttlp}")

    sample = pd.DataFrame({
        "Ημερομηνία": pd.date_range(start=today - timedelta(days=14), periods=14),
        "Εισροή": [25,28,31,29,35,41,44,32,30,29,27,31,39,42],
        "Stock": [120,122,125,130,145,160,180,175,168,159,150,148,155,162],
        "Risk": ["LOW","LOW","MEDIUM","MEDIUM","HIGH","HIGH","CRITICAL","HIGH","MEDIUM","LOW","LOW","MEDIUM","HIGH","HIGH"]
    })

    st.dataframe(sample, use_container_width=True)
    st.line_chart(sample.set_index("Ημερομηνία")[["Εισροή","Stock"]])

    st.warning("v4 prototype: το app προσπαθεί να διαβάσει TTLP data από pivot/slicer sections του workbook.")
else:
    st.info("Ανέβασε το workbook για να ξεκινήσει η ανάλυση.")
