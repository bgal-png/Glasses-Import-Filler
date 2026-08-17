# -*- coding: utf-8 -*-
"""
Standalone Barcode Checker — a minimal, fool-proof Streamlit app.

Colleagues upload a file with a Barcode column; it reports how many of those
barcodes exist in the product database (master_catalog) and which don't.

Deploy on Streamlit Cloud as its OWN app: point the app's "Main file path"
at `barcode_checker.py` and set DB_URL in that app's Secrets (same value as
the admin app).
"""
import io
import re

import pandas as pd
import streamlit as st
from sqlalchemy import create_engine

st.set_page_config(page_title="Barcode Checker", layout="centered")
st.title("🔍 Barcode Checker")
st.write(
    "Upload your file with the **Barcode** column filled in. "
    "This tool tells you how many of those barcodes are already in the database "
    "and which ones are not."
)

# ---------------------------------------------------------------------------
if "DB_URL" not in st.secrets:
    st.error(
        "⚙️ This app isn't connected to the database yet.\n\n"
        "Add a secret named **DB_URL** in this app's settings "
        "(Manage app → Settings → Secrets) with the same value as the admin app, "
        "then reboot."
    )
    st.stop()

DB_URL = st.secrets["DB_URL"]


@st.cache_resource
def _engine():
    return create_engine(DB_URL, pool_pre_ping=True, pool_recycle=300)


# (sidebar panel rendered near the end, after helpers are defined)


@st.cache_data(ttl=600, show_spinner=False)
def _load_known():
    """Return a dict: clean_barcode -> a short label (name) for display.
    Loads only the two columns needed and builds the map vectorized (fast even
    on 150k+ rows)."""
    eng = _engine()
    try:
        df = pd.read_sql('SELECT join_key, "Assembled_Name" FROM master_catalog', con=eng)
    except Exception:
        # Fallback: table exists but without Assembled_Name
        try:
            df = pd.read_sql('SELECT join_key FROM master_catalog', con=eng)
            df["Assembled_Name"] = ""
        except Exception as e:
            raise RuntimeError(f"Could not read the database: {e}")
    df["join_key"] = df["join_key"].astype(str).str.strip()
    df = df[(df["join_key"] != "") & (df["join_key"].str.lower() != "nan")]
    return dict(zip(df["join_key"], df["Assembled_Name"].astype(str).fillna("")))


@st.cache_data(ttl=3600, show_spinner=False)
def _load_ingest_log():
    """Cached read of the tiny ingest_log table (producer -> last-updated)."""
    try:
        log = pd.read_sql_table("ingest_log", con=_engine())
        return {str(r["manufacturer"]).lower(): str(r["last_updated"]) for _, r in log.iterrows()}
    except Exception:
        return {}


def _render_last_update_sidebar():
    from datetime import datetime
    try:
        from dictionaries import MANUFACTURER_CONFIG
        producers = list(MANUFACTURER_CONFIG.keys())
    except Exception:
        producers = ["safilo", "luxottica", "marcolin", "kering", "derigo", "thelios"]

    log_map = _load_ingest_log()
    for extra in log_map:
        if extra not in [p.lower() for p in producers]:
            producers.append(extra)

    def _fmt(iso):
        if not iso or iso.lower() == "none":
            return "never"
        try:
            dt = datetime.fromisoformat(iso.replace("Z", ""))
            return f"{dt.day}.{dt.month}.{dt.year}"
        except Exception:
            return iso

    with st.sidebar:
        st.markdown("### 🗓️ Last catalogue update")
        rows = [{"Producer": p.title(), "Last update": _fmt(log_map.get(p.lower()))} for p in producers]
        st.dataframe(pd.DataFrame(rows), hide_index=True, use_container_width=True)
        st.caption("When each producer's catalogue was last processed.")


def _clean_bc(x):
    """Normalize a barcode the same way the database stores it (join_key)."""
    return re.sub(r"\.0$", "", str(x).strip()).lstrip("0")


def _read_any(uploaded):
    """Read an uploaded xlsx/csv into a string DataFrame with stripped headers."""
    name = uploaded.name.lower()
    if name.endswith(".csv"):
        try:
            d = pd.read_csv(uploaded, dtype=str, sep=",", on_bad_lines="skip")
            if len(d.columns) <= 1:
                uploaded.seek(0)
                d = pd.read_csv(uploaded, dtype=str, sep=";", on_bad_lines="skip")
        except Exception:
            uploaded.seek(0)
            d = pd.read_csv(uploaded, dtype=str, sep=";", on_bad_lines="skip")
    else:
        d = pd.read_excel(uploaded, dtype=str, engine="openpyxl")
    d.columns = (
        d.columns.astype(str)
        .str.replace(r"[\r\n\t]", " ", regex=True)
        .str.replace(r"\s+", " ", regex=True)
        .str.strip()
    )
    return d


def _find_barcode_col(df):
    lower = {c.lower(): c for c in df.columns}
    for cand in ("barcode", "ean", "upc", "ean/upc", "ean code", "* ean code"):
        if cand in lower:
            return lower[cand]
    # fall back: any column containing 'barcode'/'ean'/'upc'
    for c in df.columns:
        if any(k in c.lower() for k in ("barcode", "ean", "upc")):
            return c
    return None


# ---------------------------------------------------------------------------
_render_last_update_sidebar()

uploaded = st.file_uploader("Upload your file (.xlsx or .csv)", type=["xlsx", "csv"])

if uploaded is not None:
    try:
        df = _read_any(uploaded)
    except Exception as e:
        st.error(f"❌ Could not read the file: {e}")
        st.stop()

    bc_col = _find_barcode_col(df)
    if bc_col is None:
        st.error(
            "❌ Couldn't find a **Barcode** column in your file. "
            f"Columns found: {', '.join(map(str, df.columns))}"
        )
        st.stop()

    if bc_col.lower() != "barcode":
        st.info(f"Using column **{bc_col}** as the barcode column.")

    try:
        with st.spinner("Checking against the database…"):
            known = _load_known()
    except Exception as e:
        st.error(f"❌ {e}")
        st.stop()

    # Build results
    rows = []
    seen = set()
    for raw in df[bc_col]:
        raw_s = str(raw).strip()
        if not raw_s or raw_s.lower() == "nan":
            continue
        key = _clean_bc(raw_s)
        in_db = key in known
        rows.append({
            "Barcode": raw_s,
            "Status": "✅ In database" if in_db else "❌ Not in database",
            "Product": known.get(key, "") if in_db else "",
        })
        seen.add(key)

    if not rows:
        st.warning("No barcodes found in the file.")
        st.stop()

    res = pd.DataFrame(rows)
    n_total = len(res)
    n_in = int((res["Status"] == "✅ In database").sum())
    n_out = n_total - n_in

    st.divider()
    c1, c2, c3 = st.columns(3)
    c1.metric("Checked", f"{n_total}")
    c2.metric("✅ In database", f"{n_in}")
    c3.metric("❌ Not in database", f"{n_out}")

    st.divider()
    tab_all, tab_missing = st.tabs(["📋 All results", f"❌ Not in database ({n_out})"])
    with tab_all:
        st.dataframe(res, use_container_width=True, hide_index=True)
    with tab_missing:
        missing = res[res["Status"] == "❌ Not in database"][["Barcode"]]
        if missing.empty:
            st.success("🎉 Every barcode is in the database.")
        else:
            st.dataframe(missing, use_container_width=True, hide_index=True)

    st.download_button(
        "📥 Download full result (CSV)",
        data=res.to_csv(index=False).encode("utf-8-sig"),
        file_name="barcode_check_result.csv",
        mime="text/csv",
    )
else:
    st.info("Waiting for a file…")
