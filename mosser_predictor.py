#!/usr/bin/env python3
"""
Mosser Predictor 2026
Adam Central Football – Red & Gray
"""

import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import numpy as np
from pathlib import Path
from PIL import Image

st.set_page_config(
    page_title="Mosser Predictor 2026",
    page_icon="🏈",
    layout="centered",
    initial_sidebar_state="collapsed"
)

# ------------------------------------------------------------------
# CSS – Red & Gray + bigger prediction text
# ------------------------------------------------------------------
st.markdown("""
<style>
    .stApp { background-color: #1e1e1e; color: #e8e8e8; }
    h1 {
        color: #c8102e !important;
        text-align: center;
        font-family: 'Arial Black', Arial, sans-serif;
        font-size: 2.6rem !important;
        letter-spacing: 1px;
        text-shadow: 2px 2px 4px #000;
        margin-bottom: 0.2rem;
    }
    .subtitle {
        text-align: center;
        color: #999;
        font-size: 1.05rem;
        margin-bottom: 1.2rem;
    }
    label { color: #d0d0d0 !important; font-weight: 600 !important; }

    div.stButton > button {
        background-color: #c8102e;
        color: white !important;
        font-weight: 800;
        font-size: 1.35rem;
        border: 3px solid #8b0000;
        border-radius: 10px;
        padding: 0.75rem 2rem;
        width: 100%;
        letter-spacing: 1px;
    }
    div.stButton > button:hover {
        background-color: #a00d25;
        border-color: #c8102e;
    }

    /* GOT YA box */
    .got-ya {
        background: linear-gradient(145deg, #2a0000, #1a1a1a);
        border: 4px solid #c8102e;
        border-radius: 16px;
        padding: 1.8rem 1.2rem;
        text-align: center;
        margin: 1.2rem 0;
        box-shadow: 0 0 25px rgba(200,16,46,0.45);
    }
    .got-ya .phrase {
        color: #ffcc00;
        font-size: 4.2rem;
        font-weight: 900;
        letter-spacing: 3px;
        text-transform: uppercase;
        margin-bottom: 1.2rem;
        text-shadow: 0 0 18px #c8102e;
        animation: flashGotYa 0.7s ease-in-out infinite alternate;
    }
    @keyframes flashGotYa {
        from {
            opacity: 1;
            text-shadow: 0 0 8px #c8102e, 0 0 16px #ffcc00;
            transform: scale(1);
        }
        to {
            opacity: 0.55;
            text-shadow: 0 0 20px #ff0000, 0 0 35px #ffcc00;
            transform: scale(1.06);
        }
    }
    .got-ya .pred {
        color: #ffffff;
        font-size: 3.4rem;
        font-weight: 900;
        margin: 0.6rem 0;
        line-height: 1.15;
    }
    .got-ya .detail {
        color: #bbbbbb;
        font-size: 1rem;
    }

    /* Strong trend box */
    .trend-box {
        background-color: #252525;
        border: 3px solid #888;
        border-radius: 14px;
        padding: 1.5rem 1.2rem;
        text-align: center;
        margin: 1rem 0;
    }
    .trend-box .pred {
        color: #c8102e;
        font-size: 1.9rem;
        font-weight: 800;
        margin: 0.35rem 0;
    }
    .trend-box .detail {
        color: #aaa;
        font-size: 0.95rem;
    }

    /* Not today box */
    .not-today {
        background-color: #222;
        border: 3px solid #555;
        border-radius: 14px;
        padding: 1.6rem 1rem;
        text-align: center;
        margin: 1rem 0;
    }
    .not-today .phrase {
        color: #999;
        font-size: 2.0rem;
        font-weight: 800;
        letter-spacing: 2px;
        text-transform: uppercase;
        margin-bottom: 0.8rem;
    }
    .not-today .baseline {
        color: #ddd;
        font-size: 1.4rem;
        font-weight: 600;
    }

    .stSelectbox > div > div { background-color: #333; color: #eee; }
    hr { border-color: #444; }
    .footer {
        text-align: center;
        color: #666;
        font-size: 0.85rem;
        margin-top: 1.5rem;
    }
    .filters {
        text-align: center;
        color: #888;
        font-size: 0.9rem;
        margin-bottom: 0.5rem;
    }
</style>
""", unsafe_allow_html=True)

# ------------------------------------------------------------------
# Data
# ------------------------------------------------------------------
@st.cache_data
def load_data():
    path = Path(__file__).parent / "AC_Offense_2025_Master.xlsx"
    # Try named sheet first, then fall back to first sheet
    try:
        df = pd.read_excel(path, sheet_name="AC Offense 2025", engine="openpyxl")
    except Exception:
        df = pd.read_excel(path, engine="openpyxl")
    df["DN"] = pd.to_numeric(df["DN"], errors="coerce")
    df["DIST"] = pd.to_numeric(df["DIST"], errors="coerce")
    df["YARD LN"] = pd.to_numeric(df["YARD LN"], errors="coerce")

    for col in ["HASH", "OFF FORM", "PLAY DIR", "OFF PLAY"]:
        if col in df.columns:
            df[col] = (df[col].astype(str).str.strip().str.upper()
                       .replace({"NAN": np.nan, "NONE": "NONE"}))
    df["PLAY TYPE"] = (df["PLAY TYPE"].astype(str).str.strip()
                       .str.title().replace({"Nan": np.nan}))

    # Ensure new columns exist
    if "YARD ZONE" not in df.columns or "DIST GROUP" not in df.columns:
        def yz(y):
            if pd.isna(y): return np.nan
            y = float(y)
            if y <= -40: return "Deep Own (≤ -40)"
            if y <= -20: return "Own 40–20"
            if y <= -10: return "Own 20–10"
            if y < 0: return "Own Red Zone (-10 to GL)"
            if y <= 5: return "Opp Goal Line (1–5)"
            if y <= 20: return "Opp 20–5"
            if y <= 40: return "Opp 40–20"
            return "Opp 50–40 / Midfield"
        def dg(d):
            if pd.isna(d): return np.nan
            d = float(d)
            if d >= 11: return "10+"
            if d >= 7: return "10–7"
            if d >= 5: return "7–5"
            if d >= 3: return "5–3"
            if d == 2: return "3–2"
            if d == 1: return "1"
            return "Other"
        df["YARD ZONE"] = df["YARD LN"].apply(yz)
        df["DIST GROUP"] = df["DIST"].apply(dg)

    return df

DF = load_data()

_base = DF["PLAY TYPE"].value_counts(normalize=True)
BASE_RUN = round(float(_base.get("Run", 0)) * 100, 1)
BASE_PASS = round(float(_base.get("Pass", 0)) * 100, 1)

MIN_SAMPLE = 5
THRESHOLD_STRONG = 75.0
THRESHOLD_GOTYA = 90.0

# ------------------------------------------------------------------
# Helpers
# ------------------------------------------------------------------
def filter_plays(dn=None, dist_group=None, hash_=None, yard_zone=None, form=None):
    mask = pd.Series([True] * len(DF))
    if dn is not None:
        mask &= (DF["DN"] == float(dn))
    if dist_group is not None:
        mask &= (DF["DIST GROUP"] == dist_group)
    if hash_ is not None:
        mask &= (DF["HASH"] == str(hash_).upper())
    if yard_zone is not None:
        mask &= (DF["YARD ZONE"] == yard_zone)
    if form is not None:
        mask &= (DF["OFF FORM"] == str(form).upper())
    return DF[mask].copy()

def pct_series(series):
    counts = series.dropna().value_counts()
    total = int(counts.sum())
    if total == 0:
        return []
    return [(str(val), float(round(cnt / total * 100, 1)), int(cnt))
            for val, cnt in counts.items()]

def predict(dn=None, dist_group=None, hash_=None, yard_zone=None, form=None):
    sub = filter_plays(dn=dn, dist_group=dist_group, hash_=hash_,
                       yard_zone=yard_zone, form=form)
    n = len(sub)

    filters_used = {k: v for k, v in {
        "Down": dn,
        "Distance": dist_group,
        "Hash": hash_,
        "Field Zone": yard_zone,
        "Formation": form
    }.items() if v is not None}

    result = {
        "sample_size": n,
        "filters_used": filters_used,
        "trends": [],
        "got_ya": False,
        "fallback": False,
        "run_pct": BASE_RUN,
        "pass_pct": BASE_PASS,
    }

    if n < MIN_SAMPLE:
        result["fallback"] = True
        return result

    # Play Type
    pt = pct_series(sub["PLAY TYPE"])
    if pt:
        result["run_pct"] = next((p for v, p, c in pt if v == "Run"), BASE_RUN)
        result["pass_pct"] = next((p for v, p, c in pt if v == "Pass"), BASE_PASS)
        top = pt[0]
        if top[1] >= THRESHOLD_STRONG:
            result["trends"].append((top[0], top[1], top[2]))

    # Specific Play
    plays = pct_series(sub["OFF PLAY"])
    if plays:
        top = plays[0]
        if top[1] >= THRESHOLD_STRONG:
            result["trends"].append((f"{top[0]}", top[1], top[2]))

    # Direction
    dirs = pct_series(sub["PLAY DIR"])
    if dirs:
        top = dirs[0]
        if top[1] >= THRESHOLD_STRONG:
            result["trends"].append((f"Direction {top[0]}", top[1], top[2]))

    if result["trends"]:
        result["got_ya"] = max(t[1] for t in result["trends"]) >= THRESHOLD_GOTYA
        result["fallback"] = False
    else:
        result["fallback"] = True

    return result


def show_autoplay_video(path, max_height=280):
    """Embed a local mp4 that tries to autoplay with sound."""
    import base64
    path = Path(path)
    if not path.exists():
        st.caption("Video file not found")
        return
    with open(path, "rb") as f:
        data = base64.b64encode(f.read()).decode()
    html = f"""
    <video id="predvid" width="100%" controls autoplay playsinline
           style="border-radius:10px; max-height:{max_height}px;">
        <source src="data:video/mp4;base64,{data}" type="video/mp4">
    </video>
    <script>
        const v = document.getElementById('predvid');
        if (v) {{
            v.muted = false;
            v.volume = 1.0;
            const p = v.play();
            if (p !== undefined) {{
                p.catch(() => {{
                    // Autoplay with sound blocked – try muted then unmute on first click
                    v.muted = true;
                    v.play().then(() => {{
                        document.body.addEventListener('click', () => {{
                            v.muted = false;
                        }}, {{ once: true }});
                    }}).catch(()=>{{}});
                }});
            }}
        }}
    </script>
    """
    components.html(html, height=max_height + 40)


# ------------------------------------------------------------------
# UI
# ------------------------------------------------------------------
st.markdown("<h1>MOSSER PREDICTOR 2026</h1>", unsafe_allow_html=True)
st.markdown('<p class="subtitle">Adam Central Football • 2025 Season</p>', unsafe_allow_html=True)

col1, col2 = st.columns(2)

with col1:
    yard_zones = [
        "—",
        "Deep Own (≤ -40)",
        "Own 40–20",
        "Own 20–10",
        "Own Red Zone (-10 to GL)",
        "Opp Goal Line (1–5)",
        "Opp 20–5",
        "Opp 40–20",
        "Opp 50–40 / Midfield",
    ]
    yard_sel = st.selectbox("Field Zone", yard_zones, index=0)

    hash_sel = st.selectbox("Hash", ["—", "L", "M", "R"], index=0)

    quarter_sel = st.selectbox("Quarter", ["—", "1", "2", "3", "4"], index=0,
                               help="Not in data yet – ignored for now")

with col2:
    down_sel = st.selectbox("Down", ["—", "1", "2", "3", "4"], index=0)

    dist_groups = ["—", "10+", "10–7", "7–5", "5–3", "3–2", "1"]
    dist_sel = st.selectbox("Distance", dist_groups, index=0)

    form_list = sorted([f for f in DF["OFF FORM"].dropna().unique().tolist() if f])
    form_sel = st.selectbox("Formation", ["—"] + form_list, index=0)

st.markdown("")
predict_btn = st.button("PREDICT", use_container_width=True)

# ------------------------------------------------------------------
# Results
# ------------------------------------------------------------------
if predict_btn:
    dn = int(down_sel) if down_sel != "—" else None
    dist_group = dist_sel if dist_sel != "—" else None
    hash_ = hash_sel if hash_sel != "—" else None
    yard_zone = yard_sel if yard_sel != "—" else None
    form = form_sel if form_sel != "—" else None

    result = predict(dn=dn, dist_group=dist_group, hash_=hash_,
                     yard_zone=yard_zone, form=form)

    st.markdown("---")

    if result["filters_used"]:
        used = "  •  ".join(f"{k}: {v}" for k, v in result["filters_used"].items())
        st.markdown(f'<p class="filters">Using → {used}<br>Sample: {result["sample_size"]} plays</p>',
                    unsafe_allow_html=True)
    else:
        st.markdown(f'<p class="filters">No filters • Overall baseline • {result["sample_size"]} plays</p>',
                    unsafe_allow_html=True)

    # ========== GOT YA (90%+) ==========
    if result["got_ya"]:
        st.markdown('<div class="got-ya">', unsafe_allow_html=True)

        left, right = st.columns([1.3, 1])

        with left:
            st.markdown('<p class="phrase">GOT YA FUCKER</p>', unsafe_allow_html=True)
            for label, pct, cnt in result["trends"]:
                st.markdown(f'<p class="pred">{pct}%  {label}</p>', unsafe_allow_html=True)
                st.markdown(f'<p class="detail">{cnt} of {result["sample_size"]} plays</p>',
                            unsafe_allow_html=True)

        with right:
            folder = Path(__file__).parent
            video_path = folder / "got_ya.mp4"
            if not video_path.exists():
                for p in folder.glob("*.mp4"):
                    if "not_today" not in p.name.lower():
                        video_path = p
                        break
            show_autoplay_video(video_path, max_height=300)

        st.markdown('</div>', unsafe_allow_html=True)

    # ========== Strong 75-89% ==========
    elif result["trends"]:
        st.markdown('<div class="trend-box">', unsafe_allow_html=True)
        st.markdown("### STRONG TREND")
        for label, pct, cnt in result["trends"]:
            st.markdown(f'<p class="pred">{pct}%  {label}</p>', unsafe_allow_html=True)
            st.markdown(f'<p class="detail">{cnt} of {result["sample_size"]} plays</p>',
                        unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)

    # ========== Not Today ==========
    else:
        st.markdown('<div class="not-today">', unsafe_allow_html=True)

        left, right = st.columns([1.2, 1])

        with left:
            st.markdown('<p class="phrase">NOT TODAY</p>', unsafe_allow_html=True)
            st.markdown(
                f'<p class="baseline">Run {result["run_pct"]}% &nbsp;&nbsp;|&nbsp;&nbsp; Pass {result["pass_pct"]}%</p>',
                unsafe_allow_html=True
            )
            if result["sample_size"] < MIN_SAMPLE and result["sample_size"] > 0:
                st.caption(f"Only {result['sample_size']} matching plays – showing overall baseline")

        with right:
            folder = Path(__file__).parent
            video_path = folder / "not_today.mp4"
            if not video_path.exists():
                for p in folder.glob("*.mp4"):
                    if "got_ya" not in p.name.lower():
                        video_path = p
                        break
            show_autoplay_video(video_path, max_height=260)

        st.markdown('</div>', unsafe_allow_html=True)

st.markdown("---")
st.markdown(
    '<div class="footer">Mosser Predictor 2026 &nbsp;•&nbsp; 75%+ Trend &nbsp;•&nbsp; 90%+ Got Ya</div>',
    unsafe_allow_html=True
)
