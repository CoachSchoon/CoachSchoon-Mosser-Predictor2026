#!/usr/bin/env python3
"""
Mosser Predictor 2026
Adam Central Football – Red & Gray
"""

import streamlit as st
import pandas as pd
import numpy as np
from pathlib import Path

st.set_page_config(
    page_title="Mosser Predictor 2026",
    page_icon="🏈",
    layout="centered",
    initial_sidebar_state="collapsed"
)

st.markdown("""
<style>
    .stApp { background-color: #1e1e1e; color: #e8e8e8; }
    h1 {
        color: #c8102e !important;
        text-align: center;
        font-family: 'Arial Black', Arial, sans-serif;
        font-size: 2.3rem !important;
        letter-spacing: 1px;
        text-shadow: 2px 2px 4px #000;
        margin-bottom: 0.1rem;
    }
    .subtitle {
        text-align: center; color: #999; font-size: 0.95rem; margin-bottom: 0.8rem;
    }
    label { color: #d0d0d0 !important; font-weight: 600 !important; }

    div.stButton > button {
        background-color: #c8102e;
        color: white !important;
        font-weight: 800;
        font-size: 1.2rem;
        border: 3px solid #8b0000;
        border-radius: 10px;
        padding: 0.6rem 1.4rem;
        width: 100%;
    }
    div.stButton > button:hover {
        background-color: #a00d25;
        border-color: #c8102e;
    }

    /* Field zone buttons – green turf look */
    div[data-testid="stVerticalBlock"] button[kind="secondary"] {
        background-color: #2d5a27 !important;
        border: 2px solid #3d7a35 !important;
        color: #f0f0f0 !important;
        font-weight: 700 !important;
        font-size: 0.85rem !important;
        border-radius: 4px !important;
        min-height: 38px;
    }

    .got-ya {
        background: linear-gradient(145deg, #2a0000, #1a1a1a);
        border: 4px solid #c8102e;
        border-radius: 16px;
        padding: 1.4rem 1rem;
        text-align: center;
        margin: 1rem 0;
        box-shadow: 0 0 25px rgba(200,16,46,0.45);
    }
    .got-ya .phrase {
        color: #ffcc00;
        font-size: 3.4rem;
        font-weight: 900;
        letter-spacing: 3px;
        text-transform: uppercase;
        margin-bottom: 0.7rem;
        text-shadow: 0 0 18px #c8102e;
        animation: flashGotYa 0.7s ease-in-out infinite alternate;
    }
    @keyframes flashGotYa {
        from { opacity: 1; transform: scale(1); }
        to { opacity: 0.6; transform: scale(1.05); }
    }
    .got-ya .pred {
        color: #fff; font-size: 2.6rem; font-weight: 900; margin: 0.35rem 0; line-height: 1.15;
    }
    .got-ya .detail { color: #bbb; font-size: 0.9rem; }

    .trend-box {
        background-color: #252525;
        border: 3px solid #888;
        border-radius: 14px;
        padding: 1.2rem 1rem;
        text-align: center;
        margin: 1rem 0;
    }
    .trend-box .pred {
        color: #c8102e; font-size: 1.7rem; font-weight: 800; margin: 0.3rem 0;
    }
    .trend-box .detail { color: #aaa; font-size: 0.88rem; }

    .not-today {
        background-color: #222;
        border: 3px solid #555;
        border-radius: 14px;
        padding: 1.3rem 1rem;
        text-align: center;
        margin: 1rem 0;
    }
    .not-today .phrase {
        color: #999; font-size: 1.9rem; font-weight: 800;
        letter-spacing: 2px; text-transform: uppercase; margin-bottom: 0.5rem;
    }
    .not-today .baseline { color: #ddd; font-size: 1.25rem; font-weight: 600; }

    .filters { text-align: center; color: #888; font-size: 0.85rem; margin-bottom: 0.4rem; }
    .footer { text-align: center; color: #666; font-size: 0.8rem; margin-top: 1rem; }
    .field-title {
        text-align: center; color: #ccc; font-size: 0.9rem; font-weight: 700;
        margin: 0.4rem 0 0.3rem 0; letter-spacing: 1px;
    }
    .zone-selected {
        text-align: center; color: #ffcc00; font-weight: 700; font-size: 0.95rem;
        margin: 0.4rem 0 0.6rem 0;
    }
    .endzone {
        text-align: center; font-weight: 800; font-size: 0.8rem;
        padding: 0.25rem; border-radius: 4px; margin: 0.15rem 0;
    }
    .endzone-opp { background: #8b0000; color: #fff; }
    .endzone-own { background: #1a3a6b; color: #fff; }
</style>
""", unsafe_allow_html=True)

# ------------------------------------------------------------------
# Data
# ------------------------------------------------------------------
@st.cache_data
def load_data():
    path = Path(__file__).parent / "AC_Offense_2025_Master.xlsx"
    try:
        df = pd.read_excel(path, sheet_name="AC Offense 2025", engine="openpyxl")
    except Exception:
        df = pd.read_excel(path, engine="openpyxl")

    df["DN"] = pd.to_numeric(df["DN"], errors="coerce")
    df["DIST"] = pd.to_numeric(df["DIST"], errors="coerce")
    df["YARD LN"] = pd.to_numeric(df["YARD LN"], errors="coerce")

    for col in ["HASH", "OFF FORM", "PLAY DIR", "OFF PLAY", "MOTION", "MOTION DIR"]:
        if col in df.columns:
            df[col] = (df[col].astype(str).str.strip().str.upper()
                       .replace({"NAN": np.nan, "NONE": "NONE"}))
    df["PLAY TYPE"] = (df["PLAY TYPE"].astype(str).str.strip()
                       .str.title().replace({"Nan": np.nan}))

    if "YARD ZONE" not in df.columns:
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
        df["YARD ZONE"] = df["YARD LN"].apply(yz)

    if "DIST GROUP" not in df.columns:
        def dg(d):
            if pd.isna(d): return np.nan
            d = float(d)
            if d >= 11: return "10+"
            if d >= 7: return "10–7"
            if d >= 5: return "7–5"
            if d >= 3: return "5–3"
            if d <= 2: return "2–1"
            return "Other"
        df["DIST GROUP"] = df["DIST"].apply(dg)

    if "MOTION CAT" not in df.columns:
        def mc(row):
            m = str(row.get("MOTION", "")).upper().strip()
            d = str(row.get("MOTION DIR", "")).upper().strip()
            if m in ["NONE", "NAN", ""] or pd.isna(row.get("MOTION")):
                return "No Motion"
            if d == "R": return "Motion Right"
            if d == "L": return "Motion Left"
            return "No Motion"
        df["MOTION CAT"] = df.apply(mc, axis=1)

    if "PREV_PLAY_TYPE" not in df.columns:
        sort_cols = [c for c in ["SOURCE", "ID"] if c in df.columns]
        if sort_cols:
            df = df.sort_values(sort_cols).reset_index(drop=True)
        if "SOURCE" in df.columns:
            df["PREV_PLAY_TYPE"] = df.groupby("SOURCE")["PLAY TYPE"].shift(1)
        else:
            df["PREV_PLAY_TYPE"] = df["PLAY TYPE"].shift(1)
        df["PREV_PLAY_TYPE"] = df["PREV_PLAY_TYPE"].astype(str).str.title().replace({"Nan": np.nan})

    return df

DF_MASTER = load_data()

def _dist_group(d):
    try:
        d = float(d)
    except Exception:
        return None
    if d >= 11: return "10+"
    if d >= 7: return "10–7"
    if d >= 5: return "7–5"
    if d >= 3: return "5–3"
    if d <= 2: return "2–1"
    return "Other"

def _yard_zone(y):
    try:
        y = float(y)
    except Exception:
        return None
    if y <= -40: return "Deep Own (≤ -40)"
    if y <= -20: return "Own 40–20"
    if y <= -10: return "Own 20–10"
    if y < 0: return "Own Red Zone (-10 to GL)"
    if y <= 5: return "Opp Goal Line (1–5)"
    if y <= 20: return "Opp 20–5"
    if y <= 40: return "Opp 40–20"
    return "Opp 50–40 / Midfield"

def _motion_cat(motion, motion_dir):
    m = str(motion or "").upper().strip()
    d = str(motion_dir or "").upper().strip()
    if m in ["NONE", "NAN", ""] or m == "NONE":
        return "No Motion"
    if d == "R": return "Motion Right"
    if d == "L": return "Motion Left"
    return "No Motion"

def get_working_df():
    """Master data + any plays added live this session.
    If session flag live_only is on, return only live in-game plays.
    """
    if "live_plays" not in st.session_state:
        st.session_state.live_plays = []
    if "live_only" not in st.session_state:
        st.session_state.live_only = False

    live_rows = st.session_state.live_plays
    if st.session_state.live_only:
        if not live_rows:
            # empty frame with correct columns
            return DF_MASTER.iloc[0:0].copy()
        combined = pd.DataFrame(live_rows)
    elif not live_rows:
        return DF_MASTER.copy()
    else:
        live = pd.DataFrame(live_rows)
        combined = pd.concat([DF_MASTER, live], ignore_index=True)

    sort_cols = [c for c in ["SOURCE", "ID"] if c in combined.columns]
    if sort_cols:
        combined = combined.sort_values(sort_cols).reset_index(drop=True)
    if "SOURCE" in combined.columns and len(combined) > 0:
        combined["PREV_PLAY_TYPE"] = combined.groupby("SOURCE")["PLAY TYPE"].shift(1)
    elif len(combined) > 0:
        combined["PREV_PLAY_TYPE"] = combined["PLAY TYPE"].shift(1)
    if len(combined) > 0 and "PREV_PLAY_TYPE" in combined.columns:
        combined["PREV_PLAY_TYPE"] = (
            combined["PREV_PLAY_TYPE"].astype(str).str.title().replace({"Nan": np.nan})
        )
    return combined

DF = get_working_df()

_base = DF_MASTER["PLAY TYPE"].value_counts(normalize=True)
BASE_RUN = round(float(_base.get("Run", 0)) * 100, 1)
BASE_PASS = round(float(_base.get("Pass", 0)) * 100, 1)

MIN_SAMPLE = 5
THRESHOLD_STRONG = 75.0
THRESHOLD_GOTYA = 90.0

# Vertical field zones (TOP = opponent goal line / TD)
# Display order top → bottom
FIELD_BUTTONS = [
    ("OPP 5–GL",  "Opp Goal Line (1–5)"),
    ("OPP 20–5",  "Opp 20–5"),
    ("OPP 40–20", "Opp 40–20"),
    ("OPP 50–40", "Opp 50–40 / Midfield"),
    ("OWN 40–50", "Deep Own (≤ -40)"),
    ("OWN 20–40", "Own 40–20"),
    ("OWN 10–20", "Own 20–10"),
    ("OWN GL–10", "Own Red Zone (-10 to GL)"),
]

# ------------------------------------------------------------------
# Helpers
# ------------------------------------------------------------------
def filter_plays(dn=None, dist_group=None, hash_=None, yard_zone=None,
                 form=None, prev_play=None, allowed_forms=None):
    # Start from master + live plays, optionally restricted to active formations
    base = get_working_df()
    if allowed_forms is not None and len(allowed_forms) > 0:
        allowed = [str(f).upper() for f in allowed_forms]
        base = base[base["OFF FORM"].isin(allowed)]
    mask = pd.Series([True] * len(base), index=base.index)
    if dn is not None:
        mask &= (base["DN"] == float(dn))
    if dist_group is not None:
        mask &= (base["DIST GROUP"] == dist_group)
    if hash_ is not None:
        mask &= (base["HASH"] == str(hash_).upper())
    if yard_zone is not None:
        mask &= (base["YARD ZONE"] == yard_zone)
    if form is not None:
        mask &= (base["OFF FORM"] == str(form).upper())
    if prev_play is not None:
        mask &= (base["PREV_PLAY_TYPE"] == str(prev_play).title())
    return base[mask].copy()

def pct_series(series):
    counts = series.dropna().value_counts()
    total = int(counts.sum())
    if total == 0:
        return []
    return [(str(val), float(round(cnt / total * 100, 1)), int(cnt))
            for val, cnt in counts.items()]

def _build_trends(sub):
    """Extract strong trends from a filtered subset."""
    trends = []
    run_pct, pass_pct = BASE_RUN, BASE_PASS
    pass_dir = None
    top_runs = []   # up to 2 run plays

    pt = pct_series(sub["PLAY TYPE"])
    if pt:
        run_pct = next((p for v, p, c in pt if v == "Run"), BASE_RUN)
        pass_pct = next((p for v, p, c in pt if v == "Pass"), BASE_PASS)
        top = pt[0]
        if top[1] >= THRESHOLD_STRONG:
            trends.append((top[0], top[1], top[2]))

    # Top run plays (from run subset) + direction split for each
    run_sub = sub[sub["PLAY TYPE"] == "Run"]
    top_runs = []  # list of (play, pct, count, dir_breakdown)
    if len(run_sub) > 0:
        for play, pct, cnt in pct_series(run_sub["OFF PLAY"])[:2]:
            play_rows = run_sub[run_sub["OFF PLAY"].astype(str).str.upper() == str(play).upper()]
            dir_br = pct_series(play_rows["PLAY DIR"])
            # If no dir data, still keep the play
            top_runs.append((play, pct, cnt, dir_br if dir_br else []))

    plays = pct_series(sub["OFF PLAY"])
    if plays:
        top = plays[0]
        if top[1] >= THRESHOLD_STRONG:
            trends.append((f"{top[0]}", top[1], top[2]))

    dirs = pct_series(sub["PLAY DIR"])
    if dirs:
        top = dirs[0]
        if top[1] >= THRESHOLD_STRONG:
            trends.append((f"Direction {top[0]}", top[1], top[2]))

    motion = pct_series(sub["MOTION CAT"])
    if motion:
        top = motion[0]
        if top[1] >= THRESHOLD_STRONG:
            trends.append((top[0], top[1], top[2]))

    # Pass direction / side of field
    pass_sub = sub[sub["PLAY TYPE"] == "Pass"]
    if len(pass_sub) >= 1:
        pdir = pct_series(pass_sub["PLAY DIR"])
        if pdir:
            pass_dir = pdir[0]
            if pass_pct >= THRESHOLD_STRONG and pdir[0][1] >= 60:
                trends.append((f"Pass goes {pdir[0][0]}", pdir[0][1], pdir[0][2]))

    return trends, run_pct, pass_pct, pass_dir, top_runs


def predict(dn=None, dist_group=None, hash_=None, yard_zone=None,
            form=None, prev_play=None, allowed_forms=None):
    # --- Full filter first ---
    sub = filter_plays(dn=dn, dist_group=dist_group, hash_=hash_,
                       yard_zone=yard_zone, form=form, prev_play=prev_play,
                       allowed_forms=allowed_forms)
    n = len(sub)

    filters_used = {k: v for k, v in {
        "Down": dn, "Distance": dist_group, "Hash": hash_,
        "Field Zone": yard_zone, "Formation": form, "Prev Play": prev_play
    }.items() if v is not None}

    result = {
        "sample_size": n,
        "filters_used": filters_used,
        "trends": [],
        "got_ya": False,
        "fallback": False,
        "no_data": False,       # zero matching plays
        "relaxed": False,       # fell back to down/dist/zone only
        "run_pct": BASE_RUN,
        "pass_pct": BASE_PASS,
        "pass_dir": None,
        "top_runs": [],
    }

    # Case A: some plays match full criteria
    if n >= MIN_SAMPLE:
        trends, run_pct, pass_pct, pass_dir, top_runs = _build_trends(sub)
        result["run_pct"] = run_pct
        result["pass_pct"] = pass_pct
        result["pass_dir"] = pass_dir
        result["top_runs"] = top_runs
        result["trends"] = trends
        if trends:
            result["got_ya"] = max(t[1] for t in trends) >= THRESHOLD_GOTYA
        else:
            result["fallback"] = True   # plays exist, but no 75%+ tendency → Not Today video
        return result

    # Case B: too few / zero plays with full criteria
    # Relax to Down + Distance + Field Zone only
    result["no_data"] = (n == 0)
    relaxed = filter_plays(dn=dn, dist_group=dist_group, yard_zone=yard_zone,
                            allowed_forms=allowed_forms)
    n2 = len(relaxed)

    if n2 >= MIN_SAMPLE:
        trends, run_pct, pass_pct, pass_dir, top_runs = _build_trends(relaxed)
        result["sample_size"] = n2
        result["run_pct"] = run_pct
        result["pass_pct"] = pass_pct
        result["pass_dir"] = pass_dir
        result["top_runs"] = top_runs
        result["trends"] = trends
        result["relaxed"] = True
        result["filters_used"] = {k: v for k, v in {
            "Down": dn, "Distance": dist_group, "Field Zone": yard_zone
        }.items() if v is not None}
        if trends:
            result["got_ya"] = max(t[1] for t in trends) >= THRESHOLD_GOTYA
        else:
            result["fallback"] = True
        return result

    # Case C: still nothing useful → pure baseline, no video
    result["fallback"] = True
    result["no_data"] = True
    result["sample_size"] = n2
    return result

def show_video(path, max_height=280):
    path = Path(path)
    if not path.exists():
        return
    try:
        # Prefer HTML5 autoplay (works after user clicks PREDICT)
        import base64
        with open(path, "rb") as f:
            b64 = base64.b64encode(f.read()).decode()
        html = f"""
        <video width="100%" autoplay playsinline controls
               style="border-radius:10px; max-height:{max_height}px;">
            <source src="data:video/mp4;base64,{b64}" type="video/mp4">
        </video>
        """
        st.markdown(html, unsafe_allow_html=True)
    except Exception:
        try:
            st.video(str(path), start_time=0)
        except Exception:
            pass

# ------------------------------------------------------------------
# UI – page switcher
# ------------------------------------------------------------------
if "page" not in st.session_state:
    st.session_state.page = "Predict"

nav1, nav2 = st.columns(2)
with nav1:
    if st.button("Predict", use_container_width=True,
                 type="primary" if st.session_state.page == "Predict" else "secondary",
                 key="nav_predict"):
        st.session_state.page = "Predict"
        st.rerun()
with nav2:
    if st.button("Add Play", use_container_width=True,
                 type="primary" if st.session_state.page == "Add Play" else "secondary",
                 key="nav_add"):
        st.session_state.page = "Add Play"
        st.rerun()

st.markdown("<h1>MOSSER PREDICTOR 2026</h1>", unsafe_allow_html=True)
st.markdown('<p class="subtitle">Adam Central Football • 2025 Season</p>', unsafe_allow_html=True)

# ===================== ADD PLAY PAGE =====================
if st.session_state.page == "Add Play":
    st.markdown("### Add Live Play")
    st.caption("Enter a play as it happens. It is added to this session and used in predictions immediately.")

    all_forms = sorted([f for f in DF_MASTER["OFF FORM"].dropna().unique().tolist() if f])
    all_plays = sorted([f for f in DF_MASTER["OFF PLAY"].dropna().astype(str).unique().tolist() if f and f.upper() != "NAN"])
    all_motion = sorted([f for f in DF_MASTER["MOTION"].dropna().astype(str).unique().tolist() if f and f.upper() != "NAN"])
    all_pers = sorted([str(int(float(x))) for x in DF_MASTER["PERSONNEL"].dropna().unique().tolist()])

    with st.form("add_play_form", clear_on_submit=True):
        r1c1, r1c2, r1c3 = st.columns(3)
        with r1c1:
            a_dn = st.selectbox("Down", [1, 2, 3, 4])
        with r1c2:
            a_dist = st.selectbox("Distance", list(range(1, 26)))
        with r1c3:
            a_hash = st.selectbox("Hash", ["L", "M", "R"])

        r2c1, r2c2 = st.columns(2)
        with r2c1:
            a_yard = st.selectbox(
                "Yard Line",
                list(range(-49, 51)),
                index=49,  # ~0
                help="Negative = own side, Positive = opponent side"
            )
        with r2c2:
            a_form = st.selectbox("Formation", all_forms)

        r3c1, r3c2, r3c3 = st.columns(3)
        with r3c1:
            a_type = st.selectbox("Play Type", ["Run", "Pass"])
        with r3c2:
            a_play = st.selectbox("Play Name", all_plays)
        with r3c3:
            a_dir = st.selectbox("Play Direction", ["L", "R", "—"])

        r4c1, r4c2, r4c3 = st.columns(3)
        with r4c1:
            a_motion = st.selectbox("Motion", ["NONE"] + [m for m in all_motion if m != "NONE"])
        with r4c2:
            a_mdir = st.selectbox("Motion Dir", ["—", "L", "R"])
        with r4c3:
            a_str = st.selectbox("Off Strength", ["BAL", "L", "R"])

        r5c1, r5c2 = st.columns(2)
        with r5c1:
            a_pers = st.selectbox("Personnel", all_pers if all_pers else ["11"])
        with r5c2:
            a_gnls = st.number_input("Gain/Loss (optional)", value=0, step=1)

        submitted = st.form_submit_button("Add Play to Session", use_container_width=True, type="primary")

    if submitted:
        if "live_plays" not in st.session_state:
            st.session_state.live_plays = []
        new_id = 900000 + len(st.session_state.live_plays) + 1
        row = {
            "ID": new_id,
            "PLAY #": new_id,
            "SOURCE": "LIVE",
            "DN": float(a_dn),
            "DIST": float(a_dist),
            "HASH": a_hash,
            "YARD LN": float(a_yard),
            "GN/LS": float(a_gnls),
            "PLAY TYPE": a_type,
            "OFF FORM": str(a_form).upper(),
            "OFF STR": a_str,
            "PERSONNEL": float(a_pers),
            "OFF PLAY": str(a_play).upper(),
            "PLAY DIR": None if a_dir == "—" else a_dir,
            "MOTION": a_motion,
            "MOTION DIR": None if a_mdir == "—" else a_mdir,
            "DIST GROUP": _dist_group(a_dist),
            "YARD ZONE": _yard_zone(a_yard),
            "MOTION CAT": _motion_cat(a_motion, a_mdir if a_mdir != "—" else None),
            "PREV_PLAY_TYPE": np.nan,
            "DIST CATEGORY": np.nan,
            "FIELD ZONE": np.nan,
        }
        st.session_state.live_plays.append(row)
        st.success(f"Added: {a_type} — {a_play} ({a_form})  •  Total live plays: {len(st.session_state.live_plays)}")

    # Session play log
    st.markdown("---")
    n_live = len(st.session_state.get("live_plays", []))
    st.markdown(f"**Live plays this session: {n_live}**")
    if n_live:
        live_df = pd.DataFrame(st.session_state.live_plays)
        show_cols = [c for c in ["DN", "DIST", "HASH", "YARD LN", "PLAY TYPE", "OFF FORM", "OFF PLAY", "PLAY DIR", "MOTION"] if c in live_df.columns]
        st.dataframe(live_df[show_cols], use_container_width=True, hide_index=True)
        c1, c2 = st.columns(2)
        with c1:
            if st.button("Undo last play", use_container_width=True):
                st.session_state.live_plays.pop()
                st.rerun()
        with c2:
            if st.button("Clear all live plays", use_container_width=True):
                st.session_state.live_plays = []
                st.rerun()
        # Download combined
        combined = get_working_df()
        from io import BytesIO
        buf = BytesIO()
        combined.to_excel(buf, index=False, sheet_name="AC Offense 2025")
        st.download_button(
            "Download updated Excel (master + live)",
            data=buf.getvalue(),
            file_name="AC_Offense_with_live.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    st.stop()  # don't render predictor page

# ===================== PREDICT PAGE =====================

# ---- Vertical Football Field ----
st.markdown('<p class="field-title">FIELD POSITION</p>', unsafe_allow_html=True)

if "selected_zone" not in st.session_state:
    st.session_state.selected_zone = None
if "selected_zone_label" not in st.session_state:
    st.session_state.selected_zone_label = None

# Opponent end zone (top = TD)
st.markdown('<div class="endzone endzone-opp">OPPONENT END ZONE — TOUCHDOWN</div>', unsafe_allow_html=True)

for btn_label, zone_value in FIELD_BUTTONS:
    is_sel = st.session_state.selected_zone == zone_value
    display = f"▶ {btn_label} ◀" if is_sel else btn_label
    btn_type = "primary" if is_sel else "secondary"
    if st.button(display, key=f"field_{btn_label}", use_container_width=True, type=btn_type):
        st.session_state.selected_zone = zone_value
        st.session_state.selected_zone_label = btn_label
        st.rerun()

st.markdown('<div class="endzone endzone-own">OWN END ZONE</div>', unsafe_allow_html=True)

if st.session_state.selected_zone:
    st.markdown(
        f'<p class="zone-selected">Ball on: {st.session_state.selected_zone_label}</p>',
        unsafe_allow_html=True
    )
    if st.button("Clear field zone"):
        st.session_state.selected_zone = None
        st.session_state.selected_zone_label = None
        st.rerun()

st.markdown("---")

# ---- Other inputs ----
col1, col2 = st.columns(2)

with col1:
    hash_sel = st.selectbox("Hash", ["—", "L", "M", "R"], index=0)
    down_sel = st.selectbox("Down", ["—", "1", "2", "3", "4"], index=0)
    prev_sel = st.selectbox("Previous Play", ["—", "Run", "Pass"], index=0)

with col2:
    dist_groups = ["—", "10+", "10–7", "7–5", "5–3", "2–1"]
    dist_sel = st.selectbox("Distance", dist_groups, index=0)
    all_forms = sorted([f for f in DF_MASTER["OFF FORM"].dropna().unique().tolist() if f])

# ---- Active Formations menu ----
if "active_forms" not in st.session_state:
    st.session_state.active_forms = all_forms[:]
if "show_form_menu" not in st.session_state:
    st.session_state.show_form_menu = False

# Keep active_forms valid against current list
st.session_state.active_forms = [f for f in st.session_state.active_forms if f in all_forms]

def _open_form_menu():
    st.session_state.show_form_menu = True

def _close_form_menu():
    st.session_state.show_form_menu = False

n_sel = len(st.session_state.active_forms)
_wdf = get_working_df()
n_plays = int(_wdf["OFF FORM"].isin([str(f).upper() for f in st.session_state.active_forms]).sum()) if st.session_state.active_forms else len(_wdf)

col_af1, col_af2 = st.columns([2, 1])
with col_af1:
    live_tag = " • LIVE ONLY" if st.session_state.get("live_only") else ""
    st.button(
        f"⚙  Active Formations  ({n_sel}/{len(all_forms)}){live_tag}",
        on_click=_open_form_menu,
        use_container_width=True,
        key="btn_open_forms",
    )
with col_af2:
    st.caption(f"{n_plays} plays in pool")

# Formation picker menu (full-width panel when open)
if st.session_state.show_form_menu:
    st.markdown("---")
    st.markdown("### Active Formations")
    st.caption("☑ = active (used in predictions) &nbsp;&nbsp; ☐ = ignored")

    live_count = len(st.session_state.get("live_plays", []))
    st.checkbox(
        f"Live in Game Plays only  ({live_count} plays)",
        help="When checked, predictions use ONLY plays added this game on the Add Play page — season master data is ignored.",
        key="live_only",
    )
    if st.session_state.get("live_only") and live_count == 0:
        st.warning("No live plays yet. Add plays on the Add Play page first.")

    c_all, c_none, c_done = st.columns(3)
    with c_all:
        if st.button("Select All", use_container_width=True, key="forms_all"):
            for form_name in all_forms:
                st.session_state[f"chk_form_{form_name}"] = True
            st.session_state.active_forms = all_forms[:]
            st.rerun()
    with c_none:
        if st.button("Clear All", use_container_width=True, key="forms_none"):
            for form_name in all_forms:
                st.session_state[f"chk_form_{form_name}"] = False
            st.session_state.active_forms = []
            st.rerun()
    with c_done:
        if st.button("Done ✓", use_container_width=True, type="primary", key="forms_done"):
            st.session_state.show_form_menu = False
            st.rerun()

    # Checkbox grid — checked = active
    cols = st.columns(2)
    new_active = []
    for i, form_name in enumerate(all_forms):
        key = f"chk_form_{form_name}"
        if key not in st.session_state:
            st.session_state[key] = form_name in st.session_state.active_forms
        with cols[i % 2]:
            if st.checkbox(form_name, key=key):
                new_active.append(form_name)
    st.session_state.active_forms = new_active

    st.markdown("---")

active_forms = st.session_state.active_forms
form_choices = active_forms if active_forms else all_forms
form_sel = st.selectbox("Formation (this play)", ["—"] + form_choices, index=0)

st.markdown("")
predict_btn = st.button("PREDICT", use_container_width=True)

# ------------------------------------------------------------------
# Results
# ------------------------------------------------------------------
if predict_btn:
    dn = int(down_sel) if down_sel != "—" else None
    dist_group = dist_sel if dist_sel != "—" else None
    hash_ = hash_sel if hash_sel != "—" else None
    form = form_sel if form_sel != "—" else None
    prev_play = prev_sel if prev_sel != "—" else None
    yard_zone = st.session_state.selected_zone

    allowed = active_forms if active_forms else None
    result = predict(
        dn=dn, dist_group=dist_group, hash_=hash_,
        yard_zone=yard_zone, form=form, prev_play=prev_play,
        allowed_forms=allowed
    )

    st.markdown("---")

    if result["filters_used"]:
        used = "  •  ".join(f"{k}: {v}" for k, v in result["filters_used"].items())
        st.markdown(
            f'<p class="filters">Using → {used}<br>Sample: {result["sample_size"]} plays</p>',
            unsafe_allow_html=True
        )
    else:
        st.markdown(
            f'<p class="filters">No filters • Overall baseline • {result["sample_size"]} plays</p>',
            unsafe_allow_html=True
        )

    # GOT YA
    if result["got_ya"]:
        st.markdown('<div class="got-ya">', unsafe_allow_html=True)
        left, right = st.columns([1.3, 1])
        with left:
            st.markdown('<p class="phrase">GOT YA FUCKER</p>', unsafe_allow_html=True)
            for label, pct, cnt in result["trends"]:
                st.markdown(f'<p class="pred">{pct}%  {label}</p>', unsafe_allow_html=True)
                st.markdown(
                    f'<p class="detail">{cnt} of {result["sample_size"]} plays</p>',
                    unsafe_allow_html=True
                )

            # If Run is a strong trend, always show top 2 run plays + direction
            run_is_strong = any(
                str(t[0]).lower() == "run" and t[1] >= THRESHOLD_STRONG
                for t in result["trends"]
            )
            if run_is_strong:
                top_runs = result.get("top_runs") or []
                if top_runs:
                    st.markdown(
                        '<p class="detail" style="margin-top:0.8rem;"><b>Top Run Plays</b></p>',
                        unsafe_allow_html=True
                    )
                    for item in top_runs:
                        if len(item) == 4:
                            play, pct, cnt, dir_br = item
                        else:
                            play, pct, cnt = item[0], item[1], item[2]
                            dir_br = []
                        dir_str = ""
                        if dir_br:
                            dir_str = "  →  " + "  |  ".join(
                                f"{d} {dp}%" for d, dp, dc in dir_br[:3]
                            )
                        st.markdown(
                            f'<p class="detail">{pct}%  {play}  ({cnt}){dir_str}</p>',
                            unsafe_allow_html=True
                        )

            if result.get("pass_dir"):
                d, p, c = result["pass_dir"]
                st.markdown(
                    f'<p class="detail">Pass side tendency: {d} ({p}%)</p>',
                    unsafe_allow_html=True
                )
        with right:
            folder = Path(__file__).parent
            vp = folder / "got_ya.mp4"
            if not vp.exists():
                for p in folder.glob("*.mp4"):
                    if "not_today" not in p.name.lower():
                        vp = p
                        break
            show_video(vp)
        st.markdown('</div>', unsafe_allow_html=True)

    # Strong trend (≥75%)
    elif result["trends"]:
        st.markdown('<div class="trend-box">', unsafe_allow_html=True)
        st.markdown("### STRONG TREND")

        # 1. Strongest trend first
        strongest = max(result["trends"], key=lambda t: t[1])
        st.markdown(
            f'<p class="pred">{strongest[1]}%  {strongest[0]}</p>',
            unsafe_allow_html=True
        )
        st.markdown(
            f'<p class="detail">{strongest[2]} of {result["sample_size"]} plays</p>',
            unsafe_allow_html=True
        )

        # 2. Run / Pass split
        st.markdown(
            f'<p class="pred" style="font-size:1.5rem; margin-top:0.8rem;">'
            f'Run {result["run_pct"]}% &nbsp;&nbsp;|&nbsp;&nbsp; Pass {result["pass_pct"]}%</p>',
            unsafe_allow_html=True
        )

        # 3. Pass Destination
        st.markdown("**Pass Destination**")
        if result.get("pass_dir"):
            d, p, c = result["pass_dir"]
            st.markdown(
                f'<p class="detail">{p}%  goes {d}  ({c} plays)</p>',
                unsafe_allow_html=True
            )
        else:
            st.markdown('<p class="detail">No Pass Data</p>', unsafe_allow_html=True)

        # 4. Top Run Plays with direction %
        top_runs = result.get("top_runs") or []
        st.markdown("**Top Run Plays**")
        if top_runs:
            for item in top_runs:
                # support both old (3-tuple) and new (4-tuple) shapes
                if len(item) == 4:
                    play, pct, cnt, dir_br = item
                else:
                    play, pct, cnt = item[0], item[1], item[2]
                    dir_br = []
                dir_str = ""
                if dir_br:
                    dir_str = "  →  " + "  |  ".join(
                        f"{d} {dp}%" for d, dp, dc in dir_br[:3]
                    )
                st.markdown(
                    f'<p class="detail">{pct}%  {play}  ({cnt}){dir_str}</p>',
                    unsafe_allow_html=True
                )
        else:
            st.markdown('<p class="detail">No Run Data</p>', unsafe_allow_html=True)

        st.markdown('</div>', unsafe_allow_html=True)

    # Fallback / not enough data
    else:
        st.markdown('<div class="not-today">', unsafe_allow_html=True)

        # Plays matched full criteria but no 75%+ tendency → show video
        show_vid = (result.get("fallback") and not result.get("no_data") and not result.get("relaxed") and result["sample_size"] > 0)

        left, right = st.columns([1.2, 1] if show_vid else [1, 0.01])

        with left:
            if result.get("no_data") and result["sample_size"] == 0:
                st.markdown('<p class="phrase">NOT ENOUGH DATA</p>', unsafe_allow_html=True)
                st.markdown(
                    f'<p class="baseline">0 plays matched • Run {result["run_pct"]}% | Pass {result["pass_pct"]}%</p>',
                    unsafe_allow_html=True
                )
            elif result.get("relaxed"):
                st.markdown('<p class="phrase">NOT ENOUGH DATA</p>', unsafe_allow_html=True)
                st.caption("Relaxed to Down / Distance / Field only")
                st.markdown(
                    f'<p class="baseline">{result["sample_size"]} plays • Run {result["run_pct"]}% | Pass {result["pass_pct"]}%</p>',
                    unsafe_allow_html=True
                )
            else:
                # Had plays, no strong tendency
                st.markdown('<p class="phrase">NO TENDENCY</p>', unsafe_allow_html=True)
                st.markdown(
                    f'<p class="baseline">{result["sample_size"]} plays matched • Run {result["run_pct"]}% | Pass {result["pass_pct"]}%</p>',
                    unsafe_allow_html=True
                )

        if show_vid:
            with right:
                folder = Path(__file__).parent
                vp = folder / "not_today.mp4"
                if not vp.exists():
                    for p in folder.glob("*.mp4"):
                        if "got_ya" not in p.name.lower():
                            vp = p
                            break
                show_video(vp)

        st.markdown('</div>', unsafe_allow_html=True)

st.markdown("---")
st.markdown(
    '<div class="footer">Mosser Predictor 2026 • 75%+ Trend • 90%+ Got Ya</div>',
    unsafe_allow_html=True
)
