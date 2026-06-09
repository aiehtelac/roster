import sys
import io
import copy
import os
import tempfile

import streamlit as st

sys.path.insert(0, os.path.dirname(__file__))
from RosterScheduler_Combined import (
    build_scheduler,
    ROSTER_CONFIGS,
    get_singapore_ph,
    fetch_singapore_ph,
)

st.set_page_config(page_title="Roster Scheduler", layout="wide")
st.title("Roster Scheduler")

# ── Session state ─────────────────────────────────────────────────────────────

for _k, _v in [("result_bytes", None), ("result_name", None),
               ("solve_log", None), ("error", None)]:
    if _k not in st.session_state:
        st.session_state[_k] = _v

# ── Sidebar: roster type & file uploads ──────────────────────────────────────

with st.sidebar:
    st.header("Setup")
    roster_type = st.radio("Roster type", ["HO", "MO", "REG"], horizontal=True)

    st.subheader("Files")
    csv_file  = st.file_uploader("Staff CSV *", type="csv")
    tpl_file  = st.file_uploader("Excel template", type=["xlsx", "xlsm"])
    prev_file = st.file_uploader(
        "Previous month xlsx", type=["xlsx", "xlsm"],
        disabled=tpl_file is None,
        help="Only used when a template is provided",
    )

# ── Tabs ──────────────────────────────────────────────────────────────────────

tab_ph, tab_cfg, tab_run = st.tabs(["Public Holidays", "Config", "Run"])

# ── Public Holidays ───────────────────────────────────────────────────────────

with tab_ph:
    st.subheader("Singapore Public Holidays")

    year = int(st.number_input("Year", min_value=2024, max_value=2030,
                               value=2026, step=1))

    ph_key = f"ph_text_{year}"
    if ph_key not in st.session_state:
        try:
            st.session_state[ph_key] = "\n".join(get_singapore_ph(year))
        except ValueError:
            st.session_state[ph_key] = ""
            st.warning(f"No hardcoded holidays for {year}. "
                       "Add the year to `_SG_PUBLIC_HOLIDAYS` in the scheduler.")

    if st.button("Fetch from data.gov.sg"):
        fetched = fetch_singapore_ph(year)
        if fetched:
            st.session_state[ph_key] = "\n".join(fetched)
            st.success(f"Fetched {len(fetched)} holidays")
        else:
            st.error("Fetch failed — keeping current list")

    ph_text = st.text_area(
        "Public holidays — one per line (D-Mon-YY). Edit to add or remove.",
        value=st.session_state[ph_key],
        height=320,
        key=ph_key + "_area",
    )
    st.session_state[ph_key] = ph_text

    final_phs = [h.strip() for h in ph_text.splitlines() if h.strip()]
    st.caption(f"{len(final_phs)} PHs active")

# ── Config ────────────────────────────────────────────────────────────────────

with tab_cfg:
    cfg = copy.deepcopy(ROSTER_CONFIGS[roster_type])
    sc  = cfg["shift_categories"]

    _CAT_LABELS = {
        "main":          "Main shifts",
        "half":          "Half call",
        "weekend_extra": "Extra weekend shifts",
        "cicu":          "CICU",
        "hybrid_shift":  "Hybrid shift (half-call weekdays, full call weekends)",
        "wr":            "Weekend Round",
        "sb":            "SB"
    }
    _DAY_LABELS = {
        "weekday": "Weekday", "saturday": "Saturday",
        "sunday":  "Sunday",  "ph":       "PH",
    }
    _DAY_TYPES  = ["weekday", "saturday", "sunday", "ph"]
    _CAT_ORDER  = ["main", "weekend_extra", "cicu", "hybrid_shift", "wr", "sb", "half"]

    # ── Shift Categories ──────────────────────────────────────────────────────
    st.subheader("Shift Categories")
    st.caption(
        "Edit shift names as comma-separated values. "
        "Clear a field to disable that category. "
    )

    for cat in _CAT_ORDER:
        names_str = st.text_input(
            _CAT_LABELS[cat],
            value=", ".join(sc.get(cat, {}).get("names", [])),
            key=f"sc_names_{roster_type}_{cat}",
        )
        sc.setdefault(cat, {})["names"] = [
            n.strip().upper() for n in names_str.split(",") if n.strip()
        ]

    st.divider()

    # ── Slot Counts ───────────────────────────────────────────────────────────
    st.subheader("Slot Counts per Day")
    st.caption("Slots to fill per day type for WR, SB, and half-call shifts. Main, CICU, weekend extra, and hybrid shifts always fill once each day.")

    slot_cats = [c for c in ("wr", "sb", "half") if sc.get(c, {}).get("names")]

    if slot_cats:
        hdr_cols = st.columns([2] + [1] * len(slot_cats))
        hdr_cols[0].markdown("**Day type**")
        for i, cat in enumerate(slot_cats):
            hdr_cols[i + 1].markdown(f"**{_CAT_LABELS[cat]}**")

        for dt in _DAY_TYPES:
            row_cols = st.columns([2] + [1] * len(slot_cats))
            row_cols[0].markdown(_DAY_LABELS[dt])
            for i, cat in enumerate(slot_cats):
                slots = sc[cat].setdefault("slots", {})
                slots[dt] = row_cols[i + 1].number_input(
                    label=f"{dt}_{cat}",
                    label_visibility="collapsed",
                    min_value=0, max_value=20,
                    value=int(slots.get(dt, 0)),
                    key=f"sc_slots_{roster_type}_{dt}_{cat}",
                )

        if sc.get("wr", {}).get("names"):
            sc["wr"]["ph_after_sunday"] = st.number_input(
                "Special case: WR slots on a PH following a Sunday",
                min_value=0, max_value=10,
                value=int(sc["wr"].get("ph_after_sunday", 0)),
                key=f"sc_wr_phas_{roster_type}",
            )
    else:
        st.caption("No slot-based shifts defined — add names above to configure slots.")

    st.divider()

    # ── Call Points ───────────────────────────────────────────────────────────
    st.subheader("Call Points")
    scale = cfg.get("call_points_scale", 1)
    if scale > 1:
        st.caption(f"Stored ×{scale} for integer arithmetic for code — actual call points = stored ÷ {scale}.")

    cp      = cfg["call_points"]
    cp_cols = st.columns(len(cp))
    for i, (day, pts) in enumerate(cp.items()):
        cp[day] = cp_cols[i].number_input(
            day, min_value=0, max_value=50, value=int(pts),
            key=f"cp_{roster_type}_{day}",
        )

    if "r3_points" in cfg:
        st.markdown("**Hybrid shift points** — half-call on weekdays, full call on weekends")
        r3p      = cfg["r3_points"]
        r3p_cols = st.columns(len(r3p))
        for i, (day, pts) in enumerate(r3p.items()):
            r3p[day] = r3p_cols[i].number_input(
                day, min_value=0, max_value=50, value=int(pts),
                key=f"r3p_{roster_type}_{day}",
            )

    if "ho6_points" in cfg:
        cfg["ho6_points"] = st.number_input(
            "Half call points (HO6)",
            min_value=0, max_value=20, value=int(cfg["ho6_points"]),
            help=f"Stored ×{scale}. Set to {scale} for 1.0 displayed point.",
            key=f"ho6pts_{roster_type}",
        )

    st.divider()

    # ── Limits ────────────────────────────────────────────────────────────────
    st.subheader("Limits")
    lim      = cfg["limits"]
    lim_cols = st.columns(min(len(lim), 4))
    for i, (k, v) in enumerate(lim.items()):
        with lim_cols[i % len(lim_cols)]:
            lim[k] = st.number_input(
                k, min_value=0, max_value=20, value=int(v),
                key=f"lim_{roster_type}_{k}",
            )

    st.divider()

    # ── Soft Penalty Weights ──────────────────────────────────────────────────
    st.subheader("Soft Penalty Weights")
    pen      = cfg["soft_penalties"]
    pen_cols = st.columns(len(pen))
    for i, (k, v) in enumerate(pen.items()):
        pen[k] = pen_cols[i].number_input(
            k, min_value=0, max_value=500, value=int(v),
            key=f"pen_{roster_type}_{k}",
        )

    st.divider()

    # ── Fairness Weights ──────────────────────────────────────────────────────
    st.subheader("Fairness Weights")
    for pool in cfg["fairness_pools"]:
        st.markdown(f"**{pool['label']}**")
        metric_keys = list(pool["metrics"].items())
        m_cols      = st.columns(min(len(metric_keys), 4))
        for i, (k, default) in enumerate(metric_keys):
            with m_cols[i % len(m_cols)]:
                pool["metrics"][k] = st.number_input(
                    k, min_value=0, max_value=200, value=int(default),
                    key=f"m_{roster_type}_{pool['label']}_{k}",
                )

    st.divider()

    # ── Solver ────────────────────────────────────────────────────────────────
    st.subheader("Solver")
    time_limit = st.slider("Time limit (seconds)", 30, 600, 300, step=30)

# ── Run ───────────────────────────────────────────────────────────────────────

with tab_run:
    st.subheader("Generate Roster")

    if not csv_file:
        st.info("Upload a staff CSV in the sidebar to continue.")

    if st.button("Generate Roster", disabled=not csv_file, type="primary"):
        st.session_state.result_bytes = None
        st.session_state.result_name  = None
        st.session_state.solve_log    = None
        st.session_state.error        = None

        log_buf    = io.StringIO()
        old_stdout = sys.stdout

        try:
            with st.spinner("Solving… this may take a few minutes"):
                with tempfile.TemporaryDirectory() as tmpdir:

                    csv_path = os.path.join(tmpdir, "input.csv")
                    with open(csv_path, "wb") as f:
                        f.write(csv_file.getvalue())

                    tpl_path = None
                    if tpl_file:
                        tpl_path = os.path.join(tmpdir, "template.xlsx")
                        with open(tpl_path, "wb") as f:
                            f.write(tpl_file.getvalue())

                    prev_path = None
                    if prev_file and tpl_file:
                        prev_path = os.path.join(tmpdir, "prev.xlsx")
                        with open(prev_path, "wb") as f:
                            f.write(prev_file.getvalue())

                    sys.stdout = log_buf
                    try:
                        scheduler = build_scheduler(
                            roster_type, csv_path, tmpdir,
                            copy.deepcopy(cfg), tpl_path, prev_path,
                        )
                        scheduler.set_public_holidays(final_phs)
                        scheduler.load_data()
                        df = scheduler.solve_and_export(
                            time_limit=float(time_limit),
                        )
                    finally:
                        sys.stdout = old_stdout

                    st.session_state.solve_log = log_buf.getvalue()

                    if df is None:
                        st.session_state.error = (
                            "Solver found no solution. "
                            "Check the log — likely infeasible constraints or too few staff."
                        )
                    else:
                        skip = {"template.xlsx", "prev.xlsx", "input.csv"}

                        # Prefer xlsx output, fall back to csv
                        output_bytes, output_name = None, None
                        for ext in (".xlsx", ".csv"):
                            for fname in sorted(os.listdir(tmpdir)):
                                if fname not in skip and fname.endswith(ext):
                                    with open(os.path.join(tmpdir, fname), "rb") as f:
                                        output_bytes = f.read()
                                    output_name = fname
                                    break
                            if output_bytes:
                                break

                        st.session_state.result_bytes = output_bytes
                        st.session_state.result_name  = output_name

        except Exception as exc:
            sys.stdout = old_stdout
            st.session_state.error     = str(exc)
            st.session_state.solve_log = log_buf.getvalue()

    # ── Output ────────────────────────────────────────────────────────────────

    if st.session_state.error:
        st.error(st.session_state.error)

    if st.session_state.result_bytes:
        fname = st.session_state.result_name or "roster_output"
        mime  = (
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            if fname.endswith(".xlsx") else "text/csv"
        )
        st.success(f"Roster generated: **{fname}**")
        st.download_button(
            label="⬇ Download",
            data=st.session_state.result_bytes,
            file_name=fname,
            mime=mime,
        )

    if st.session_state.solve_log:
        with st.expander("Solver log"):
            st.text(st.session_state.solve_log)
