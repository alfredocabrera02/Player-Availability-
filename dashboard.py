"""
Sacramento Republic FC — Player Availability Risk Dashboard
Run with: streamlit run dashboard.py
"""

import streamlit as st
import pandas as pd
import numpy as np
import openpyxl
import os
import plotly.graph_objects as go

BASE = r"C:\Users\fredo\500 Project"
WATERFALL_DIR = os.path.join(BASE, "waterfall_plots")

st.set_page_config(
    page_title="Sacramento Republic FC — Availability Risk",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown("""
<style>
[data-testid="stAppViewContainer"] { background: #F8FAFC; }
.block-container { padding-top: 1.2rem; padding-bottom: 1rem; }
[data-testid="stMetricValue"] { font-size: 1.6rem !important; }
</style>
""", unsafe_allow_html=True)


# ── Data loading ─────────────────────────────────────────────────────────────
@st.cache_data
def load_data():
    wb = openpyxl.load_workbook(os.path.join(BASE, "data columns v2.xlsx"))

    ws = wb["Risk Scores"]
    hdrs = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    rows = [[ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
            for r in range(2, ws.max_row + 1)]
    risk_df = pd.DataFrame(rows, columns=hdrs)
    risk_df["risk_score"] = risk_df["risk_score"].astype(float)
    risk_df["risk_pct"] = (risk_df["risk_score"] * 100).round(1)

    ws2 = wb["Row per Player"]
    h2 = [ws2.cell(1, c).value for c in range(1, ws2.max_column + 1)]
    r2 = [[ws2.cell(r, c).value for c in range(1, ws2.max_column + 1)]
          for r in range(2, ws2.max_row + 1)]
    player_df = pd.DataFrame(r2, columns=h2)
    player_df["match_date"] = pd.to_datetime(player_df["match_date"])

    return risk_df, player_df


risk_df, player_df = load_data()

# Dashboard view is pinned to this date; data after it is ignored
_as_of = pd.Timestamp("2026-05-03")
_risk_dates = pd.to_datetime(risk_df["as_of_date"])
risk_df = risk_df[_risk_dates == _risk_dates[_risk_dates <= _as_of].max()].copy()
player_df = player_df[player_df["match_date"] <= _as_of].copy()
# %#d strips leading zero on Windows; fall back to zero-padded %d elsewhere
try:
    AS_OF_STR = _as_of.strftime("%B %#d, %Y")
except ValueError:
    AS_OF_STR = _as_of.strftime("%B %-d, %Y")

# ── Header (placed after data load so AS_OF_STR is available) ────────────────
st.markdown(f"""
<div style="background:#1B2E5E; border-radius:10px; padding:14px 24px; margin-bottom:16px;">
    <span style="font-size:22px; font-weight:800; color:#F4D03F;">
        Sacramento Republic FC
    </span>
    <span style="font-size:14px; color:#94A3B8; margin-left:16px;">
        Player Availability Risk Dashboard &nbsp;&middot;&nbsp; As of {AS_OF_STR}
    </span>
</div>
""", unsafe_allow_html=True)


# ── Helpers ──────────────────────────────────────────────────────────────────
def tier_color(tier):
    t = str(tier)
    if "High" in t:   return "#B9182F"
    if "Medium" in t: return "#D97706"
    return "#167A3C"

def tier_bg(tier):
    t = str(tier)
    if "High" in t:   return "#FEE2E2"
    if "Medium" in t: return "#FEF3C7"
    return "#D1FAE5"

def waterfall_path(name):
    safe = name.replace(" ", "_").replace("'", "").replace("/", "-")
    return os.path.join(WATERFALL_DIR, f"waterfall_{safe}.png")

def safe_float(v):
    """Convert openpyxl cell value to float; return NaN for formulas/None."""
    if v is None:
        return np.nan
    try:
        return float(v)
    except (TypeError, ValueError):
        return np.nan

def fmt_metric(v, is_absence=False):
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return "N/A"
    fv = safe_float(v)
    if np.isnan(fv):
        return "N/A"
    if is_absence and fv >= 999:
        return "None recent"
    return str(int(fv))


# ── Tabs ─────────────────────────────────────────────────────────────────────
tab1, tab2, tab3 = st.tabs(
    ["Squad Overview", "Player Detail", "Model Performance"]
)


# ══════════════════════════════════════════════════════════════════════════════
# TAB 1 — SQUAD OVERVIEW
# ══════════════════════════════════════════════════════════════════════════════
with tab1:
    n_high   = risk_df["risk_tier"].str.contains("High").sum()
    n_medium = risk_df["risk_tier"].str.contains("Medium").sum()
    n_low    = len(risk_df) - n_high - n_medium

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total Players",          len(risk_df))
    c2.metric("🔴  High Risk  (>35%)",  int(n_high))
    c3.metric("🟡  Medium Risk (15-35%)", int(n_medium))
    c4.metric("🟢  Low Risk  (<15%)",   int(n_low))

    st.markdown("---")

    # Horizontal bar chart — sorted lowest → highest so highest appears on top
    sorted_df = risk_df.sort_values("risk_pct")
    bar_colors = [tier_color(t) for t in sorted_df["risk_tier"]]

    fig_bar = go.Figure(go.Bar(
        x=sorted_df["risk_pct"],
        y=sorted_df["player_name"],
        orientation="h",
        marker_color=bar_colors,
        text=[f"{v:.1f}%" for v in sorted_df["risk_pct"]],
        textposition="outside",
        cliponaxis=False,
        hovertemplate="<b>%{y}</b><br>Risk: %{x:.1f}%<extra></extra>",
    ))
    fig_bar.add_vline(x=15, line_dash="dash", line_color="#167A3C", line_width=1.5,
                      annotation_text="15% threshold", annotation_font_size=11,
                      annotation_position="top right")
    fig_bar.add_vline(x=35, line_dash="dash", line_color="#B9182F", line_width=1.5,
                      annotation_text="35% threshold", annotation_font_size=11,
                      annotation_position="top right")
    fig_bar.update_layout(
        title=dict(text=f"Player Unavailability Risk — Current Squad ({AS_OF_STR})",
                   font=dict(size=16, color="#1E293B")),
        xaxis_title="Risk Score (%)",
        yaxis_title="",
        height=660,
        plot_bgcolor="white",
        paper_bgcolor="#F8FAFC",
        margin=dict(l=190, r=80, t=55, b=40),
        xaxis=dict(range=[0, 100], gridcolor="#E2E8F0", ticksuffix="%"),
        font=dict(family="Inter, Arial, sans-serif"),
        showlegend=False,
    )
    st.plotly_chart(fig_bar, use_container_width=True)

    # Styled table
    st.markdown("### Player Risk Table")
    display_df = risk_df[["player_name", "risk_pct", "risk_tier"]].copy()
    display_df.columns = ["Player", "Risk Score (%)", "Risk Tier"]
    display_df = display_df.sort_values("Risk Score (%)", ascending=False).reset_index(drop=True)

    def _style_tier(val):
        if "High" in str(val):
            return "background-color:#FEE2E2; color:#B9182F; font-weight:bold"
        if "Medium" in str(val):
            return "background-color:#FEF3C7; color:#D97706; font-weight:bold"
        return "background-color:#D1FAE5; color:#167A3C; font-weight:bold"

    styled = display_df.style.map(_style_tier, subset=["Risk Tier"])
    st.dataframe(styled, use_container_width=True, hide_index=True)


# ══════════════════════════════════════════════════════════════════════════════
# TAB 2 — PLAYER DETAIL
# ══════════════════════════════════════════════════════════════════════════════
with tab2:
    player_names = (risk_df.sort_values("risk_pct", ascending=False)["player_name"].tolist())
    selected = st.selectbox("Select a player", player_names)

    player_row = risk_df[risk_df["player_name"] == selected].iloc[0]
    risk_pct = float(player_row["risk_pct"])
    tier     = str(player_row["risk_tier"])

    # Risk score card + feature metrics side by side
    col_card, col_feats = st.columns([1, 3], gap="large")

    with col_card:
        st.markdown(f"""
        <div style="background:{tier_bg(tier)}; border:2px solid {tier_color(tier)};
                    border-radius:12px; padding:28px 20px; text-align:center; margin-top:8px;">
            <div style="font-size:52px; font-weight:900; color:{tier_color(tier)};
                        line-height:1;">{risk_pct:.1f}%</div>
            <div style="font-size:14px; color:{tier_color(tier)}; margin-top:6px;
                        font-weight:600;">Unavailability Risk</div>
            <div style="margin-top:10px; font-size:12px; color:#64748B; background:white;
                        border-radius:6px; padding:4px 8px; display:inline-block;">
                {tier}
            </div>
        </div>
        """, unsafe_allow_html=True)

    with col_feats:
        p_latest = (player_df[
            (player_df["player_name"] == selected) &
            (player_df["match_date"].dt.year == 2026)
        ].sort_values("match_date", ascending=False))

        if len(p_latest) > 0:
            lat = p_latest.iloc[0]
            fc1, fc2, fc3 = st.columns(3)

            played_label = (
                "Yes" if lat.get("played_last_match") == 1
                else "No" if lat.get("played_last_match") == 0
                else "N/A"
            )

            with fc1:
                st.metric("Min. Last Match",    fmt_metric(lat.get("minutes_last_match")))
                st.metric("Min. Last 7 Days",   fmt_metric(lat.get("minutes_last_7_days")))
                st.metric("Min. Last 14 Days",  fmt_metric(lat.get("minutes_last_14_days")))
                st.metric("Min. Last 30 Days",  fmt_metric(lat.get("minutes_last_30_days")))
            with fc2:
                st.metric("Consecutive Starts", fmt_metric(lat.get("consecutive_starts")))
                st.metric("Matches (14 Days)",  fmt_metric(lat.get("matches_last_14_days")))
                st.metric("Played Last Match",  played_label)
                st.metric("Days Since Last Match", fmt_metric(lat.get("days_since_last_match")))
            with fc3:
                st.metric("Missed Matches (30d)", fmt_metric(lat.get("missed_matches_last_30_days")))
                st.metric("Days Since Absence",   fmt_metric(lat.get("days_since_last_absence"), is_absence=True))
                st.metric("Injuries (30 Days)",   fmt_metric(lat.get("injuries_last_30_days")))
                st.metric("Injuries (60 Days)",   fmt_metric(lat.get("injuries_last_60_days")))
        else:
            st.info("No 2026 feature data available for this player.")

    st.markdown("---")

    # Workload history (minutes per match across all seasons)
    p_history = (player_df[player_df["player_name"] == selected]
                 .sort_values("match_date")
                 .copy())
    p_history["mins"] = p_history["minutes_on_date"].apply(safe_float).fillna(0)

    if len(p_history) > 0:
        st.markdown("#### Workload History — Minutes Per Match")

        season_colors = {2024: "#1B2E5E", 2025: "#C4A416", 2026: "#B9182F"}
        fig_hist = go.Figure()
        for season, grp in p_history.groupby(p_history["match_date"].dt.year):
            fig_hist.add_trace(go.Bar(
                x=grp["match_date"],
                y=grp["mins"],
                name=str(season),
                marker_color=season_colors.get(season, "#94A3B8"),
                hovertemplate="<b>%{x|%b %d, %Y}</b><br>Minutes: %{y}<extra></extra>",
            ))
        fig_hist.add_hline(y=90, line_dash="dot", line_color="#94A3B8", line_width=1,
                           annotation_text="Full match (90 min)", annotation_font_size=10,
                           annotation_position="bottom right")
        fig_hist.update_layout(
            xaxis_title="Match Date",
            yaxis_title="Minutes Played",
            height=270,
            plot_bgcolor="white",
            paper_bgcolor="#F8FAFC",
            margin=dict(l=50, r=20, t=20, b=40),
            barmode="stack",
            legend=dict(orientation="h", yanchor="bottom", y=1.02, x=0),
            xaxis=dict(gridcolor="#E2E8F0"),
            yaxis=dict(gridcolor="#E2E8F0", range=[0, 110]),
            font=dict(family="Inter, Arial, sans-serif"),
        )
        st.plotly_chart(fig_hist, use_container_width=True)

    # SHAP waterfall plot
    wf = waterfall_path(selected)
    if os.path.exists(wf):
        st.markdown("#### SHAP Waterfall — Risk Factor Breakdown")
        st.image(wf, use_container_width=True)
    else:
        st.warning(f"Waterfall plot not found for {selected}.")


# ══════════════════════════════════════════════════════════════════════════════
# TAB 3 — MODEL PERFORMANCE
# ══════════════════════════════════════════════════════════════════════════════
with tab3:
    st.markdown("### Model Performance Metrics")
    st.caption("Train: 2024 season (196 rows) · Test: 2025–2026 seasons (564 rows) · Threshold: 0.35")

    # Metrics row
    r1c1, r1c2, r1c3, r1c4, r1c5, r1c6 = st.columns(6)
    r1c1.metric("LR  ROC-AUC",  "0.7825")
    r1c2.metric("LR  PR-AUC",   "0.6839")
    r1c3.metric("LR  Brier",    "0.2022")
    r1c4.metric("LGB ROC-AUC",  "0.7710")
    r1c5.metric("LGB PR-AUC",   "0.6657")
    r1c6.metric("LGB Brier",    "0.1854")

    st.markdown("---")

    col_left, col_right = st.columns(2)

    with col_left:
        st.markdown("#### Precision-Recall Curves")
        pr = os.path.join(BASE, "pr_curves.png")
        if os.path.exists(pr):
            st.image(pr, use_container_width=True)

        st.markdown("#### Confusion Matrices (threshold = 0.35)")
        cm = os.path.join(BASE, "confusion_matrices.png")
        if os.path.exists(cm):
            st.image(cm, use_container_width=True)

    with col_right:
        st.markdown("#### SHAP Feature Importance (LightGBM)")
        si = os.path.join(BASE, "shap_importance.png")
        if os.path.exists(si):
            st.image(si, use_container_width=True)

        st.markdown("#### SHAP Beeswarm — Risk Drivers & Mitigators")
        sb = os.path.join(BASE, "shap_beeswarm.png")
        if os.path.exists(sb):
            st.image(sb, use_container_width=True)

    st.markdown("---")
    st.markdown("#### Top 6 Highest-Risk Players — SHAP Waterfall Grid")
    top6 = os.path.join(BASE, "shap_waterfall_top6.png")
    if os.path.exists(top6):
        st.image(top6, use_container_width=True)
