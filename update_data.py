"""
update_data.py
Add April 21 and May 3 2026 match rows to 'Row per Player', then regenerate
all model outputs (Risk Scores, waterfall plots, model charts).

Run: python update_data.py
"""
import os, warnings, subprocess, sys
warnings.filterwarnings("ignore")

import numpy as np
import pandas as pd
import openpyxl
from datetime import datetime, timedelta

BASE = r"C:\Users\fredo\500 Project"

# ── New match lineups: player_id -> (is_starter: bool, minutes_played: int) ──
# April 21 2026: Brooklyn FC 1-1 Sacramento (Away)
# May 3  2026: Sacramento 3-2 Orange County SC (Home)
LINEUPS = {
    datetime(2026, 4, 21): {
        #  id  start  min   (based on confirmed ESPN/search lineup)
        0: (False,  0),   # Jacob Randolph    – backup GK, unused
        1: (True,  96),   # Danny Vitello     – started, full game
        2: (True,  60),   # Chibuike Ukaegbu  – started, subbed off 60'
        3: (False,  0),   # Rohan Chivukula   – backup GK, unused
        4: (False,  0),   # Rayan Djedje      – unused
        5: (True,  96),   # Michel Benitez    – started, scored pen 90+6
        6: (False, 51),   # Dominik Wanner    – sub on 45', played 51 min
        7: (True,  96),   # Freddy Kleemann   – started, full game
        8: (False,  0),   # Jared Timmer      – unused
        9: (True,  96),   # Lee Desmond (C)   – started, full game
       10: (True,  96),   # Ryan Spaulding    – started, full game
       11: (False, 36),   # Jack Gurr         – sub on 60', played 36 min
       12: (False,  0),   # Danny Crisostomo  – unused
       13: (False,  0),   # Pep Casas         – unused
       14: (True,  96),   # Blake Willey      – started, full game
       15: (False,  0),   # Mark-Anthony Kaye – unused
       16: (True,  45),   # Aaron Essel       – started, subbed off HT (45')
       17: (False,  0),   # Pierre Reedy      – unused
       18: (True,  71),   # Arturo Rodriguez  – started, subbed off 71'
       19: (False,  0),   # Memo Rodriguez    – unused
       20: (False, 25),   # Kyle Edwards      – sub on 71', played 25 min
       21: (False,  0),   # Da'vian Kimbrough – unused
       22: (True,  96),   # Forster Ajago     – started, full game
       23: (True,  71),   # Mayele Malango    – started, subbed off 71'
    },
    datetime(2026, 5, 3): {
        0: (False,  0),   # Jacob Randolph    – backup GK, unused
        1: (True,  96),   # Danny Vitello     – started, full game
        2: (False,  9),   # Chibuike Ukaegbu  – sub on 87', 9 min
        3: (False,  0),   # Rohan Chivukula   – unused
        4: (False,  0),   # Rayan Djedje      – unused
        5: (True,  96),   # Michel Benitez    – started, scored 90+5
        6: (False, 22),   # Dominik Wanner    – sub on 74', 22 min
        7: (False,  0),   # Freddy Kleemann   – unused
        8: (False,  0),   # Jared Timmer      – unused
        9: (True,  96),   # Lee Desmond (C)   – started, full game
       10: (True,  78),   # Ryan Spaulding    – started, subbed off 78'
       11: (True,  96),   # Jack Gurr         – started, full game
       12: (False,  0),   # Danny Crisostomo  – unused
       13: (True,  96),   # Pep Casas         – started, full game
       14: (False, 18),   # Blake Willey      – sub on 78', 18 min
       15: (False,  0),   # Mark-Anthony Kaye – unused
       16: (True,  87),   # Aaron Essel       – started, subbed off 87'
       17: (False,  0),   # Pierre Reedy      – unused
       18: (True,  96),   # Arturo Rodriguez  – started, scored 57'
       19: (False,  0),   # Memo Rodriguez    – unused
       20: (False, 22),   # Kyle Edwards      – sub on 74', scored 88'
       21: (False,  0),   # Da'vian Kimbrough – unused
       22: (True,  74),   # Forster Ajago     – started, subbed off 74'
       23: (True,  96),   # Mayele Malango    – started, full game
    },
}

# ── Helpers ───────────────────────────────────────────────────────────────────
def safe_num(v):
    """openpyxl cell value → float; handles None, ints, formula strings."""
    if v is None:
        return np.nan
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip()
    if s.startswith("="):
        try:
            return float(eval(s[1:]))
        except Exception:
            return np.nan
    try:
        return float(s)
    except (TypeError, ValueError):
        return np.nan


# ── Load existing Row per Player ──────────────────────────────────────────────
print("Loading existing data...")
wb = openpyxl.load_workbook(
    os.path.join(BASE, "data columns v2.xlsx"), data_only=True
)
ws = wb["Row per Player"]
headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
raw_rows = [
    [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
    for r in range(2, ws.max_row + 1)
]
df = pd.DataFrame(raw_rows, columns=headers)
df["match_date"] = pd.to_datetime(df["match_date"])
df["player_id"] = df["player_id"].astype(int)
df["minutes_on_date_num"] = df["minutes_on_date"].apply(safe_num).fillna(0)

print(f"  Loaded {len(df)} existing rows. Last date: {df['match_date'].max().date()}")

# ── Build player metadata (name, position) ────────────────────────────────────
player_meta = {}
for pid in range(24):
    prows = df[df["player_id"] == pid].sort_values("match_date")
    if len(prows) > 0:
        player_meta[pid] = {
            "name": prows.iloc[-1]["player_name"],
            "position": prows.iloc[-1]["position"],
        }

# ── Compute new rows in chronological order ───────────────────────────────────
new_rows = []  # list of dicts in column order

for match_date in sorted(LINEUPS.keys()):
    lineup = LINEUPS[match_date]
    print(f"\nComputing rows for {match_date.date()} ...")

    for pid in range(24):
        is_start, new_min = lineup[pid]
        was_absent = 1 if new_min == 0 else 0

        # All historical entries for this player (from original df + already-computed new rows)
        p_orig = df[df["player_id"] == pid].sort_values("match_date")

        hist = []
        for _, row in p_orig.iterrows():
            m = row["minutes_on_date_num"]
            ab_raw = row.get("was_absent_this_match")
            ab_num = safe_num(ab_raw)
            ab = int(ab_num) if not np.isnan(ab_num) else (1 if m == 0 else 0)
            hist.append({"date": row["match_date"], "minutes": m, "absent": ab,
                         "consecutive_starts": row.get("consecutive_starts"),
                         "injuries_30": safe_num(row.get("injuries_last_30_days")),
                         "injuries_60": safe_num(row.get("injuries_last_60_days"))})

        # Also include any rows already computed in this loop run (for May 3 needing April 21)
        for r in new_rows:
            if r["player_id"] == pid:
                hist.append({"date": r["match_date"], "minutes": float(r["minutes_on_date"]),
                              "absent": r["was_absent_this_match"],
                              "consecutive_starts": r["consecutive_starts"],
                              "injuries_30": r["injuries_last_30_days"],
                              "injuries_60": r["injuries_last_60_days"]})

        hist.sort(key=lambda x: x["date"])
        prev = [h for h in hist if h["date"] < match_date]

        # ── Core rolling features ──────────────────────────────────────────
        if prev:
            last = prev[-1]
            minutes_last_match   = last["minutes"]
            days_since_last_match = (match_date - last["date"]).days
            played_last_match    = 1 if last["minutes"] > 0 else 0
        else:
            minutes_last_match   = None
            days_since_last_match = None
            played_last_match    = None

        w7  = [h for h in prev if h["date"] >= match_date - timedelta(days=7)]
        w14 = [h for h in prev if h["date"] >= match_date - timedelta(days=14)]
        w30 = [h for h in prev if h["date"] >= match_date - timedelta(days=30)]

        minutes_last_7_days  = sum(h["minutes"] for h in w7)
        minutes_last_14_days = sum(h["minutes"] for h in w14)
        minutes_last_30_days = sum(h["minutes"] for h in w30)
        matches_last_14_days = len(w14)

        missed_matches_last_30_days = sum(1 for h in w30 if h["absent"] == 1)

        absences = [h for h in prev if h["absent"] == 1]
        days_since_last_absence = (
            (match_date - absences[-1]["date"]).days if absences else None
        )

        # ── Consecutive starts ─────────────────────────────────────────────
        if prev:
            cs_prev_val = prev[-1]["consecutive_starts"]
            cs_prev = int(cs_prev_val) if cs_prev_val is not None else 0
        else:
            cs_prev = 0
        consecutive_starts = cs_prev + 1 if is_start else 0

        # ── Injuries: carry forward last known values ──────────────────────
        if prev:
            inj30_v = prev[-1]["injuries_30"]
            inj60_v = prev[-1]["injuries_60"]
            injuries_30 = int(inj30_v) if not np.isnan(inj30_v) else 0
            injuries_60 = int(inj60_v) if not np.isnan(inj60_v) else 0
        else:
            injuries_30 = 0
            injuries_60 = 0

        # ── Target variable ────────────────────────────────────────────────
        # April 21: next match = May 3 (now known)
        # May 3: next match unknown
        if match_date == datetime(2026, 4, 21):
            next_min = LINEUPS[datetime(2026, 5, 3)][pid][1]
            player_unavailable_next_match = 1 if next_min == 0 else 0
        else:
            player_unavailable_next_match = None

        meta = player_meta.get(pid, {"name": f"Player_{pid}", "position": ""})

        row_dict = {
            "player_id":                    pid,
            "player_name":                  meta["name"],
            "position":                     meta["position"],
            "minutes_on_date":              new_min,
            "match_date":                   match_date,
            "minutes_last_match":           minutes_last_match,
            "minutes_last_7_days":          minutes_last_7_days,
            "minutes_last_14_days":         minutes_last_14_days,
            "minutes_last_30_days":         minutes_last_30_days,
            "matches_last_14_days":         matches_last_14_days,
            "consecutive_starts":           consecutive_starts,
            "played_last_match":            played_last_match,
            "days_since_last_match":        days_since_last_match,
            "was_absent_this_match":        was_absent,
            "missed_matches_last_30_days":  missed_matches_last_30_days,
            "days_since_last_absence":      days_since_last_absence,
            "injuries_last_30_days":        injuries_30,
            "injuries_last_60_days":        injuries_60,
            "player_unavailable_next_match": player_unavailable_next_match,
        }
        new_rows.append(row_dict)

    played = sum(1 for pid in range(24) if LINEUPS[match_date][pid][1] > 0)
    print(f"  {played} players played, {24 - played} absent")

# ── Append new rows to Excel ──────────────────────────────────────────────────
print(f"\nAppending {len(new_rows)} new rows to Excel...")
import gc
del wb, ws
gc.collect()

xlsx_path = os.path.join(BASE, "data columns v2.xlsx")

# Pre-flight: check if the file can be opened for writing
_can_write = False
try:
    with open(xlsx_path, "r+b"):
        _can_write = True
except PermissionError:
    pass

if not _can_write:
    print("\nWARNING: 'data columns v2.xlsx' is locked (likely open in Excel).")
    print("   Please CLOSE the file in Excel, then re-run this script.")
    sys.exit(1)

wb2 = openpyxl.load_workbook(xlsx_path)
ws2 = wb2["Row per Player"]

for row_dict in new_rows:
    ws2.append([row_dict.get(h) for h in headers])

wb2.save(xlsx_path)
print(f"Saved. Row per Player now has {ws2.max_row - 1} data rows.")

# ── Re-run model to regenerate risk scores and plots ─────────────────────────
print("\nRunning availability_model.py to regenerate outputs...\n")
result = subprocess.run(
    [sys.executable, os.path.join(BASE, "availability_model.py")],
    capture_output=True, text=True
)
print(result.stdout)
if result.returncode != 0:
    print("STDERR:", result.stderr[-2000:])
    sys.exit(1)

print("\nDone. Risk Scores updated and waterfall plots regenerated.")
