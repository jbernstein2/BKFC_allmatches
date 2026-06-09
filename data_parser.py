import pandas as pd
import re

# ─────────────────────────────────────────────
# CLEAN METRICS MAP (NO COLUMN INDEXING)
# ─────────────────────────────────────────────
STATS = [
    ("Goals", 6),
    ("xG", 7),
    ("Shots", 8),
    ("Shots on Target", 9),
    ("Passes", 11),
    ("Pass Accuracy %", 13),
    ("Possession %", 14),
    ("Duels Won %", 25),
    ("Corners", 38),
    ("Interceptions", 73),
    ("Clearances", 74),
    ("Fouls", 75),
    ("PPDA", 108),
]


# ─────────────────────────────────────────────
# SAFE CONVERSION
# ─────────────────────────────────────────────
def safe_float(x):
    try:
        return float(x)
    except:
        return 0.0


# ─────────────────────────────────────────────
# CONVERT ROW → DICT (KEY FIX)
# ─────────────────────────────────────────────
def row_to_dict(row):
    return {name: safe_float(row[col]) for name, col in STATS}


# ─────────────────────────────────────────────
# LOAD MATCH DATA
# ─────────────────────────────────────────────
def load_match_data(bkfc_file, opponent_file):

    bkfc_df = pd.read_excel(bkfc_file, header=None, engine="openpyxl")
    opp_df = pd.read_excel(opponent_file, header=None, engine="openpyxl")

    # ── BASELINES ─────────────────────────────
    bkfc_season_avg = row_to_dict(bkfc_df.iloc[1])
    all_opp_avg = row_to_dict(bkfc_df.iloc[2])
    opp_season_avg = row_to_dict(opp_df.iloc[1])

    # ── MATCHES ───────────────────────────────
    matches = []

    for i in range(3, len(bkfc_df)):
        row = bkfc_df.iloc[i]

        if pd.isna(row[0]) or pd.isna(row[1]):
            continue

        match_title = str(row[1])
        if "-" not in match_title:
            continue

        score = ""
        m = re.search(r"\d+:\d+", match_title)
        if m:
            score = m.group()

        opponent_name = match_title.split("-")[-1].replace(score, "").strip()

        matches.append({
            "index": i,
            "title": match_title,
            "date": str(row[0]),
            "competition": str(row[2]),
            "score": score,
            "opponent": opponent_name,
            "stats": row_to_dict(row)
        })

    if not matches:
        raise ValueError("No valid matches found.")

    return {
        "matches": matches,
        "selected_match": matches[-1],  # default latest
        "bkfc_season_avg": bkfc_season_avg,
        "all_opp_avg": all_opp_avg,
        "opp_season_avg": opp_season_avg
    }


# ─────────────────────────────────────────────
# FORMAT FOR REPORT GENERATOR
# ─────────────────────────────────────────────
def build_match_data(parsed, opponent_file=None):

    m = parsed["selected_match"]

    return {
        "match_title": m["title"],
        "match_date": m["date"],
        "competition": m["competition"],
        "score": m["score"],
        "opponent_name": m["opponent"],

        "match_bkfc": m["stats"],

        "bkfc_season_avg": parsed["bkfc_season_avg"],
        "all_opp_avg": parsed["all_opp_avg"],
        "opp_season_avg": parsed["opp_season_avg"],
    }
