import pandas as pd
import re

# ── METRICS MAP ─────────────────────────────────────────────
STATS = [
    {"label": "Goals", "col": 6},
    {"label": "xG (Expected Goals)", "col": 7},
    {"label": "Shots", "col": 8},
    {"label": "Shots on Target", "col": 9},
    {"label": "Shot Accuracy %", "col": 10},
    {"label": "Possession %", "col": 14},
    {"label": "Passes", "col": 11},
    {"label": "Pass Accuracy %", "col": 13},
    {"label": "Corners", "col": 38},
    {"label": "Interceptions", "col": 73},
    {"label": "Clearances", "col": 74},
    {"label": "Fouls", "col": 75},
    {"label": "PPDA", "col": 108},
]


# ── CORE PARSER ─────────────────────────────────────────────
def load_match_data(bkfc_file, opponent_file):
    bkfc_df = pd.read_excel(bkfc_file, header=None, engine="openpyxl")
    opp_df = pd.read_excel(opponent_file, header=None, engine="openpyxl")

    # ── BASELINE ROWS ───────────────────────────────────────
    bkfc_season_avg = bkfc_df.iloc[1]
    all_opp_avg = bkfc_df.iloc[2]
    opp_season_avg = opp_df.iloc[1]

    # ── MATCH ROW DETECTION ────────────────────────────────
    matches = []

    for i in range(3, len(bkfc_df)):
        row = bkfc_df.iloc[i]

        # skip empty rows
        if pd.isna(row[0]) or pd.isna(row[1]):
            continue

        match_title = str(row[1])
        match_date = str(row[0])
        competition = str(row[2])

        # only keep actual match rows
        if "-" not in match_title:
            continue

        # extract score
        score = ""
        m = re.search(r"\d+:\d+", match_title)
        if m:
            score = m.group()

        opponent_name = match_title.split("-")[-1].replace(score, "").strip()

        matches.append({
            "match_index": i,
            "match_title": match_title,
            "match_date": match_date,
            "competition": competition,
            "score": score,
            "opponent_name": opponent_name,
            "match_row": row
        })

    if not matches:
        raise ValueError("No match rows detected in BKFC file.")

    # default = most recent match
    selected = matches[-1]

    return {
        "matches": matches,
        "selected_match": selected,
        "bkfc_season_avg": bkfc_season_avg,
        "all_opp_avg": all_opp_avg,
        "opp_season_avg": opp_season_avg
    }


# ── HELPER: build full report-ready structure ──────────────
def build_match_data(parsed_data, opponent_file):
    """
    Converts selected match into report generator format
    """
    opp_df = pd.read_excel(opponent_file, header=None, engine="openpyxl")

    match = parsed_data["selected_match"]
    bkfc_row = match["match_row"]

    # opponent row in BKFC dataset (same match, second row)
    match_index = match["match_index"]
    opp_row = parsed_data["all_opp_avg"]  # fallback safety baseline

    return {
        "match_title": match["match_title"],
        "match_date": match["match_date"],
        "competition": match["competition"],
        "score": match["score"],
        "opponent_name": match["opponent_name"],

        "bkfc_season_avg": parsed_data["bkfc_season_avg"],
        "all_opp_avg": parsed_data["all_opp_avg"],
        "opp_season_avg": parsed_data["opp_season_avg"],

        "match_bkfc": bkfc_row,
        "match_opp": opp_row,   # fallback baseline opponent context
    }
