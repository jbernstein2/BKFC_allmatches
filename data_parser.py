import pandas as pd
import re

# --------------------------------------------------
# STATS (UNCHANGED — KEEP YOUR EXISTING VERSION)
# --------------------------------------------------

STATS = [
    {"label": "Goals", "col": 6},
    {"label": "xG (Expected Goals)", "col": 7},
    {"label": "Shots", "col": 8},
    {"label": "Shots on Target", "col": 9},
    {"label": "Shot Accuracy %", "col": 10},
    {"label": "Possession %", "col": 14},
    {"label": "Passes", "col": 11},
    {"label": "Pass Accuracy %", "col": 13},
    {"label": "Positional Attacks", "col": 29},
    {"label": "Positional Attacks w/ Shots", "col": 30},
    {"label": "Counterattacks", "col": 32},
    {"label": "Corners", "col": 38},
    {"label": "Set Pieces w/ Shots", "col": 36},
    {"label": "Duels Won %", "col": 25},
    {"label": "Crosses", "col": 47},
    {"label": "Cross Accuracy %", "col": 49},
    {"label": "Touches in Penalty Area", "col": 55},
    {"label": "Offensive Duels Won %", "col": 58},
    {"label": "Defensive Duels Won %", "col": 66},
    {"label": "Aerial Duels Won %", "col": 69},
    {"label": "Interceptions", "col": 73},
    {"label": "Clearances", "col": 74},
    {"label": "Fouls", "col": 75},
    {"label": "Yellow Cards", "col": 76},
    {"label": "Shots Against", "col": 61},
    {"label": "PPDA", "col": 108},
]

# --------------------------------------------------
# MATCH DISCOVERY FUNCTION (NEW)
# --------------------------------------------------

def get_available_matches(bkfc_file):
    """
    Builds dropdown list of all matches in BKFC workbook.
    Each match = BKFC row + Opponent row (every 2 rows).
    """

    bkfc_file.seek(0)

    df = pd.read_excel(
        bkfc_file,
        header=None,
        engine="openpyxl"
    )

    matches = []

    for i in range(3, len(df), 2):
        try:
            bkfc_row = df.iloc[i]

            if pd.isna(bkfc_row[0]):
                continue

            match_title = str(bkfc_row[1])

            if match_title.lower() == "nan":
                continue

            score_match = re.search(r"\d+:\d+", match_title)
            score = score_match.group() if score_match else ""

            opponent = (
                match_title
                .split("-")[-1]
                .replace(score, "")
                .strip()
            )

            matches.append({
                "display": f"{bkfc_row[0]} | {opponent} | {bkfc_row[2]}",
                "row_idx": i,
                "opponent": opponent,
                "date": str(bkfc_row[0])
            })

        except Exception:
            continue

    return matches


# --------------------------------------------------
# CORE MATCH LOADER (UPDATED)
# --------------------------------------------------

def load_match_data(bkfc_file, opponent_file, selected_row):
    """
    Loads a specific match using dropdown selection.
    """

    bkfc_file.seek(0)
    opponent_file.seek(0)

    bkfc_df = pd.read_excel(
        bkfc_file,
        header=None,
        engine="openpyxl"
    )

    opp_df = pd.read_excel(
        opponent_file,
        header=None,
        engine="openpyxl"
    )

    # Season baselines
    bkfc_season_avg = bkfc_df.iloc[1]
    all_opp_avg = bkfc_df.iloc[2]

    # Match-specific rows
    match_bkfc = bkfc_df.iloc[selected_row]
    match_opp = bkfc_df.iloc[selected_row + 1]

    # Opponent season baseline (from opponent file)
    opp_season_avg = opp_df.iloc[1]

    # Metadata parsing
    match_title = str(match_bkfc[1])
    competition = str(match_bkfc[2])
    match_date = str(match_bkfc[0])

    score_match = re.search(r"\d+:\d+", match_title)
    score = score_match.group() if score_match else ""

    opponent_name = (
        match_title
        .split("-")[-1]
        .replace(score, "")
        .strip()
    )

    return {
        "match_title": match_title,
        "match_date": match_date,
        "competition": competition,
        "score": score,
        "opponent_name": opponent_name,
        "bkfc_season_avg": bkfc_season_avg,
        "all_opp_avg": all_opp_avg,
        "match_bkfc": match_bkfc,
        "match_opp": match_opp,
        "opp_season_avg": opp_season_avg
    }