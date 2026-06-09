import pandas as pd
import re

# Comprehensive metrics mapping transferred from advanced telemetry definitions
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

def extract_match_info(row):
    """Extract match information from a row."""
    match_title = str(row[1])
    competition = str(row[2])
    match_date = str(row[0])
    
    score = ""
    m = re.search(r"\d+:\d+", match_title)
    if m:
        score = m.group()
    
    opponent_name = match_title.split("-")[-1].replace(score, "").strip()
    
    return {
        "match_title": match_title,
        "match_date": match_date,
        "competition": competition,
        "score": score,
        "opponent_name": opponent_name,
    }

def load_available_matches(bkfc_file, opponent_file):
    """Load all available matches from both files and return selectable options."""
    bkfc_df = pd.read_excel(bkfc_file, header=None, engine="openpyxl")
    opp_df = pd.read_excel(opponent_file, header=None, engine="openpyxl")
    
    # Extract seasonal averages (rows 1-2)
    bkfc_season_avg = bkfc_df.iloc[1]
    all_opp_avg = bkfc_df.iloc[2]
    opp_season_avg = opp_df.iloc[1]
    
    # Collect all matches (starting from row 3, pairs of rows)
    matches = []
    row_idx = 3
    
    while row_idx < len(bkfc_df):
        match_bkfc = bkfc_df.iloc[row_idx]
        
        # Check if row has valid data
        if pd.isna(match_bkfc[0]) or pd.isna(match_bkfc[1]):
            break
        
        match_info = extract_match_info(match_bkfc)
        
        # Create display label for dropdown
        display_label = f"{match_info['match_date']} - {match_info['match_title']} ({match_info['competition']})"
        
        matches.append({
            "display_label": display_label,
            "row_index": row_idx,
            "match_info": match_info,
            "bkfc_season_avg": bkfc_season_avg,
            "all_opp_avg": all_opp_avg,
            "opp_season_avg": opp_season_avg,
            "bkfc_df": bkfc_df,
            "opp_df": opp_df
        })
        
        row_idx += 1
    
    return matches

def load_match_data(bkfc_file, opponent_file, match_index=0):
    """Load specific match data. If match_index is provided, load that match; otherwise load the most recent (last)."""
    bkfc_df = pd.read_excel(bkfc_file, header=None, engine="openpyxl")
    opp_df = pd.read_excel(opponent_file, header=None, engine="openpyxl")

    # Extract seasonal averages (rows 1-2)
    bkfc_season_avg = bkfc_df.iloc[1]
    all_opp_avg = bkfc_df.iloc[2]
    opp_season_avg = opp_df.iloc[1]
    
    # Collect all matches
    matches = []
    row_idx = 3
    
    while row_idx < len(bkfc_df):
        match_bkfc = bkfc_df.iloc[row_idx]
        
        if pd.isna(match_bkfc[0]) or pd.isna(match_bkfc[1]):
            break
        
        matches.append(row_idx)
        row_idx += 1
    
    # Load the selected match (default to last/most recent if index out of range)
    if match_index >= len(matches):
        match_index = len(matches) - 1
    
    match_row_idx = matches[match_index]
    match_bkfc = bkfc_df.iloc[match_row_idx]
    match_opp = bkfc_df.iloc[match_row_idx + 1] if match_row_idx + 1 < len(bkfc_df) else bkfc_df.iloc[match_row_idx]
    
    match_info = extract_match_info(match_bkfc)

    return {
        "match_title": match_info["match_title"],
        "match_date": match_info["match_date"],
        "competition": match_info["competition"],
        "score": match_info["score"],
        "opponent_name": match_info["opponent_name"],
        "bkfc_season_avg": bkfc_season_avg,
        "all_opp_avg": all_opp_avg,
        "match_bkfc": match_bkfc,
        "match_opp": match_opp,
        "opp_season_avg": opp_season_avg
    }
