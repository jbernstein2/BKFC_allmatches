def generate_insights(stats_list):
    insights = []
    for s in stats_list:
        if s["season"] == 0:
            continue
        
        diff = (s["match"] - s["season"]) / s["season"]
        
        if diff > 0.15:
            insights.append(f"{s['label']} was significantly higher (+{diff:.1%}) than our season baseline.")
        elif diff < -0.15:
            insights.append(f"{s['label']} dropped sharply ({diff:.1%}) compared to our season baseline.")
            
    return insights[:6]  # Visual safeguard to prevent text boxes from overflowing slides
