import io
import matplotlib.pyplot as plt
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

from branding import COLORS, get_rgb
from insights import generate_insights


# ─────────────────────────────────────────────
# SAFE PRIMITIVE
# ─────────────────────────────────────────────
def safe_float(val):
    try:
        return float(val)
    except:
        return 0.0


# ─────────────────────────────────────────────
# PRESENTATION INIT
# ─────────────────────────────────────────────
def create_prs():
    prs = Presentation()
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)
    return prs


# ─────────────────────────────────────────────
# CHART ENGINE
# ─────────────────────────────────────────────
def generate_chart(values, labels, colors, title):

    fig, ax = plt.subplots(figsize=(5.5, 4), dpi=200)

    ax.bar(labels, values, color=colors)
    ax.set_title(title)

    buf = io.BytesIO()
    plt.savefig(buf, format="png")
    buf.seek(0)
    plt.close(fig)

    return buf


# ─────────────────────────────────────────────
# TITLE SLIDE
# ─────────────────────────────────────────────
def add_title_slide(prs, data):

    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = get_rgb(COLORS["BLACK"])

    title = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(11), Inches(1))
    tf = title.text_frame
    tf.text = "BROOKLYN FC MATCH REPORT"
    tf.paragraphs[0].font.size = Pt(40)
    tf.paragraphs[0].font.bold = True
    tf.paragraphs[0].font.color.rgb = get_rgb(COLORS["GOLD"])


# ─────────────────────────────────────────────
# SUMMARY SLIDE
# ─────────────────────────────────────────────
def add_summary_slide(prs, data):

    slide = prs.slides.add_slide(prs.slide_layouts[6])

    stats = data["match_bkfc"]

    table = slide.shapes.add_table(rows=4, cols=3,
                                   left=Inches(1),
                                   top=Inches(1.5),
                                   width=Inches(11),
                                   height=Inches(3)).table

    table.cell(0, 0).text = "Metric"
    table.cell(0, 1).text = "BKFC"
    table.cell(0, 2).text = data["opponent_name"]

    keys = ["Goals", "xG", "Shots", "Passes"]

    for i, k in enumerate(keys, start=1):
        table.cell(i, 0).text = k
        table.cell(i, 1).text = str(stats[k])
        table.cell(i, 2).text = str(safe_float(stats[k]))  # placeholder opponent fallback


# ─────────────────────────────────────────────
# INSIGHTS SLIDE
# ─────────────────────────────────────────────
def add_insights_slide(prs, data):

    slide = prs.slides.add_slide(prs.slide_layouts[6])

    stats = data["match_bkfc"]

    season = data["bkfc_season_avg"]

    stats_list = []
    for k in stats:
        stats_list.append({
            "label": k,
            "match": stats[k],
            "season": season.get(k, 0)
        })

    insights = generate_insights(stats_list)

    y = Inches(1.5)

    for text in insights:

        box = slide.shapes.add_textbox(Inches(1), y, Inches(11), Inches(0.5))
        tf = box.text_frame
        tf.text = text
        tf.paragraphs[0].font.size = Pt(14)

        y += Inches(0.6)


# ─────────────────────────────────────────────
# STAT SLIDE (SIMPLE SAFE VERSION)
# ─────────────────────────────────────────────
def add_stat_slide(prs, data, label, key):

    slide = prs.slides.add_slide(prs.slide_layouts[6])

    stats = data["match_bkfc"]
    season = data["bkfc_season_avg"]

    match_val = safe_float(stats.get(label, 0))
    season_val = safe_float(season.get(label, 0))

    title = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(11), Inches(1))
    tf = title.text_frame
    tf.text = label.upper()
    tf.paragraphs[0].font.size = Pt(24)

    content = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(11), Inches(1))
    tf2 = content.text_frame
    tf2.text = f"Match: {match_val} | Season: {season_val}"


# ─────────────────────────────────────────────
# MAIN GENERATOR
# ─────────────────────────────────────────────
def generate_report(data):

    prs = create_prs()

    add_title_slide(prs, data)
    add_summary_slide(prs, data)
    add_insights_slide(prs, data)

    # optional deeper stat slides
    for k in list(data["match_bkfc"].keys())[:8]:
        add_stat_slide(prs, data, k, k)

    output = io.BytesIO()
    prs.save(output)
    output.seek(0)

    return output
