import streamlit as st

from data_parser import load_match_data, build_match_data
from report_generator import generate_report

st.set_page_config(
    page_title="BKFC Technical Match Intelligence",
    page_icon="⚽",
    layout="wide"
)

st.title("⚽ Brooklyn FC")
st.subheader("Automated Match Analysis & Scouting Report Generator")
st.markdown("---")


# ─────────────────────────────────────────────
# FILE UPLOADS
# ─────────────────────────────────────────────
col1, col2 = st.columns(2)

with col1:
    st.markdown("### 📥 BKFC Data")
    bkfc_file = st.file_uploader("Upload BKFC Wyscout File (.xlsx)", type=["xlsx"])

with col2:
    st.markdown("### 📥 Opponent Data")
    opp_file = st.file_uploader("Upload Opponent Wyscout File (.xlsx)", type=["xlsx"])


# ─────────────────────────────────────────────
# MAIN PIPELINE
# ─────────────────────────────────────────────
if bkfc_file and opp_file:

    try:
        parsed = load_match_data(bkfc_file, opp_file)

        # ── MATCH DROPDOWN ─────────────────────
        match_options = {
            f"{m['date']} | {m['opponent']} | {m['competition']}": m
            for m in parsed["matches"]
        }

        selected_label = st.selectbox(
            "Select Match to Generate Report",
            list(match_options.keys())
        )

        parsed["selected_match"] = match_options[selected_label]

        # convert to report-ready structure
        data = build_match_data(parsed)

        # ── MATCH INFO DISPLAY ──────────────────
        st.success(f"Loaded Match: {data['match_title']}")

        colA, colB, colC = st.columns(3)
        colA.metric("Opponent", data["opponent_name"])
        colB.metric("Date", data["match_date"])
        colC.metric("Score", data["score"])

        st.markdown("---")

        # quick sanity check
        st.write("Goals (BKFC):", data["match_bkfc"]["Goals"])

        confirm = st.checkbox("Confirm selection before generating report")

        if confirm:

            if st.button("Generate Match Report", use_container_width=True):

                with st.spinner("Building tactical report..."):
                    report_stream = generate_report(data)

                filename = f"BKFC_vs_{data['opponent_name'].replace(' ', '_')}.pptx"

                st.download_button(
                    "Download PowerPoint Report",
                    data=report_stream,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    use_container_width=True
                )

    except Exception as e:
        st.error(f"Pipeline error: {str(e)}")
