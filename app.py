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


col1, col2 = st.columns(2)

with col1:
    st.markdown("### 📥 BKFC Data")
    bkfc_file = st.file_uploader("Upload BKFC Wyscout Season Database (.xlsx)", type=["xlsx"])

with col2:
    st.markdown("### 📥 Opponent Data")
    opp_file = st.file_uploader("Upload Opponent Wyscout Season Database (.xlsx)", type=["xlsx"])


if bkfc_file and opp_file:

    try:
        parsed = load_match_data(bkfc_file, opp_file)

        # ── MATCH DROPDOWN ───────────────────────────────
        match_options = {
            f"{m['match_date']} | {m['opponent_name']} | {m['competition']}": m
            for m in parsed["matches"]
        }

        selected_label = st.selectbox(
            "Select Match to Generate Report",
            list(match_options.keys())
        )

        parsed["selected_match"] = match_options[selected_label]

        data = build_match_data(parsed, opp_file)

        st.success(f"Loaded: {data['match_title']}")

        # ── METADATA ─────────────────────────────────────
        colA, colB, colC = st.columns(3)
        colA.metric("Competition", data["competition"])
        colB.metric("Date", data["match_date"])
        colC.metric("Score", data["score"])

        st.markdown("---")

        confirm = st.checkbox("Confirm match selection")

        if confirm:
            if st.button("Generate Report Deck", use_container_width=True):

                with st.spinner("Building tactical report..."):
                    report_stream = generate_report(data)

                    clean_name = data["opponent_name"].replace(" ", "_")
                    filename = f"BKFC_vs_{clean_name}_report.pptx"

                st.success("Report generated successfully!")
                st.download_button(
                    "Download PPTX",
                    data=report_stream,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    use_container_width=True
                )

    except Exception as e:
        st.error(f"Pipeline error: {str(e)}")
