import streamlit as st
from mem0_store import get_user_profile, get_or_create_profile_path, update_user_profile

st.set_page_config(page_title="Slide Generation Capstone", layout="centered")

st.title("Personalized Slide Generator")
st.write("Upload a research paper and we'll generate a personalized slide deck for you.")

# --- User Identity ---
st.subheader("Who are you?")
user_id = st.text_input("Enter your name or user ID", placeholder="e.g. kate_higgins")

profile_path = None

if user_id:
    profile = get_user_profile(user_id)
    
    if profile and profile['count'] > 0:
        st.success(f"Welcome back, {user_id}!")
        
        # Get profile path for pipeline
        profile_path = get_or_create_profile_path(user_id, profile)
        
        # Display preferences
        st.subheader("Your Learned Preferences")
        st.caption("Based on your previous presentations. These are applied automatically.")
        
        for memory in profile['results']:
            st.write(f"• {memory['memory']}")
    
    else:
        st.info(f"Welcome, {user_id}! No preferences stored yet. Your profile will be built after your first session.")

st.divider()

# --- PDF Upload ---
st.subheader("Upload your paper")
uploaded_file = st.file_uploader("Upload your PDF", type=["pdf"])

if uploaded_file is not None:
    st.success(f"Got it: {uploaded_file.name}")
    
    if st.button("Generate Slides"):
        with st.spinner("Generating your personalized slides..."):
            import time
            time.sleep(3)
            # Pipeline call will go here
            # subprocess.run([
            #     "python", "-m", "SlidesAgent.new_pipeline",
            #     "--paper_path", saved_pdf_path,
            #     "--use_author_preferences",
            #     "--author_id", user_id,
            #     "--author_profile_path", profile_path
            # ])
        st.success("Slides generated!")
        st.info("(Pipeline not yet connected — coming soon)")

        # --- Feedback ---
        st.subheader("How were your slides?")
        feedback = st.text_area("Any preferences to update for next time?", 
                                placeholder="e.g. I prefer fewer slides, more visuals")
        if st.button("Save Feedback"):
            if feedback:
                update_user_profile(user_id, feedback)
                st.success("Preferences updated for your next session!")

# --- Evaluation Results ---
st.divider()
st.subheader("Past Evaluation Results")

import json
from pathlib import Path

bundle_root = Path("Capstone/evaluations/deck_bundles")

if user_id:
    summaries = list(bundle_root.rglob("summary.json"))
    
    if summaries:
        rows = []
        for s in summaries:
            data = json.loads(s.read_text())
            metrics = data.get("metrics", {})
            rows.append({
                "Paper": data.get("title", data.get("paper_id", s.parent.name)),
                "ID": data.get("paper_id", ""),
                "Type": "Personalized" if "personalized" in s.parent.name else "Baseline",
                "Coverage IoU": metrics.get("core_coverage", {}).get("topic_iou"),
                "GAD": metrics.get("geometry_aware_density", {}).get("gad_geom"),
                "Aesthetic": metrics.get("slidetailor_aesthetic_quality", {}).get("deck_score"),
                "Content": metrics.get("slidetailor_content_informativeness", {}).get("deck_score"),
            })
        st.dataframe(rows)
    else:
        st.info("No evaluation results found yet.")
else:
    st.info("Enter your user ID above to see evaluation results.")
