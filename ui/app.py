import streamlit as st
import json
import copy
import difflib
import os
import time

st.set_page_config(
    page_title="SlideGen · Personalized",
    page_icon="🎯",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ── Custom CSS ─────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Serif+Display:ital@0;1&family=DM+Mono:wght@300;400;500&family=DM+Sans:wght@300;400;500;600&display=swap');

html, body, [class*="css"] { font-family: 'DM Sans', sans-serif; }
.main { background-color: #0e0e0f; color: #f0ede8; }
.block-container { padding: 2.5rem 3rem 4rem 3rem; max-width: 1200px; }

.site-title {
    font-family: 'DM Serif Display', serif;
    font-size: 2.6rem;
    color: #f0ede8;
    letter-spacing: -0.02em;
    margin: 0;
    line-height: 1;
}

.site-badge {
    font-family: 'DM Mono', monospace;
    font-size: 0.65rem;
    font-weight: 500;
    letter-spacing: 0.12em;
    text-transform: uppercase;
    color: #c8b89a;
    background: rgba(200, 184, 154, 0.12);
    border: 1px solid rgba(200, 184, 154, 0.25);
    padding: 0.2rem 0.6rem;
    border-radius: 2px;
    margin-left: 0.75rem;
    position: relative;
    top: -0.4rem;
}

.section-label {
    font-family: 'DM Mono', monospace;
    font-size: 0.65rem;
    font-weight: 500;
    letter-spacing: 0.15em;
    text-transform: uppercase;
    color: #c8b89a;
    margin-bottom: 0.75rem;
    display: block;
}

.card {
    background: #161617;
    border: 1px solid #262626;
    border-radius: 8px;
    padding: 1.5rem;
    margin-bottom: 1.25rem;
}

.card-accent { border-left: 3px solid #c8b89a; }

.bullet-item {
    font-size: 0.875rem;
    color: #a09d99;
    padding: 0.35rem 0;
    border-bottom: 1px solid #1e1e1e;
    line-height: 1.5;
}

.bullet-item:last-child { border-bottom: none; }

.sub-bullet {
    font-size: 0.8rem;
    color: #6b6860;
    padding: 0.2rem 0 0.2rem 1.25rem;
    line-height: 1.4;
}

.slide-label {
    font-family: 'DM Mono', monospace;
    font-size: 0.7rem;
    letter-spacing: 0.08em;
    text-transform: uppercase;
    color: #6b6860;
    margin-bottom: 0.4rem;
}

.status-ok {
    color: #7a9e8a;
    font-size: 0.8rem;
    font-family: 'DM Mono', monospace;
}

.thin-divider {
    border: none;
    border-top: 1px solid #1e1e1e;
    margin: 2rem 0;
}

.col-header {
    font-family: 'DM Mono', monospace;
    font-size: 0.65rem;
    letter-spacing: 0.12em;
    text-transform: uppercase;
    color: #6b6860;
    margin-bottom: 0.75rem;
}

.helper-text {
    font-size: 0.8rem;
    color: #6b6860;
    line-height: 1.5;
    margin-bottom: 0.5rem;
}

.placeholder-note {
    font-family: 'DM Mono', monospace;
    font-size: 0.65rem;
    color: #3d3d3d;
    letter-spacing: 0.08em;
}

/* Streamlit overrides */
.stTextInput > div > div > input {
    background: #161617 !important;
    border: 1px solid #262626 !important;
    border-radius: 6px !important;
    color: #f0ede8 !important;
    font-family: 'DM Mono', monospace !important;
    font-size: 0.85rem !important;
}
.stTextInput > div > div > input:focus {
    border-color: #c8b89a !important;
    box-shadow: 0 0 0 1px rgba(200, 184, 154, 0.3) !important;
}
.stTextArea > div > div > textarea {
    background: #161617 !important;
    border: 1px solid #262626 !important;
    border-radius: 6px !important;
    color: #f0ede8 !important;
    font-family: 'DM Sans', sans-serif !important;
    font-size: 0.85rem !important;
}
.stFileUploader {
    background: #161617 !important;
    border: 1px dashed #262626 !important;
    border-radius: 8px !important;
}
.stButton > button {
    background: #c8b89a !important;
    color: #0e0e0f !important;
    border: none !important;
    border-radius: 6px !important;
    font-family: 'DM Mono', monospace !important;
    font-size: 0.75rem !important;
    font-weight: 500 !important;
    letter-spacing: 0.08em !important;
    text-transform: uppercase !important;
    padding: 0.6rem 1.5rem !important;
    transition: all 0.15s ease !important;
}
.stButton > button:hover {
    background: #d4c6ad !important;
    transform: translateY(-1px) !important;
}
.stSelectbox > div > div {
    background: #161617 !important;
    border: 1px solid #262626 !important;
    color: #f0ede8 !important;
}
.stTabs [data-baseweb="tab-list"] {
    background: transparent !important;
    border-bottom: 1px solid #262626 !important;
    gap: 0 !important;
}
.stTabs [data-baseweb="tab"] {
    font-family: 'DM Mono', monospace !important;
    font-size: 0.7rem !important;
    letter-spacing: 0.1em !important;
    text-transform: uppercase !important;
    color: #6b6860 !important;
    background: transparent !important;
    border: none !important;
    padding: 0.75rem 1.25rem !important;
}
.stTabs [aria-selected="true"] {
    color: #c8b89a !important;
    border-bottom: 2px solid #c8b89a !important;
}
</style>
""", unsafe_allow_html=True)


# ── Helpers ────────────────────────────────────────────────────────────────────
def find_slide(slides, section, subsection):
    best_score = 0.0
    best_idx = None
    for i, slide in enumerate(slides):
        sec_score = difflib.SequenceMatcher(
            None, slide["section"].lower(), section.lower()
        ).ratio()
        sub_score = difflib.SequenceMatcher(
            None, slide["subsection"].lower(), subsection.lower()
        ).ratio()
        combined = (sec_score + sub_score) / 2
        if combined > best_score:
            best_score = combined
            best_idx = i
    if best_idx is not None and best_score >= 0.55:
        return best_idx, slides[best_idx]
    return None, None


def revise_bullets(slide, instruction):
    from openai import AzureOpenAI
    client = AzureOpenAI(
        api_key=os.environ.get("AZURE_OPENAI_API_KEY", ""),
        azure_endpoint=os.environ.get("AZURE_OPENAI_ENDPOINT", ""),
        api_version="2024-02-15-preview",
    )

    system_prompt = """You are a slide revision assistant.
You will receive a slide's current bullet points and a user instruction.
Your ONLY task: return a revised bullets array as valid JSON.
Each bullet must follow this exact schema:
[{"text": "<string>", "sub": ["<string>", ...]}, ...]
Rules:
- Keep bullets factually faithful to the original content
- Apply the instruction to tone, depth, length, or structure
- Return ONLY the JSON array, no explanation, no markdown fences
- Top-level bullets: max 20 words
- Sub-bullets: max 25 words
- Max 6 top-level bullets"""

    user_prompt = f"""Current bullets:
{json.dumps(slide["bullets"], indent=2)}

Slide context:
Section: {slide["section"]}
Subsection: {slide["subsection"]}

User instruction: {instruction}

Return the revised bullets JSON array only."""

    response = client.chat.completions.create(
        model="gpt-4.1-mini",
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt},
        ],
        temperature=0.3,
        max_tokens=800,
    )

    raw = response.choices[0].message.content.strip()
    if raw.startswith("```"):
        raw = raw.split("```")[1]
        if raw.startswith("json"):
            raw = raw[4:]
    return json.loads(raw.strip())


def render_bullets(bullets):
    for b in bullets:
        st.markdown(
            f'<div class="bullet-item">• {b["text"]}</div>',
            unsafe_allow_html=True
        )
        for s in b.get("sub", []):
            st.markdown(
                f'<div class="sub-bullet">↳ {s}</div>',
                unsafe_allow_html=True
            )


# ── Sample plan (sandbox) ──────────────────────────────────────────────────────
SAMPLE_PLAN = {
    "slides": [
        {
            "section": "Motivation and Problem Statement",
            "subsection": "Importance of Constituency Parsing",
            "template_id": "T2_ImageRight",
            "bullets": [
                {"text": "Constituency parsing is fundamental in natural language processing applications.", "sub": []},
                {"text": "Supports relation extraction, paraphrase detection, natural language inference, and machine translation.", "sub": []},
                {"text": "Fast and accurate parsing remains a long-standing challenge.", "sub": []}
            ],
            "images": ["image_1.png"], "tables": [], "formulas": []
        },
        {
            "section": "Motivation and Problem Statement",
            "subsection": "Limitations of Existing Methods",
            "template_id": "T1_TextOnly",
            "bullets": [
                {"text": "Transition-based models suffer from compounding errors due to sequential local decisions.", "sub": []},
                {"text": "Chart-based models have high computational costs from complex structured inference.", "sub": []},
                {"text": "These limitations motivate the need for a more efficient and robust parsing approach.", "sub": []}
            ],
            "images": [], "tables": [], "formulas": []
        },
        {
            "section": "Empirical Evaluation and Results",
            "subsection": "Performance on Penn Treebank",
            "template_id": "T4_ImageTop",
            "bullets": [
                {"text": "Achieves labeled F1 score of 91.8 on Penn Treebank test set.", "sub": []},
                {"text": "Competitive with or surpasses recent single-model discriminative parsers.", "sub": []},
                {"text": "Uses standard splits and preprocessing; POS tags predicted externally.", "sub": []}
            ],
            "images": [], "tables": ["table_2.png"], "formulas": []
        },
        {
            "section": "Conclusions and Implications",
            "subsection": "Summary of Contributions",
            "template_id": "T1_TextOnly",
            "bullets": [
                {"text": "Introduces a fully parallel constituency parsing method based on syntactic distances.", "sub": []},
                {"text": "Achieves competitive accuracy with faster parsing speeds and simpler training.", "sub": []},
                {"text": "Outperforms existing methods in efficiency and robustness.", "sub": []}
            ],
            "images": [], "tables": [], "formulas": []
        }
    ]
}


# ── Session state ──────────────────────────────────────────────────────────────
if "plan" not in st.session_state:
    st.session_state["plan"] = copy.deepcopy(SAMPLE_PLAN)
if "deck_generated" not in st.session_state:
    st.session_state["deck_generated"] = False
if "revised_bullets" not in st.session_state:
    st.session_state["revised_bullets"] = None
if "original_bullets_snapshot" not in st.session_state:
    st.session_state["original_bullets_snapshot"] = None


# ── Header ─────────────────────────────────────────────────────────────────────
st.markdown("""
<h1 class="site-title">
    [Project Name] <span class="site-badge">Personalized Slide Generation</span>
</h1>
<br>
""", unsafe_allow_html=True)

# ── Tabs ───────────────────────────────────────────────────────────────────────
tab_generate, tab_revise = st.tabs(["Generate", "Revise"])


# ════════════════════════════════════════════════════════════════════════════════
# TAB 1 — GENERATE
# ════════════════════════════════════════════════════════════════════════════════
with tab_generate:
    col_left, col_right = st.columns([1, 1.6], gap="large")

    # ── Left column ───────────────────────────────────────────────────────────
    with col_left:
        st.markdown('<span class="section-label">Paper</span>', unsafe_allow_html=True)
        uploaded_file = st.file_uploader(
            "Upload PDF",
            type=["pdf"],
            label_visibility="collapsed"
        )
        if uploaded_file:
            st.markdown(
                f'<div class="status-ok">✓ {uploaded_file.name}</div>',
                unsafe_allow_html=True
            )

        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown('<span class="section-label">Outline Mode</span>', unsafe_allow_html=True)
        st.markdown(
            '<p class="helper-text">High Level produces a compact narrative deck. Technical preserves the paper\'s section structure closely.</p>',
            unsafe_allow_html=True
        )
        outline_mode = st.selectbox(
            "Outline Mode",
            ["high_level", "technical"],
            format_func=lambda x: "High Level" if x == "high_level" else "Technical",
            label_visibility="collapsed"
        )

        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown(
            '<span class="section-label">Pre-Generation Instructions <span style="color:#3d3d3d">(optional)</span></span>',
            unsafe_allow_html=True
        )
        st.markdown(
            '<p class="helper-text">Shape the whole deck before it is generated. Applied at generation time.</p>',
            unsafe_allow_html=True
        )
        pre_gen_instructions = st.text_area(
            "Pre-generation instructions",
            placeholder="e.g. Keep it under 10 slides, focus on the results section, make it highly technical...",
            label_visibility="collapsed",
            height=100,
        )

    # ── Right column ──────────────────────────────────────────────────────────
    with col_right:
        st.markdown('<span class="section-label">About This Tool</span>', unsafe_allow_html=True)
        st.markdown("""
        <div class="card">
            <p class="helper-text" style="margin-bottom:0.75rem;">
                This system generates a personalized slide deck from your paper,
                conditioned on your presentation history and any instructions you provide.
            </p>
            <p class="helper-text" style="margin-bottom:0.75rem;">
                After generation, use the <strong style="color:#c8b89a">Revise</strong> tab
                to refine individual slides with free-form instructions.
            </p>
            <p class="placeholder-note">↳ Pipeline connection coming soon</p>
        </div>
        """, unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)
        generate_btn = st.button("Generate Slides", use_container_width=True)

        if generate_btn:
            if not uploaded_file:
                st.warning("Please upload a PDF first.")
            else:
                with st.spinner("Generating your personalized slides..."):
                    time.sleep(2)
                st.session_state["deck_generated"] = True
                st.session_state["plan"] = copy.deepcopy(SAMPLE_PLAN)
                st.session_state["revised_bullets"] = None
                st.session_state["original_bullets_snapshot"] = None
                st.success("Deck generated. Head to the Revise tab to refine individual slides.")


# ════════════════════════════════════════════════════════════════════════════════
# TAB 2 — REVISE
# ════════════════════════════════════════════════════════════════════════════════
with tab_revise:
    if not st.session_state["deck_generated"]:
        st.markdown("""
        <div class="card" style="text-align:center; padding:3rem; margin-top:2rem;">
            <div style="font-family:'DM Serif Display',serif; font-size:1.4rem; color:#3d3d3d; margin-bottom:0.5rem;">
                No deck generated yet
            </div>
            <div class="helper-text" style="text-align:center;">
                Go to the Generate tab, upload a paper, and hit Generate first.
            </div>
        </div>
        """, unsafe_allow_html=True)
    else:
        plan = st.session_state["plan"]
        slides = plan["slides"]

        col_left, col_right = st.columns([1, 1.4], gap="large")

        # ── Left: slide selector + current bullets ────────────────────────────
        with col_left:
            st.markdown('<span class="section-label">Select a Slide</span>', unsafe_allow_html=True)
            slide_labels = [
                f"{s['section']} → {s['subsection']}"
                for s in slides
            ]
            selected_label = st.selectbox(
                "Slide",
                slide_labels,
                label_visibility="collapsed"
            )
            selected_idx = slide_labels.index(selected_label)
            selected_slide = slides[selected_idx]

            st.markdown("<br>", unsafe_allow_html=True)
            st.markdown('<span class="section-label">Current Bullets</span>', unsafe_allow_html=True)
            st.markdown('<div class="card">', unsafe_allow_html=True)
            st.markdown(
                f'<div class="slide-label">{selected_slide["template_id"]}</div>',
                unsafe_allow_html=True
            )
            render_bullets(selected_slide["bullets"])
            st.markdown('</div>', unsafe_allow_html=True)

        # ── Right: instruction + before/after ─────────────────────────────────
        with col_right:
            st.markdown('<span class="section-label">Post-Generation Instruction</span>', unsafe_allow_html=True)
            st.markdown(
                '<p class="helper-text">Revise this specific slide with a free-form instruction.</p>',
                unsafe_allow_html=True
            )
            instruction = st.text_area(
                "Instruction",
                placeholder="e.g. Make this less technical, shorten to 2 bullets, add sub-bullets with more detail...",
                label_visibility="collapsed",
                height=100,
            )

            revise_btn = st.button("Revise Slide", use_container_width=True)

            if revise_btn:
                if not instruction.strip():
                    st.warning("Please enter an instruction first.")
                else:
                    with st.spinner("Revising..."):
                        try:
                            snapshot = copy.deepcopy(selected_slide["bullets"])
                            revised = revise_bullets(selected_slide, instruction)

                            st.session_state["plan"]["slides"][selected_idx]["bullets"] = revised
                            st.session_state["revised_bullets"] = revised
                            st.session_state["original_bullets_snapshot"] = snapshot

                        except Exception as e:
                            st.error(f"Something went wrong: {e}")

            # ── Before / after ─────────────────────────────────────────────────
            if st.session_state["revised_bullets"] is not None:
                st.markdown('<hr class="thin-divider">', unsafe_allow_html=True)

                bc, ac = st.columns(2, gap="medium")

                with bc:
                    st.markdown('<div class="col-header">Before</div>', unsafe_allow_html=True)
                    st.markdown('<div class="card">', unsafe_allow_html=True)
                    render_bullets(st.session_state["original_bullets_snapshot"])
                    st.markdown('</div>', unsafe_allow_html=True)

                with ac:
                    st.markdown('<div class="col-header">After</div>', unsafe_allow_html=True)
                    st.markdown('<div class="card card-accent">', unsafe_allow_html=True)
                    render_bullets(st.session_state["revised_bullets"])
                    st.markdown('</div>', unsafe_allow_html=True)

                st.success("Slide revised. Select another slide to keep refining, or download below.")

        # ── Download ───────────────────────────────────────────────────────────
        st.markdown('<hr class="thin-divider">', unsafe_allow_html=True)
        st.markdown('<span class="section-label">Export</span>', unsafe_allow_html=True)
        st.markdown(
            '<p class="helper-text">Download the full revised slide plan. Re-render through the pipeline to produce an updated PPTX.</p>',
            unsafe_allow_html=True
        )
        st.download_button(
            label="Download Revised Plan",
            data=json.dumps(st.session_state["plan"], indent=4, ensure_ascii=False),
            file_name="revised_plan.json",
            mime="application/json",
        )