# Slide Generation Capstone

This repository is an adapted and expanded version of **SlideGen**:

- Paper: [SlideGen: Collaborative Multimodal Agents for Scientific Slide Generation](https://arxiv.org/pdf/2512.04529)
- Original system focus: multi-agent scientific slide generation from research papers

This capstone fork keeps the core SlideGen pipeline, but extends it with:

- author-conditioned personalization
- retrieval-based preference profiles
- color and font conditioning
- stronger batch experimentation and evaluation tooling
- additional UI work
- newer prompt-driven planning and refinement workflows

In short, this repo is not just a reproduction of SlideGen. It is a research/engineering extension of SlideGen for **personalized scientific presentation generation**.

## What This Repo Adds

Compared with the original SlideGen release, this repo adds or expands:

- **Baseline vs personalized generation**
  - generate a normal deck
  - generate a personalized deck using an author profile
- **Retrieval-conditioned personalization**
  - build a profile from historically similar paper-deck pairs
  - condition planning on quantitative style preferences such as average bullets/slide, words/slide, image rate, formula rate, slide count, section preferences, font, and color palette
- **Template-conditioned rendering**
  - render a generated plan into a user-supplied PowerPoint template
- **Prompt-based revision / refinement**
  - revise slides or repair rendered decks with explicit user instructions
- **Evaluation**
  - personalization evaluation
  - bundle evaluation
  - layout-defect evaluation
- **Batch studies**
  - fixed-paper cohort studies
  - 5-paper and 50-paper experiments

## High-Level Pipeline

At a high level, the system works like this:

1. **PDF parsing**
   - the paper PDF is parsed into machine-usable structured content
   - this includes text blocks, headers, images, tables, and formulas
2. **Raw content extraction**
   - the parser and LLM-based content extraction produce a structured representation of the paper
3. **Asset extraction and filtering**
   - figures, tables, and formula crops are extracted
   - noisy or low-value assets are filtered
4. **Asset matching**
   - figures/tables/formulas are matched to relevant slide sections
5. **Slide planning**
   - an LLM planner chooses sections, subsections, templates, and asset placements
   - in personalized modes, the planner is conditioned on a profile JSON
6. **Rendering**
   - the slide plan is rendered into a PPTX using built-in templates or a user template
7. **Optional refinement / evaluation**
   - post-render refinement and evaluation scripts can inspect and score the deck

## Core Prompt Files

If you are trying to understand the main LLM behavior, these are the most important prompt files:

- [utils/prompt_templates/layout_agent_xin.yaml](./utils/prompt_templates/layout_agent_xin.yaml)
  - main baseline planner prompt
- [utils/prompt_templates/layout_agent_xin_retrieval.yaml](./utils/prompt_templates/layout_agent_xin_retrieval.yaml)
  - retrieval-personalized planner prompt
- [utils/prompt_templates/layout_agent_xin_strong.yaml](./utils/prompt_templates/layout_agent_xin_strong.yaml)
  - stronger planner variant
- [utils/prompt_templates/preference_distiller_target_conditioned.yaml](./utils/prompt_templates/preference_distiller_target_conditioned.yaml)
  - retrieval-style target-conditioned profile prompt
- [utils/prompt_templates/post_render_repair.yaml](./utils/prompt_templates/post_render_repair.yaml)
  - prompt used for LLM-based post-render repair
- [utils/prompt_templates/figure_match.yaml](./utils/prompt_templates/figure_match.yaml)
  - figure-to-outline matching
- [utils/prompt_templates/formula_match.yaml](./utils/prompt_templates/formula_match.yaml)
  - formula-to-outline matching
- [utils/prompt_templates/image_table_filter_agent.yaml](./utils/prompt_templates/image_table_filter_agent.yaml)
  - image/table filtering

## Repository Structure

- [SlidesAgent/](./SlidesAgent)
  - main pipeline code
- [Capstone/](./Capstone)
  - batch generation, personalization, evaluation, and experiment scripts
- [ui/](./ui)
  - Streamlit UI
- [webui/](./webui)
  - legacy web interface
- [utils/prompt_templates/](./utils/prompt_templates)
  - LLM prompts
- [SlideGen_Original/](./SlideGen_Original)
  - local copy of the original SlideGen code for comparison experiments

## Setup

### 1. Create an environment

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install --no-build-isolation \
  "python-pptx @ https://codeload.github.com/Force1ess/python-pptx/zip/dc356685d4d210a10abe1ffab3c21315cdfae63d"
pip install -r requirements.txt
```

### 2. Install LibreOffice

LibreOffice is used for PPTX/PDF rendering workflows and template rendering.

**macOS**
```bash
brew install --cask libreoffice
```

**Ubuntu**
```bash
sudo apt install libreoffice
```

### 3. Configure API access

This repo supports either direct OpenAI or Azure OpenAI.

Example `.env`:

```bash
OPENAI_API_KEY=your_openai_key
```

or Azure:

```bash
AZURE_OPENAI_API_KEY=your_azure_key
AZURE_OPENAI_BASE_URL=https://your-resource.openai.azure.com/
AZURE_API_VERSION=2025-03-01-preview
AZURE_DEPLOYMENT_NAME=gpt-5.4-nano
SLIDEGEN_USE_AZURE_OPENAI=1
```

The repo’s client utilities are in [slidegen_openai_utils.py](./slidegen_openai_utils.py).

## How To Run

All commands below assume:

```bash
cd /path/to/SlideGen
source .venv/bin/activate
```

### 1. Normal / baseline generation

This is the standard non-personalized pipeline.

```bash
python -m SlidesAgent.new_pipeline \
  --paper_path data_raw/acl18/74/74_paper.pdf \
  --paper_name acl18_74 \
  --model_name_t gpt-5.4-nano \
  --model_name_v gpt-5.4-nano \
  --outline_mode high_level \
  --formula_mode 1 \
  --output_dir outputs/example_baseline
```

Useful options:

- `--outline_mode high_level`
  - more presentation-like summary structure
- `--outline_mode technical`
  - preserves paper subsection structure more closely
- `--formula_mode 1|2|3`
  - different formula-handling modes

### 2. Personalized generation

This repo’s supported personalization path is **retrieval-conditioned personalization**.

First build the retrieval profile:

```bash
python Capstone/retrieval_profile_pilot.py \
  --author-id hinrich_sch_tze \
  --target-paper-id acl18:74 \
  --model gpt-5.4-nano \
  --output-dir outputs/retrieval_profiles
```

Then run the personalized deck:

```bash
python -m SlidesAgent.new_pipeline \
  --paper_path data_raw/acl18/74/74_paper.pdf \
  --paper_name acl18_74 \
  --model_name_t gpt-5.4-nano \
  --model_name_v gpt-5.4-nano \
  --outline_mode high_level \
  --formula_mode 1 \
  --output_dir outputs/example_personalized_retrieval \
  --use_author_preferences \
  --author_profile_path outputs/retrieval_profiles/hinrich_sch_tze.acl18_74.retrieval.json \
  --personalization_mode retrieval
```

### 3. Use a user PowerPoint template

There are two supported ways to do this.

#### 3a. Run SlideGen normally, then bind to a user template

Generate the plan/deck first:

```bash
python -m SlidesAgent.new_pipeline_logtime \
  --paper_path data_raw/acl18/74/74_paper.pdf \
  --paper_name acl18_74 \
  --model_name_t gpt-5.4-nano \
  --model_name_v gpt-5.4-nano \
  --outline_mode high_level \
  --formula_mode 1
```

Then bind that plan into a user template:

```bash
python -m scripts.render_with_user_template \
  --paper_name acl18_74 \
  --template_path "/absolute/path/to/template_deck.pptx" \
  --template_label mytemplate \
  --model_name_t gpt-5.4-nano \
  --model_name_v gpt-5.4-nano \
  --outline_mode high_level
```

#### 3b. Use `new_pipeline_logtime.py` with `--template_path`

This path runs the normal generation flow and then invokes the layout binder automatically:

```bash
python -m SlidesAgent.new_pipeline_logtime \
  --paper_path data_raw/acl18/74/74_paper.pdf \
  --paper_name acl18_74 \
  --model_name_t gpt-5.4-nano \
  --model_name_v gpt-5.4-nano \
  --outline_mode high_level \
  --formula_mode 1 \
  --template_path "/absolute/path/to/template_deck.pptx"
```

### 4. Run with the UI

#### Streamlit UI

The Streamlit app lives at [ui/app.py](./ui/app.py):

```bash
streamlit run ui/app.py
```

What it currently provides:

- a styled UI shell for the project
- upload/generation-oriented interaction
- a slide revision sandbox
- post-generation bullet rewriting with an LLM

Important note:
- the current Streamlit UI is useful as an interaction/demo surface, but the CLI pipeline is still the main production path for full experiments

#### Legacy web UI

If you want to use the older web UI:

Backend:

```bash
cd webui/backend
uvicorn main:app --host 0.0.0.0 --port 8000 --reload
```

Frontend:

```bash
cd webui/frontend
npm install
npm run dev
```

### 5. Use LLM prompting for revision / repair

There are two main prompt-driven workflows in this repo.

#### 5a. Slide-level prompting in the Streamlit UI

The UI supports free-form post-generation slide revision instructions such as:

- “make this slide more technical”
- “shorten to 3 bullets”
- “add a sub-bullet explaining the formula”

This is implemented in [ui/app.py](./ui/app.py).

#### 5b. Prompt-based post-render refinement from the CLI

For rendered decks, you can inspect and repair a generated run with the post-render refinement agent:

Inspect a generated run:

```bash
python -m SlidesAgent.post_render_refinement local inspect \
  --run-dir contents/acl18_74_high_level
```

Then perform repair:

```bash
python -m SlidesAgent.post_render_refinement local repair \
  --run-dir contents/acl18_74_high_level \
  --repair-mode llm \
  --model gpt-5.4-nano
```

This workflow uses:

- [SlidesAgent/post_render_refinement.py](./SlidesAgent/post_render_refinement.py)
- [utils/prompt_templates/post_render_repair.yaml](./utils/prompt_templates/post_render_repair.yaml)

## Outputs

Generated artifacts typically appear under:

- `contents/<paper_name>/`
- `<model_t>_<model_v>_images_and_tables/<paper_name>/`
- custom experiment folders under `outputs/`

Depending on the run, you may see:

- PPTX decks
- slide plans
- raw content JSON
- figure match JSON
- formula match JSON
- rendered slide images
- personalization traces
- cost summaries
- evaluation outputs

## Batch Experiments

The main cohort-study script is:

- [Capstone/run_batch_experiment.py](./Capstone/run_batch_experiment.py)

This is used for baseline vs personalized deck generation over multiple papers.

Other useful experiment scripts:

- [Capstone/run_all_bundle_eval_5papers.py](./Capstone/run_all_bundle_eval_5papers.py)
- [Capstone/run_bundle_eval_from_manifest.py](./Capstone/run_bundle_eval_from_manifest.py)
- [Capstone/summarize_generation_costs.py](./Capstone/summarize_generation_costs.py)

## Evaluation

The capstone branch includes evaluation beyond the original SlideGen release.

Main evaluation entry points:

- [Capstone/evaluate_pptx_bundle.py](./Capstone/evaluate_pptx_bundle.py)
- [Capstone/slidetailor_eval/run_slidetailor_eval.py](./Capstone/slidetailor_eval/run_slidetailor_eval.py)
- [Capstone/slidetailor_eval/evaluate_layout_correctness.py](./Capstone/slidetailor_eval/evaluate_layout_correctness.py)

Examples of metrics used in recent experiments:

- core coverage
- geometry-aware density
- visual appeal
- paper faithfulness
- logical flow
- layout defects / defect rate
- personalization alignment metrics

## Original SlideGen Comparison

This repo also contains a local copy of the original system in:

- [SlideGen_Original/](./SlideGen_Original)

That copy is used for direct experimental comparisons between:

- **SlideGen Original**
- **SlideGen Baseline** in this repo
- **SlideGen Personalized** in this repo

## Citation

If you use this repository in academic work, please cite the original SlideGen paper as the upstream system:

```bibtex
@article{liang2025slidegen,
  title={SlideGen: Collaborative Multimodal Agents for Scientific Slide Generation},
  author={Liang, Xin and Zhang, Xiang and Xu, Yiwei and Sun, Siqi and You, Chenyu},
  journal={arXiv preprint arXiv:2512.04529},
  year={2025}
}
```

If you describe this repository, it is best referred to as:

- an adapted SlideGen codebase
- a SlideGen-derived personalization and evaluation extension

## Acknowledgments

This project builds on SlideGen and related open-source tools. In particular:

- [SlideGen](https://arxiv.org/pdf/2512.04529)
- [Docling](https://www.docling.ai/)
- [Marker](https://github.com/datalab-to/marker)
- [python-pptx](https://github.com/scanny/python-pptx)
- [Paper2Poster](https://github.com/Paper2Poster/Paper2Poster)
