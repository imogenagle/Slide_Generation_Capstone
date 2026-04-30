# Changelog

This file tracks changes made on the `dev` branch of the capstone project.

Use this as a shared running log:

- add only changes that are committed or ready to merge
- write in plain language
- group entries by date
- include the files or subsystem touched when that helps

## 2026-04-27

### Added
- Added a Streamlit UI layer in `ui/`.
- Added a Mem0 integration helper in `ui/mem0_store.py`.
- Added `.env.example` support for the new UI path.
- Added `Capstone/slidetailor_eval/` as a dedicated home for SlideTailor-derived evaluation code and outputs.

### Notes
- This work adds an alternate app-facing interface layer on top of the pipeline. It does not replace the `SlidesAgent` CLI pipeline.

## 2026-04-24

### Added
- Added author-preference distillation and runtime integration on the `dev` branch.
- Added evaluation tooling in `Capstone/`:
  - `evaluate_core_coverage.py`
  - `evaluate_generated_decks.py`
  - `generate_random_decks.py`
  - `summarize_core_coverage.py`
  - `summarize_evaluation_summaries.py`

### Changed
- Updated `Capstone/preference_distill.py` to support a usable preference-distillation flow.
- Updated the main pipeline entry points:
  - `SlidesAgent/new_pipeline.py`
  - `SlidesAgent/new_pipeline_logtime.py`
- Updated personalization-aware runtime components:
  - `SlidesAgent/parse_raw.py`
  - `SlidesAgent/layout_agent_xin.py`
  - `SlidesAgent/layout_filler.py`
  - `SlidesAgent/gen_figure_match.py`
  - `SlidesAgent/gen_formula.py`
  - `SlidesAgent/gen_speaker.py`
- Updated prompts to condition generation on an author preference profile:
  - `utils/prompts/gen_slides_raw_content_v2.txt`
  - `utils/prompt_templates/layout_agent_xin.yaml`
  - `utils/prompt_templates/preference_distiller.yaml`
  - `utils/prompt_templates/figure_match.yaml`
- Updated `slidegen_openai_utils.py`.
- Updated `webui/backend/main.py`.

### Behavior now present on `dev`
- The pipeline supports author-profile-based personalization through `--use_author_preferences`.
- The pipeline can produce baseline and personalized output variants.
- The repo includes evaluation scripts for scoring generated decks against reference decks.
- The repo includes batch-generation scripts for running larger experiments.

## 2026-04-23 and earlier

### Baseline context
- The repository started from the SlideGen codebase and has since diverged into the capstone branch.
- Earlier work established the preference-distillation direction before full pipeline integration.

---

## Suggested entry format for future updates

```md
## YYYY-MM-DD

### Added
- ...

### Changed
- ...

### Fixed
- ...

### Notes
- ...
```
