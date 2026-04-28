<div align="center">
  
  <img src="./asset/logo2.jpg" height="200" style="object-fit: contain;">

  <h2>Slide Generation Capstone</h2>
  
</div>

## Project Overview

This repository is the working capstone branch derived from **SlideGen** and extended for experimentation with **personalized scientific slide generation**.

The current `dev` branch is not just a copy of upstream SlideGen. It now includes capstone-specific work around:

- author-preference distillation from prior presentation history
- preference-aware slide generation through the main pipeline
- baseline vs personalized run variants
- evaluation tooling for generated decks
- batch experiment scripts
- additional UI work for interacting with the system

The underlying pipeline still follows the SlideGen multi-agent structure:

- **Outliner Agent**: builds a slide-level outline from the paper
- **Mapper Agent**: aligns figures and tables with relevant content
- **Formulizer Agent**: identifies and assigns formulas
- **Arranger Agent**: selects templates and plans slide layouts
- **Speaker Agent**: generates presenter notes
- **Refiner Agent**: merges and adjusts slides for readability

The capstone work extends this backbone with personalization and evaluation rather than replacing it outright.

![](./asset/teaser.jpg)

## Current dev branch focus

The current development direction is:

1. generate slides from scientific papers
2. personalize planning behavior using an author preference profile
3. compare baseline and personalized outputs against reference decks
4. support iterative experimentation without losing reproducibility

See [CHANGELOG.md](./CHANGELOG.md) for the shared branch history.

## Quick start

### 1) Environment

Requirements:
- Python 3.10+
- OpenAI or Azure OpenAI credentials
- LibreOffice for conversion workflows

Example setup:

```bash
conda create -n paper2pptx python=3.12 -y
conda activate paper2pptx

cd Slide_Generation_Capstone

python -m pip install --no-build-isolation \
  "python-pptx @ https://codeload.github.com/Force1ess/python-pptx/zip/dc356685d4d210a10abe1ffab3c21315cdfae63d"

python -m pip install -r requirements.txt
```

Direct OpenAI:

```bash
export OPENAI_API_KEY=your_key
```

If you are using Azure OpenAI, configure the Azure environment variables expected by `slidegen_openai_utils.py`.

### 2) Install LibreOffice

LibreOffice is used for slide-format conversion and headless rendering.

**macOS**
```bash
brew install --cask libreoffice
```

**Ubuntu/Linux**
```bash
sudo apt install libreoffice
```

**Windows**
- Install LibreOffice
- Add the LibreOffice `program` directory to `PATH`

## Running the pipeline

### Baseline generation

```bash
conda activate paper2pptx
cd Slide_Generation_Capstone

python -m SlidesAgent.new_pipeline_logtime \
    --paper_path=your_path \
    --model_name_t="4o" \
    --model_name_v="4o"
```

### Personalized generation

The `dev` branch supports author-profile-based personalization:

```bash
python -m SlidesAgent.new_pipeline \
    --paper_path=your_path \
    --model_name_t="4o-mini" \
    --model_name_v="4o-mini" \
    --use_author_preferences \
    --author_id=your_author_id
```

Notes:
- Replace `--paper_path` with your PDF path.
- `author_id` values come from `Capstone/author_tables/`.
- The current branch supports both baseline and personalized output variants.

## Output

Generated artifacts are written under `contents/<paper_name>/`.

Depending on run mode, outputs may include:

- baseline PPTX decks
- personalized PPTX decks
- raw content JSON
- slide plan JSON
- speaker notes JSON
- logs and detail logs
- evaluation JSON files

## Evaluation and experiments

The `dev` branch includes experiment helpers in `Capstone/`:

- `generate_random_decks.py` for batch generation
- `evaluate_core_coverage.py` for single-deck evaluation
- `evaluate_generated_decks.py` for scanning generated outputs
- `summarize_core_coverage.py`
- `summarize_evaluation_summaries.py`

These scripts are part of the capstone workflow and are not part of the original SlideGen release.

## Interfaces

### Web UI

Backend:

```bash
cd webui/backend
python -m uvicorn main:app --host 0.0.0.0 --port 8000 --reload
```

Frontend:

```bash
cd webui/frontend
npm install
npm run dev
```

### Streamlit UI

The current `dev` branch also includes a Streamlit-based UI layer in `ui/`.

Relevant files:

- `ui/app.py`
- `ui/.env.example`
- `ui/mem0_store.py`

## Example results

![Example 1](./asset/4o_4o_output_slides1_01.jpg)

![Example 2](./asset/4o_4o_output_slides2_01.jpg)
![Example 3](./asset/4o_4o_output_slidesshengwu_01.jpg)

## Origin

This project started from the SlideGen codebase and has been adapted for the capstone's personalization and evaluation goals.

## Citation

If you need to cite the original SlideGen work, use the upstream citation:

```bibtex
@article{liang2025slidegen,
  title={SlideGen: Collaborative Multimodal Agents for Scientific Slide Generation},
  author={Liang, Xin and Zhang, Xiang and Xu, Yiwei and Sun, Siqi and You, Chenyu},
  journal={arXiv preprint arXiv:2512.04529},
  year={2025}
}
```

## Acknowledgments

This capstone codebase is built on SlideGen and related open-source tools. We express our sincere gratitude to:

- **[Docling](https://www.docling.ai/)** for document parsing and conversion
- **[Marker](https://github.com/datalab-to/marker)** for PDF parsing support
- **[python-pptx](https://github.com/scanny/python-pptx)** for PPTX generation
- **[Paper2poster](https://github.com/Paper2Poster/Paper2Poster)** for upstream multi-agent slide-generation ideas
