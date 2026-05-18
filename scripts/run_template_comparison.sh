#!/bin/bash
# Run 2 SlideGen pipelines + 8 binder-only template renderings.
# One plan per author shared across 4 user templates (true A/B comparison).

set -e
set -o pipefail

# Load Azure credentials from .env into the environment.
set -a
source .env
set +a

PYBIN="/opt/anaconda3/envs/paper2pptx/bin/python"
LOG_DIR="contents/_run_logs"
mkdir -p "$LOG_DIR"

UCHICAGO="$HOME/Downloads/UChicago Powerpoint Template_16-9.pptx"
BEE="$HOME/Downloads/bee template.pptx"
BEIGE="$HOME/Downloads/beige template.pptx"
MACARON="$HOME/Downloads/macaron template.pptx"

run_pipeline () {
  local paper_path="$1"
  local paper_name="$2"
  local author_id="$3"
  local log="$LOG_DIR/${paper_name}__pipeline.log"
  echo "=== $(date +%H:%M:%S) PIPELINE $paper_name (author=$author_id) ==="
  "$PYBIN" -m SlidesAgent.new_pipeline_logtime \
    --paper_path="$paper_path" \
    --paper_name="$paper_name" \
    --model_name_t=gpt-5 \
    --model_name_v=gpt-5 \
    --use_author_preferences \
    --author_id="$author_id" 2>&1 | tee "$log"
}

run_template () {
  local paper_name="$1"
  local template_path="$2"
  local label="$3"
  local log="$LOG_DIR/${paper_name}__${label}.log"
  echo "=== $(date +%H:%M:%S) TEMPLATE $paper_name [$label] ==="
  "$PYBIN" scripts/render_with_user_template.py \
    --paper_name="$paper_name" \
    --template_path="$template_path" \
    --template_label="$label" \
    --model_name_t=gpt-5 \
    --model_name_v=gpt-5 \
    --use_author_preferences \
    --author_id=placeholder 2>&1 | tee "$log"
}

echo "=== START $(date) ==="

# ---- URTASUN ----
run_pipeline "data_raw/icml20/709/709_paper.pdf" "urtasun_709" "raquel_urtasun"
run_template "urtasun_709" "$UCHICAGO" "uchicago"
run_template "urtasun_709" "$BEE"      "bee"
run_template "urtasun_709" "$BEIGE"    "beige"
run_template "urtasun_709" "$MACARON"  "macaron"

# ---- HERSHCOVICH ----
run_pipeline "data_raw/acl17/103/103_paper.pdf" "hershcovich_103" "daniel_hershcovich"
run_template "hershcovich_103" "$UCHICAGO" "uchicago"
run_template "hershcovich_103" "$BEE"      "bee"
run_template "hershcovich_103" "$BEIGE"    "beige"
run_template "hershcovich_103" "$MACARON"  "macaron"

echo "=== DONE $(date) ==="
