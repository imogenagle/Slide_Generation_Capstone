"""Render an already-generated SlideGen plan against a user PPTX template.

Use this AFTER running new_pipeline_logtime.py once with the same paper_name
(and same --use_author_preferences setting). This script skips all SlideGen
agents and only calls the layout binder + user-template renderer.

Example:
    python -m scripts.render_with_user_template \\
        --paper_name=urtasun_709 \\
        --author_id=raquel_urtasun \\
        --use_author_preferences \\
        --template_path="~/Downloads/UChicago Powerpoint Template_16-9.pptx" \\
        --template_label=uchicago
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path
from types import SimpleNamespace

REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from dotenv import load_dotenv

load_dotenv(REPO_ROOT / ".env")

from SlidesAgent.new_pipeline_logtime import (
    append_outline_mode_suffix,
    append_personalized_folder_suffix,
)
from SlidesAgent.layout_binder import bind_and_render


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--paper_name", required=True,
                        help="Same value passed to new_pipeline_logtime.py (before suffixing).")
    parser.add_argument("--template_path", required=True,
                        help="Path to the user-supplied .pptx template.")
    parser.add_argument("--template_label", default=None,
                        help="Short label for output filenames (e.g. uchicago). Derived from filename if omitted.")
    parser.add_argument("--model_name_t", default="gpt-5")
    parser.add_argument("--model_name_v", default="gpt-5")
    parser.add_argument("--outline_mode", choices=["high_level", "technical"], default="high_level")
    parser.add_argument("--formula_mode", type=int, choices=[1, 2, 3], default=1,
                        help="Must match what was used in the SlideGen run.")
    parser.add_argument("--use_author_preferences", action="store_true")
    parser.add_argument("--author_id", default=None,
                        help="Carried for completeness; only used if other code reads it.")
    cli = parser.parse_args()

    paper_name = append_outline_mode_suffix(cli.paper_name, cli.outline_mode)
    paper_name = append_personalized_folder_suffix(paper_name, cli.use_author_preferences)

    args = SimpleNamespace(
        paper_name=paper_name,
        model_name_t=cli.model_name_t,
        model_name_v=cli.model_name_v,
        outline_mode=cli.outline_mode,
        formula_mode=cli.formula_mode,
        use_author_preferences=cli.use_author_preferences,
        author_id=cli.author_id,
    )

    output_pptx = bind_and_render(args, cli.template_path, template_label=cli.template_label)
    if output_pptx is None:
        return 1
    print(f"[render_with_user_template] DONE -> {output_pptx}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
