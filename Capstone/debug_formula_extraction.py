#!/usr/bin/env python3
"""Run only the formula-extraction path for a single paper.

This avoids running the full slide-generation pipeline when debugging
formula detection / grouping / matching.
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from dotenv import load_dotenv

REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from SlidesAgent.gen_formula import gen_formula_match_v1
from SlidesAgent.new_pipeline import append_outline_mode_suffix, find_target_paper_id, infer_output_key_from_paper_path, output_key_from_paper_id
from SlidesAgent.output_paths import formula_match_path, formula_sections_path, paper_image_tables_dir, raw_content_path
from SlidesAgent.parse_raw import export_formula_crops_from_texts, export_formula_sections_grouped_json_from_texts, parse_raw
from utils.wei_utils import get_agent_config


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Debug formula extraction without running the full pipeline.")
    parser.add_argument("--paper_path", required=True)
    parser.add_argument("--paper_name", default=None)
    parser.add_argument("--model_name_t", default="gpt-5.4-nano")
    parser.add_argument("--model_name_v", default="gpt-5.4-nano")
    parser.add_argument("--outline_mode", choices=["high_level", "technical"], default="high_level")
    parser.add_argument("--formula_mode", type=int, choices=[1, 2, 3], default=1)
    parser.add_argument("--output_dir", default="outputs/formula_debug")
    parser.add_argument("--tmp_dir", default="tmp")
    parser.add_argument(
        "--page-cap",
        type=int,
        default=12,
        help="Only used for grouped formula section export. Increase to test whether later pages matter.",
    )
    parser.add_argument(
        "--run-match",
        action="store_true",
        help="Also run formula-to-subsection matching after section export.",
    )
    return parser


def prepare_args(args: argparse.Namespace) -> argparse.Namespace:
    if args.paper_name is None:
        target_paper_id = find_target_paper_id(args.paper_path)
        paper_name = output_key_from_paper_id(target_paper_id)
        if paper_name is None:
            paper_name = infer_output_key_from_paper_path(args.paper_path)
        args.paper_name = append_outline_mode_suffix(paper_name, args.outline_mode)
    else:
        args.paper_name = append_outline_mode_suffix(args.paper_name, args.outline_mode)

    args.asset_paper_name = args.paper_name
    args.output_variant_suffix = "_baseline"
    args.output_folder_suffix = ""
    args.use_author_preferences = False
    args.personalization_mode = "retrieval"
    return args


def main() -> None:
    parser = build_arg_parser()
    args = prepare_args(parser.parse_args())

    load_dotenv(REPO_ROOT / ".env")
    Path(args.tmp_dir).mkdir(parents=True, exist_ok=True)

    agent_config_t = get_agent_config(args.model_name_t)

    print(f"[formula-debug] paper_path={args.paper_path}", flush=True)
    print(f"[formula-debug] paper_name={args.paper_name}", flush=True)
    print(f"[formula-debug] output_dir={args.output_dir}", flush=True)

    _, _, _, raw_result = parse_raw(args, agent_config_t, version=2)
    formulas, conv_res = export_formula_crops_from_texts(args, raw_result)
    export_formula_sections_grouped_json_from_texts(args, conv_res, max_page_no_exclusive=args.page_cap)

    print(f"[formula-debug] raw_content={raw_content_path(args)}", flush=True)
    print(f"[formula-debug] formula_sections={formula_sections_path(args)}", flush=True)
    print(
        f"[formula-debug] formula_debug={paper_image_tables_dir(args) / f'{args.asset_paper_name}_formula_debug.json'}",
        flush=True,
    )
    print(f"[formula-debug] cropped_formulas={len(formulas) if isinstance(formulas, list) else 'unknown'}", flush=True)

    if args.run_match:
        total_input_token, total_output_token, time_taken = gen_formula_match_v1(args, agent_config_t, raw_result)
        print(f"[formula-debug] formula_match={formula_match_path(args)}", flush=True)
        print(
            json.dumps(
                {
                    "formula_match_input_tokens": total_input_token,
                    "formula_match_output_tokens": total_output_token,
                    "formula_match_time_taken": time_taken,
                },
                indent=2,
            ),
            flush=True,
        )


if __name__ == "__main__":
    main()
