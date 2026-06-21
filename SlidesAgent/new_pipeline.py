from math import ceil
import sys
 
from pathlib import Path 
from dotenv import load_dotenv
 
import argparse
import csv
import json
import os
import time
from SlidesAgent.output_paths import (
    detail_log_path,
    figures_json_path,
    formula_match_path,
    images_filtered_path,
    log_json_path,
    output_pptx_path,
    paper_output_dir,
    raw_content_path,
    tables_filtered_path,
)
 
# Create a theme profile here
theme_title_text_color = (255,255,0)
theme_title_fill_color = (255,255,0)
theme = {
    'panel_visible': True,
    'textbox_visible': False,
    'figure_visible': False,
    'panel_theme': {
        'color': theme_title_fill_color,
        'thickness': 5,
        'line_style': 'solid',
    },
    'textbox_theme': None,
    'figure_theme': None,
}


def output_key_from_paper_id(paper_id: str | None) -> str | None:
    if not paper_id:
        return None
    key = str(paper_id).strip().replace(":", "_")
    key = "".join(ch if ch.isalnum() or ch in ("_", "-") else "_" for ch in key)
    key = key.strip("_")
    return key or None


def infer_output_key_from_paper_path(paper_path: str) -> str | None:
    path = Path(paper_path)
    parts = list(path.parts)

    # Common dataset layout:
    # .../<split>/<record_id>/paper.pdf
    if len(parts) >= 3:
        record_id = parts[-2].strip()
        split = parts[-3].strip()
        if record_id.isdigit() and split:
            split_key = "".join(ch if ch.isalnum() or ch in ("_", "-") else "_" for ch in split)
            return f"{split_key}_{record_id}"

    stem = path.stem.replace(" ", "_")
    if stem.lower() == "paper" and len(parts) >= 2:
        record_id = parts[-2].strip()
        if record_id:
            return record_id
    return stem or None


def append_outline_mode_suffix(paper_name: str, outline_mode: str) -> str:
    base = paper_name.strip().replace(" ", "_")
    if base.endswith("_high_level") or base.endswith("_technical"):
        return base
    return f"{base}_{outline_mode}"


def append_output_folder_suffix(paper_name: str, folder_suffix: str | None) -> str:
    base = paper_name.strip().replace(" ", "_")
    if not folder_suffix or base.endswith(folder_suffix):
        return base
    return f"{base}{folder_suffix}"


def find_target_paper_id(paper_path: str) -> str | None:
    papers_csv = Path("Capstone/author_tables/papers.csv")
    if not papers_csv.exists():
        return None

    target_candidates = set()
    raw_target = Path(paper_path)
    target_candidates.add(str(raw_target))
    try:
        target_candidates.add(str(raw_target.resolve()))
    except Exception:
        pass
    try:
        target_candidates.add(str((Path.cwd() / raw_target).resolve()))
    except Exception:
        pass

    with papers_csv.open(newline="", encoding="utf-8") as handle:
        for row in csv.DictReader(handle):
            paper_pdf_path = row.get("paper_pdf_path", "")
            candidates = {paper_pdf_path}
            try:
                candidates.add(str(Path(paper_pdf_path).resolve()))
            except Exception:
                pass
            try:
                if paper_pdf_path.startswith("SlideGen/"):
                    candidates.add(str((Path.cwd() / Path(*Path(paper_pdf_path).parts[1:])).resolve()))
            except Exception:
                pass
            if target_candidates & candidates:
                return row.get("paper_id")
    return None

def extract_title_text(title_raw):
    """ title 为 str / list / dict / list[dict]"""
    if isinstance(title_raw, list):
        parts = []
        for t in title_raw:
            if isinstance(t, dict) and "runs" in t:
                for run in t["runs"]:
                    parts.append(run.get("text", ""))
            else:
                parts.append(str(t))
        return ' '.join(parts)
    elif isinstance(title_raw, dict):
        return str(title_raw.get('text', ''))
    else:
        return str(title_raw)

def extract_bullet_text(bullet_raw): 
    if isinstance(bullet_raw, list):
        return ' '.join([extract_bullet_text(b) for b in bullet_raw])
    elif isinstance(bullet_raw, dict):
        if "text" in bullet_raw:
            return bullet_raw["text"]
        elif "runs" in bullet_raw:
            return ''.join([r.get("text", "") for r in bullet_raw["runs"]])
        else:
            return ""
    else:
        return str(bullet_raw)

def save_panels(panels, paper_name, save_dir="outputs"):
    os.makedirs(save_dir, exist_ok=True)
    with open(os.path.join(save_dir, f"{paper_name}_panels.json"), "w") as f:
        json.dump(panels, f, indent=4)

def load_panels(paper_name, save_dir="outputs"):
    with open(os.path.join(save_dir, f"{paper_name}_panels.json"), "r") as f:
        return json.load(f)


from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx import Presentation
def _import_pipeline_modules(personalization_mode: str = "standard"):
    from pptx.util import Inches
    from SlidesAgent.parse_raw import (
        parse_raw,
        gen_image_and_table,
        export_formula_crops_from_texts,
        export_formula_sections_grouped_json_from_texts,
    )
    from SlidesAgent.gen_figure_match import gen_figure_match, filter_image_table
    from SlidesAgent.gen_formula import build_formula_json, gen_formula_match_v1
    if personalization_mode == "retrieval":
        from SlidesAgent.layout_agent_xin_retrieval import generate_slide_plan
    else:
        from SlidesAgent.layout_agent_xin import generate_slide_plan
    from SlidesAgent.layout_filler import generate_pptx_from_plan
    from Capstone.preference_distill import distill_author_profile
    from utils.wei_utils import get_agent_config

    return {
        "Inches": Inches,
        "parse_raw": parse_raw,
        "gen_image_and_table": gen_image_and_table,
        "export_formula_crops_from_texts": export_formula_crops_from_texts,
        "export_formula_sections_grouped_json_from_texts": export_formula_sections_grouped_json_from_texts,
        "gen_figure_match": gen_figure_match,
        "filter_image_table": filter_image_table,
        "build_formula_json": build_formula_json,
        "gen_formula_match_v1": gen_formula_match_v1,
        "generate_slide_plan": generate_slide_plan,
        "generate_pptx_from_plan": generate_pptx_from_plan,
        "distill_author_profile": distill_author_profile,
        "get_agent_config": get_agent_config,
    }


def build_arg_parser(*, description: str = 'Poster Generation Pipeline') -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=description)
    parser.add_argument('--paper_path', type=str)
    parser.add_argument('--model_name_t', type=str, default='4o')
    parser.add_argument('--model_name_v', type=str, default='4o')
    parser.add_argument('--index', type=int, default=0)
    parser.add_argument('--paper_name', type=str, default=None)
    parser.add_argument(
        '--output_dir',
        type=str,
        default='.',
        help=(
            'Base directory for run artifacts. The pipeline will create '
            '"contents/<paper_name>" and "<model>_images_and_tables/<paper_name>" under this path.'
        ),
    )
    parser.add_argument('--tmp_dir', type=str, default='tmp')
    parser.add_argument('--no_blank_detection', action='store_true', help='When overflow is severe, try this option.')
    parser.add_argument('--ablation_no_tree_layout', action='store_true', help='Ablation study: no tree layout')
    parser.add_argument('--ablation_no_commenter', action='store_true', help='Ablation study: no commenter')
    parser.add_argument('--ablation_no_example', action='store_true', help='Ablation study: no example')
    parser.add_argument(
        '--outline_mode',
        choices=['high_level', 'technical'],
        default='high_level',
        help='Use high_level for a compact presentation narrative, or technical to preserve major paper subsections.',
    )
    parser.add_argument(
        "--formula_mode",
        type=int,
        choices=[1, 2, 3],
        default=1,
        help=(
            "Method to add formulas: "
            "1 = use bbox crop from docling, "
            "2 = use LaTeX code rendering, "
            "3 = use user-marked boxes"
        ),
    )
    parser.add_argument(
        '--use_author_preferences',
        action='store_true',
        help='Use a distilled author preference profile when generating the slide plan.',
    )
    parser.add_argument('--author_id', type=str, default=None, help='Canonical author_id used by the preference distiller.')
    parser.add_argument(
        '--author_profile_path',
        type=str,
        default=None,
        help='Optional path to an existing distilled author profile JSON.',
    )
    parser.add_argument(
        '--preference_model',
        type=str,
        default=None,
        help='Model used to generate the author preference profile if needed. Defaults to --model_name_t.',
    )
    parser.add_argument(
        '--preference_max_papers',
        type=int,
        default=5,
        help='Maximum number of prior decks to sample for preference distillation.',
    )
    parser.add_argument(
        '--force_refresh_preferences',
        action='store_true',
        help='Regenerate the author profile even if a cached profile JSON already exists.',
    )
    parser.add_argument(
        '--personalization_mode',
        choices=['standard', 'retrieval'],
        default='standard',
        help='Choose the standard personalization path or the retrieval-conditioned personalization path.',
    )
    return parser


def configure_variant_args(args: argparse.Namespace) -> bool:
    if getattr(args, "preference_model", None) is None:
        args.preference_model = args.model_name_t

    requested_personalized = bool(getattr(args, "use_author_preferences", False))

    if not getattr(args, "output_variant_suffix", None):
        if requested_personalized:
            args.output_variant_suffix = (
                "_personalized_retrieval" if getattr(args, "personalization_mode", "standard") == "retrieval"
                else "_personalized"
            )
        else:
            args.output_variant_suffix = "_baseline"

    if not hasattr(args, "output_folder_suffix") or args.output_folder_suffix is None:
        if requested_personalized:
            args.output_folder_suffix = (
                "_personalized_retrieval" if getattr(args, "personalization_mode", "standard") == "retrieval"
                else "_personalized"
            )
        else:
            args.output_folder_suffix = ""

    return requested_personalized


def prepare_author_profile(args: argparse.Namespace, distill_author_profile, detail_log: dict) -> None:
    if not getattr(args, "use_author_preferences", False):
        return
    if getattr(args, "personalization_mode", "standard") == "retrieval" and not args.author_profile_path:
        raise ValueError(
            "--author_profile_path is required when --personalization_mode retrieval is enabled. "
            "This mode expects a prebuilt retrieval profile JSON."
        )
    if args.author_profile_path:
        profile_path = Path(args.author_profile_path)
    else:
        if not args.author_id:
            raise ValueError("--author_id is required when --use_author_preferences is enabled without --author_profile_path.")
        profile_path = Path("Capstone/profiles") / f"{args.author_id}.json"

    if profile_path.exists() and not args.force_refresh_preferences:
        print(f"[preferences] Reusing existing author profile: {profile_path}", flush=True)
        args.author_profile_path = str(profile_path)
        detail_log['author_profile_path'] = args.author_profile_path
        return

    if not args.author_id:
        raise ValueError("--author_id is required when regenerating an author preference profile.")

    exclude_pdf_paths = set()
    try:
        exclude_pdf_paths.add(str(Path(args.paper_path).resolve()))
    except Exception:
        exclude_pdf_paths.add(args.paper_path)
    target_paper_id = find_target_paper_id(args.paper_path)
    exclude_paper_ids = {target_paper_id} if target_paper_id else set()
    try:
        print(
            f"[preferences] Starting profile distillation for author_id={args.author_id} "
            f"(force_refresh={args.force_refresh_preferences}, max_papers={args.preference_max_papers})",
            flush=True,
        )
        profile = distill_author_profile(
            args.author_id,
            output_dir=profile_path.parent,
            max_papers=args.preference_max_papers,
            model=args.preference_model,
            force_refresh=args.force_refresh_preferences,
            exclude_paper_ids=exclude_paper_ids,
            exclude_pdf_paths=exclude_pdf_paths,
        )
        print(f"[preferences] Profile distillation finished: {profile_path}", flush=True)
        profile_path.write_text(json.dumps(profile, indent=2, ensure_ascii=False), encoding="utf-8")
        args.author_profile_path = str(profile_path)
        detail_log['author_profile_path'] = args.author_profile_path
        detail_log['preference_target_excluded_paper_id'] = target_paper_id
    except ValueError as exc:
        print(f"[preferences] {exc}")
        print("[preferences] No non-target history remains; falling back to baseline planning.")
        args.use_author_preferences = False
        args.author_profile_path = None
        detail_log['author_preferences_fallback_reason'] = str(exc)


def run_pipeline(args: argparse.Namespace) -> None:
    load_dotenv(Path(__file__).resolve().parents[1] / ".env")
    requested_personalized = configure_variant_args(args)

    if args.formula_mode == 1:
        print("👉 Using Docling bbox crop method...")
    elif args.formula_mode == 2:
        print("👉 Using Docling LaTeX rendering method...")
    elif args.formula_mode == 3:
        print("👉 Using user-marked boxes method...")

    start_time = time.time()
    os.makedirs(args.tmp_dir, exist_ok=True)

    detail_log = {}
    pipeline = _import_pipeline_modules(getattr(args, "personalization_mode", "standard"))
    Inches = pipeline["Inches"]
    parse_raw = pipeline["parse_raw"]
    gen_image_and_table = pipeline["gen_image_and_table"]
    export_formula_crops_from_texts = pipeline["export_formula_crops_from_texts"]
    export_formula_sections_grouped_json_from_texts = pipeline["export_formula_sections_grouped_json_from_texts"]
    gen_figure_match = pipeline["gen_figure_match"]
    filter_image_table = pipeline["filter_image_table"]
    build_formula_json = pipeline["build_formula_json"]
    gen_formula_match_v1 = pipeline["gen_formula_match_v1"]
    generate_slide_plan = pipeline["generate_slide_plan"]
    generate_pptx_from_plan = pipeline["generate_pptx_from_plan"]
    distill_author_profile = pipeline["distill_author_profile"]
    get_agent_config = pipeline["get_agent_config"]

    detail_log['outline_mode'] = args.outline_mode
    detail_log['personalization_mode'] = getattr(args, "personalization_mode", "standard")
    slide_width_inches = 13.33
    slide_height_inches = 7.5
    slide_width = Inches(slide_width_inches)
    slide_height = Inches(slide_height_inches)

    if args.paper_name is None:
        target_paper_id = find_target_paper_id(args.paper_path)
        paper_name = output_key_from_paper_id(target_paper_id)
        if paper_name is None:
            paper_name = infer_output_key_from_paper_path(args.paper_path)
        args.paper_name = append_outline_mode_suffix(paper_name, args.outline_mode)
    else:
        paper_name = append_outline_mode_suffix(args.paper_name, args.outline_mode)
        args.paper_name = paper_name

    agent_config_t = get_agent_config(args.model_name_t)
    agent_config_v = get_agent_config(args.model_name_v)
    total_input_tokens_t, total_output_tokens_t = 0, 0
    total_input_tokens_v, total_output_tokens_v = 0, 0

    print(f'slides size: {slide_width_inches} x {slide_height_inches} inches')

    prepare_author_profile(args, distill_author_profile, detail_log)

    args.paper_name = append_output_folder_suffix(args.paper_name, getattr(args, "output_folder_suffix", ""))
    detail_log['output_paper_name'] = args.paper_name
    detail_log['output_dir'] = str(getattr(args, "output_dir", "."))
    detail_log['requested_personalized_run'] = requested_personalized
    detail_log['effective_use_author_preferences'] = getattr(args, "use_author_preferences", False)

    figs_json_path = figures_json_path(args)
    formula_json_path = formula_match_path(args)
    paper_outline_json = raw_content_path(args)
    output_pptx = output_pptx_path(args, args.output_variant_suffix)
    images_filtered_json = images_filtered_path(args)
    tables_filtered_json = tables_filtered_path(args)
    reuse_cached_parse_artifacts = all(
        os.path.exists(path)
        for path in (
            paper_outline_json,
            figs_json_path,
            formula_json_path,
            images_filtered_json,
            tables_filtered_json,
        )
    )

    if reuse_cached_parse_artifacts:
        print(
            "[pipeline] Reusing cached raw/figure/formula artifacts; skipping docling parse",
            flush=True,
        )
        detail_log['reused_cached_parse_artifacts'] = True
    else:
        print("[pipeline] Starting raw paper parsing", flush=True)
        input_token, output_token, _parse_time_taken, raw_result = parse_raw(args, agent_config_t, version=2)
        print("[pipeline] Raw paper parsing finished", flush=True)

        total_input_tokens_t += input_token
        total_output_tokens_t += output_token
        print("[pipeline] Extracting images and tables", flush=True)
        _, _, images, tables = gen_image_and_table(args, raw_result)
        print("[pipeline] Image/table extraction finished", flush=True)

        if args.formula_mode == 1:
            print("start export_formula_crops_from_texts")
            formulas, conv_res = export_formula_crops_from_texts(args, raw_result)
            print("start export_formula_sections_grouped_json_from_texts")
            export_formula_sections_grouped_json_from_texts(args, conv_res)
        elif args.formula_mode == 3:
            print("add formula")
            build_formula_json(args, raw_result)

        print(f'Parsing token consumption: {input_token} -> {output_token}')

        detail_log['parser_in_t'] = input_token
        detail_log['parser_out_t'] = output_token
        input_token, output_token = filter_image_table(args, agent_config_t)
        total_input_tokens_t += input_token
        total_output_tokens_t += output_token
        print(f'Filter figures token consumption: {input_token} -> {output_token}')

        detail_log['filter_in_t'] = input_token
        detail_log['filter_out_t'] = output_token
        input_token, output_token, _figure_match_time, figures = gen_figure_match(args, agent_config_t, raw_result)
        total_input_tokens_t += input_token
        total_output_tokens_t += output_token

        input_token, output_token, _formula_match_time = gen_formula_match_v1(args, agent_config_t, raw_result)
        total_input_tokens_t += input_token
        total_output_tokens_t += output_token

    input_token, output_token, _time_taken = generate_slide_plan(args)
    total_input_tokens_t += input_token
    total_output_tokens_t += output_token

    end_time = time.time()
    time_taken = end_time - start_time
    print("time_taken:", time_taken)

    output_dir = paper_output_dir(args)
    variant_suffix = args.output_variant_suffix
    output_dir.mkdir(parents=True, exist_ok=True)
    log_file = log_json_path(args, variant_suffix)
    with open(log_file, 'w') as f:
        log_data = {
            'input_tokens_t': total_input_tokens_t,
            'output_tokens_t': total_output_tokens_t,
            'input_tokens_v': total_input_tokens_v,
            'output_tokens_v': total_output_tokens_v,
            'time_taken': time_taken,
        }
        json.dump(log_data, f, indent=4)

    print("✅ all files exist……")
    generate_pptx_from_plan(args, 3)

    detail_log_file = detail_log_path(args, variant_suffix)
    with open(detail_log_file, 'w') as f:
        json.dump(detail_log, f, indent=4)


def main() -> None:
    parser = build_arg_parser()
    args = parser.parse_args()
    run_pipeline(args)


if __name__ == '__main__':
    main()
