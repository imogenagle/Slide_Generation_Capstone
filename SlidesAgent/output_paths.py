from __future__ import annotations

from pathlib import Path
from typing import Any


def asset_paper_name(args: Any) -> str:
    return str(getattr(args, "asset_paper_name", None) or args.paper_name)


def shared_content_paper_name(args: Any) -> str:
    return str(getattr(args, "asset_paper_name", None) or args.paper_name)


def output_root(args: Any) -> Path:
    raw = getattr(args, "output_dir", None)
    return Path(raw) if raw else Path(".")


def contents_root(args: Any) -> Path:
    return output_root(args) / "contents"


def paper_output_dir(args: Any) -> Path:
    return contents_root(args) / str(args.paper_name)


def shared_content_output_dir(args: Any) -> Path:
    return contents_root(args) / shared_content_paper_name(args)


def image_tables_root(args: Any) -> Path:
    return output_root(args) / f"<{args.model_name_t}_{args.model_name_v}>_images_and_tables"


def paper_image_tables_dir(args: Any) -> Path:
    return image_tables_root(args) / asset_paper_name(args)


def raw_content_path(args: Any) -> Path:
    return shared_content_output_dir(args) / f"<{args.model_name_t}_{args.model_name_v}>_raw_content.json"


def figures_json_path(args: Any) -> Path:
    return shared_content_output_dir(args) / f"<{args.model_name_t}_{args.model_name_v}>_figures.json"


def formula_match_path(args: Any) -> Path:
    return shared_content_output_dir(args) / f"<{args.model_name_t}_{args.model_name_v}>_formula_match.json"


def slide_plan_path(args: Any, variant_suffix: str) -> Path:
    return paper_output_dir(args) / f"<{args.model_name_t}_{args.model_name_v}>_slide_plan{variant_suffix}.json"


def slide_plan_path_for(
    *,
    output_dir: str | Path | None,
    paper_name: str,
    model_name_t: str,
    model_name_v: str,
    variant_suffix: str,
) -> Path:
    base_root = Path(output_dir) if output_dir else Path(".")
    return (
        base_root
        / "contents"
        / str(paper_name)
        / f"<{model_name_t}_{model_name_v}>_slide_plan{variant_suffix}.json"
    )


def slide_plan_draft_path(args: Any, variant_suffix: str) -> Path:
    return paper_output_dir(args) / f"<{args.model_name_t}_{args.model_name_v}>_slide_plan_draft{variant_suffix}.json"


def slide_plan_repair_report_path(args: Any, variant_suffix: str) -> Path:
    return paper_output_dir(args) / f"<{args.model_name_t}_{args.model_name_v}>_slide_plan_repair_report{variant_suffix}.json"


def personalization_trace_path(args: Any, variant_suffix: str) -> Path:
    return paper_output_dir(args) / f"{args.model_name_t}_{args.model_name_v}_personalization_trace{variant_suffix}.json"


def output_pptx_path(args: Any, variant_suffix: str) -> Path:
    return paper_output_dir(args) / f"{args.model_name_t}_{args.model_name_v}_output_slides{variant_suffix}.pptx"


def themed_output_pptx_path(args: Any, variant_suffix: str) -> Path:
    return paper_output_dir(args) / f"{args.model_name_t}_{args.model_name_v}_output_slides{variant_suffix}_themed.pptx"


def log_json_path(args: Any, variant_suffix: str) -> Path:
    return paper_output_dir(args) / f"<{args.model_name_t}_{args.model_name_v}>_log{variant_suffix}.json"


def detail_log_path(args: Any, variant_suffix: str) -> Path:
    return paper_output_dir(args) / f"detail_log{variant_suffix}.json"


def images_json_path(args: Any) -> Path:
    return image_tables_root(args) / f"{asset_paper_name(args)}_images.json"


def tables_json_path(args: Any) -> Path:
    return image_tables_root(args) / f"{asset_paper_name(args)}_tables.json"


def images_filtered_path(args: Any) -> Path:
    return paper_image_tables_dir(args) / "images_filtered.json"


def tables_filtered_path(args: Any) -> Path:
    return paper_image_tables_dir(args) / "tables_filtered.json"


def formula_sections_path(args: Any) -> Path:
    return paper_image_tables_dir(args) / f"{asset_paper_name(args)}_formula_sections.json"


def formulas_json_path(args: Any) -> Path:
    return paper_image_tables_dir(args) / f"{asset_paper_name(args)}_formulas.json"


def formula_crop_path(args: Any, index: str | int) -> Path:
    return paper_image_tables_dir(args) / f"{asset_paper_name(args)}-formula-{index}.png"


def page_image_path(args: Any, page_no: int) -> Path:
    return paper_image_tables_dir(args) / f"{asset_paper_name(args)}-{page_no}.png"


def picture_image_path(args: Any, index: int) -> Path:
    return paper_image_tables_dir(args) / f"{asset_paper_name(args)}-picture-{index}.png"


def table_image_path(args: Any, index: int) -> Path:
    return paper_image_tables_dir(args) / f"{asset_paper_name(args)}-table-{index}.png"


def markdown_embedded_path(args: Any) -> Path:
    return paper_image_tables_dir(args) / f"{asset_paper_name(args)}-with-images.md"


def markdown_referenced_path(args: Any) -> Path:
    return paper_image_tables_dir(args) / f"{asset_paper_name(args)}-with-image-refs.md"


def html_referenced_path(args: Any) -> Path:
    return paper_image_tables_dir(args) / f"{asset_paper_name(args)}-with-image-refs.html"


def referenced_artifacts_dir(args: Any) -> Path:
    return paper_image_tables_dir(args) / f"{asset_paper_name(args)}-with-image-refs_artifacts"


def formula_images_dir(args: Any) -> Path:
    return shared_content_output_dir(args) / "formula_images"


def formula_index_path(args: Any) -> Path:
    return shared_content_output_dir(args) / "formula_index.json"


def formula_mode3_index_path(args: Any) -> Path:
    return shared_content_output_dir(args) / "formula_index_formula_mode3.json"
