#!/usr/bin/env python3
"""Audit where a personalization experiment is failing stage by stage."""

from __future__ import annotations

import argparse
import csv
import json
import sys
from pathlib import Path
from typing import Any

REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from SlidesAgent.personalization_targets import build_numeric_target_summary
from SlidesAgent.slide_plan_summary import summarize_slide_plan

NUMERIC_METRIC_SPECS = {
    "target_avg_slides_per_section": "avg_slides_per_section",
    "target_avg_bullets_per_slide": "avg_bullets_per_slide",
    "target_avg_words_per_slide": "avg_words_per_slide",
    "target_fraction_figure_slides": "figure_slide_fraction",
    "target_fraction_table_slides": "table_slide_fraction",
    "target_fraction_formula_slides": "formula_slide_fraction",
    "target_fraction_text_only_slides": "text_only_fraction",
    "target_fraction_multi_visual_slides": "multi_visual_fraction",
    "target_fraction_formula_capable_slides": "formula_capable_fraction",
    "target_fraction_image_right_slides": "image_right_fraction",
    "target_fraction_image_left_slides": "image_left_fraction",
    "target_fraction_image_top_slides": "image_top_fraction",
}


def load_json(path: Path) -> dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8"))


def safe_float(value: Any) -> float | None:
    try:
        return float(value)
    except Exception:
        return None


def mean(values: list[float]) -> float:
    return sum(values) / len(values) if values else 0.0


def infer_personalized_variant(personalized_pptx: Path) -> str:
    name = personalized_pptx.name
    if "personalized_strong" in name:
        return "_personalized_strong"
    if "personalized" in name:
        return "_personalized"
    raise ValueError(f"Could not infer personalized variant from {personalized_pptx}")


def plan_paths_from_summary_row(
    row: dict[str, Any],
    *,
    model_name_t: str,
    model_name_v: str,
) -> dict[str, Path]:
    baseline_pptx = Path(row["baseline_pptx"])
    personalized_pptx = Path(row["personalized_pptx"])
    baseline_dir = baseline_pptx.parent
    personalized_dir = personalized_pptx.parent
    personalized_variant = infer_personalized_variant(personalized_pptx)
    model_tag = f"<{model_name_t}_{model_name_v}>"
    trace_tag = f"{model_name_t}_{model_name_v}"
    return {
        "baseline_plan": baseline_dir / f"{model_tag}_slide_plan_baseline.json",
        "personalized_draft_plan": personalized_dir / f"{model_tag}_slide_plan_draft{personalized_variant}.json",
        "personalized_final_plan": personalized_dir / f"{model_tag}_slide_plan{personalized_variant}.json",
        "personalization_trace": personalized_dir / f"{trace_tag}_personalization_trace{personalized_variant}.json",
        "repair_report": personalized_dir / f"{model_tag}_slide_plan_repair_report{personalized_variant}.json",
    }


def resolve_profile_path(row: dict[str, Any], trace: dict[str, Any]) -> Path | None:
    raw_path = str(trace.get("author_profile_path") or "").strip()
    if raw_path:
        candidate = Path(raw_path)
        if not candidate.is_absolute():
            candidate = REPO_ROOT / candidate
        if candidate.exists():
            return candidate

    author_id = str(row.get("author_id") or "").strip()
    if author_id:
        fallback = REPO_ROOT / "Capstone" / "profiles" / f"{author_id}.json"
        if fallback.exists():
            return fallback
    return None


def compute_numeric_distances(
    plan_summary: dict[str, Any],
    target_summary: dict[str, Any],
) -> dict[str, float]:
    distances: dict[str, float] = {}
    for target_key, summary_key in NUMERIC_METRIC_SPECS.items():
        target_value = safe_float(target_summary.get(target_key))
        observed_value = safe_float(plan_summary.get(summary_key))
        if target_value is None or observed_value is None:
            continue
        distances[target_key] = round(abs(observed_value - target_value), 4)
    return distances


def shared_metric_keys(*metric_sets: dict[str, float]) -> list[str]:
    keys: set[str] | None = None
    for metric_set in metric_sets:
        current = set(metric_set)
        keys = current if keys is None else keys & current
    return sorted(keys or [])


def average_distance(metric_distances: dict[str, float], keys: list[str]) -> float:
    values = [metric_distances[key] for key in keys if key in metric_distances]
    return round(mean(values), 4) if values else 0.0


def relative_gain(before: float, after: float) -> float:
    if before <= 0:
        return 0.0
    return round((before - after) / before, 4)


def metric_delta_map(before: dict[str, float], after: dict[str, float], keys: list[str]) -> dict[str, float]:
    deltas: dict[str, float] = {}
    for key in keys:
        if key in before and key in after:
            deltas[key] = round(before[key] - after[key], 4)
    return deltas


def top_metric_names(delta_map: dict[str, float], *, positive: bool, limit: int = 4) -> list[str]:
    filtered = [
        (metric, delta)
        for metric, delta in delta_map.items()
        if (delta > 0.0 if positive else delta < 0.0)
    ]
    filtered.sort(key=lambda item: abs(item[1]), reverse=True)
    return [metric for metric, _ in filtered[:limit]]


def classify_failure_mode(
    *,
    baseline_avg: float,
    draft_avg: float,
    final_avg: float,
    planner_gain: float,
    repair_gain: float,
    total_gain: float,
    blocked_ratio: float,
    profile_source_paper_count: int,
    repair_attempted: bool,
    accepted: bool,
    measurable_target_improvements: list[str],
) -> tuple[str, str]:
    if blocked_ratio >= 0.35:
        return "constraint_limited", "A large share of mismatch score is blocked by missing assets or unsupported targets."
    if profile_source_paper_count <= 2:
        return "profile_risk", "The profile is built from very few source decks, so targets may be unstable."
    if planner_gain <= 0.03 and draft_avg >= baseline_avg * 0.97:
        return "planner_conditioning_weak", "The personalized draft barely moved closer to the target profile than baseline."
    if repair_attempted and repair_gain <= 0.03 and final_avg >= draft_avg * 0.97:
        return "repair_ineffective", "Repair ran, but the final plan barely improved over the personalized draft."
    if repair_attempted and accepted and not measurable_target_improvements:
        return "repair_acceptance_too_permissive", "Repair was accepted without producing measurable gains on actionable target metrics."
    if total_gain <= 0.0 and draft_avg < baseline_avg and final_avg > draft_avg:
        return "repair_regressed_draft", "The personalized draft was closer than baseline, but repair moved the final plan away again."
    if total_gain <= 0.0:
        return "end_to_end_personalization_weak", "The final personalized plan is not closer to the target profile than baseline."
    return "improving_but_incomplete", "Personalization is helping overall, but the final plan is still missing several target dimensions."


def build_paper_audit(
    row: dict[str, Any],
    *,
    model_name_t: str,
    model_name_v: str,
    alignment_dir: Path,
) -> dict[str, Any]:
    paths = plan_paths_from_summary_row(row, model_name_t=model_name_t, model_name_v=model_name_v)
    missing = [label for label, path in paths.items() if not path.exists()]
    if missing:
        return {
            "paper_id": row["paper_id"],
            "author_id": row.get("author_id"),
            "status": "missing_inputs",
            "missing_inputs": missing,
            "paths": {key: str(value) for key, value in paths.items()},
        }

    trace = load_json(paths["personalization_trace"])
    repair_report = load_json(paths["repair_report"]) if paths["repair_report"].exists() else {}
    baseline_plan = load_json(paths["baseline_plan"])
    draft_plan = load_json(paths["personalized_draft_plan"])
    final_plan = load_json(paths["personalized_final_plan"])

    baseline_summary = summarize_slide_plan(baseline_plan)
    draft_summary = summarize_slide_plan(draft_plan)
    final_summary = summarize_slide_plan(final_plan)

    author_profile_path = resolve_profile_path(row, trace)
    author_profile = load_json(author_profile_path) if author_profile_path and author_profile_path.exists() else {}
    target_summary = build_numeric_target_summary(author_profile)

    baseline_distances = compute_numeric_distances(baseline_summary, target_summary)
    draft_distances = compute_numeric_distances(draft_summary, target_summary)
    final_distances = compute_numeric_distances(final_summary, target_summary)
    metric_keys = shared_metric_keys(baseline_distances, draft_distances, final_distances)

    baseline_avg = average_distance(baseline_distances, metric_keys)
    draft_avg = average_distance(draft_distances, metric_keys)
    final_avg = average_distance(final_distances, metric_keys)
    planner_gain = relative_gain(baseline_avg, draft_avg)
    repair_gain = relative_gain(draft_avg, final_avg)
    total_gain = relative_gain(baseline_avg, final_avg)

    planner_deltas = metric_delta_map(baseline_distances, draft_distances, metric_keys)
    repair_deltas = metric_delta_map(draft_distances, final_distances, metric_keys)
    total_deltas = metric_delta_map(baseline_distances, final_distances, metric_keys)

    draft_directives = dict(trace.get("planner", {}).get("draft_repair_directives") or {})
    final_directives = dict(trace.get("repair", {}).get("repaired_repair_directives") or {})
    acceptance = dict(trace.get("repair", {}).get("acceptance") or {})
    blocked_score = safe_float(draft_directives.get("blocked_total_mismatch_score")) or 0.0
    actionable_score = safe_float(draft_directives.get("total_mismatch_score")) or 0.0
    blocked_ratio = round(blocked_score / (blocked_score + actionable_score), 4) if (blocked_score + actionable_score) > 0 else 0.0
    measurable_target_improvements = list(acceptance.get("measurable_target_improvements") or [])

    alignment_path = (
        alignment_dir / f"{row['paper_id'].replace(':', '_')}.alignment.json"
        if (alignment_dir / f"{row['paper_id'].replace(':', '_')}.alignment.json").exists()
        else alignment_dir / f"{row['paper_id'].replace(':', '_')}.pairwise.json"
    )
    alignment = load_json(alignment_path) if alignment_path.exists() else {}
    lift = dict(alignment.get("lift") or {})
    dimensions = dict(alignment.get("dimensions") or {})
    positive_lift_dims = sorted(key for key, value in lift.items() if safe_float(value) and safe_float(value) > 0)
    negative_lift_dims = sorted(key for key, value in lift.items() if safe_float(value) and safe_float(value) < 0)

    profile_source_paper_count = int(target_summary.get("source_paper_count", 0) or 0)
    failure_mode, diagnosis = classify_failure_mode(
        baseline_avg=baseline_avg,
        draft_avg=draft_avg,
        final_avg=final_avg,
        planner_gain=planner_gain,
        repair_gain=repair_gain,
        total_gain=total_gain,
        blocked_ratio=blocked_ratio,
        profile_source_paper_count=profile_source_paper_count,
        repair_attempted=bool(trace.get("repair", {}).get("attempted")),
        accepted=bool(trace.get("repair", {}).get("accepted")),
        measurable_target_improvements=measurable_target_improvements,
    )

    return {
        "paper_id": row["paper_id"],
        "author_id": row.get("author_id"),
        "status": "ok",
        "paths": {key: str(value) for key, value in paths.items()},
        "profile_path": str(author_profile_path) if author_profile_path else None,
        "profile_source_paper_count": profile_source_paper_count,
        "profile_softened_for_sparse": bool(target_summary.get("softened_for_sparse_profile")),
        "target_summary": target_summary,
        "stage_average_target_distance": {
            "baseline": baseline_avg,
            "draft": draft_avg,
            "final": final_avg,
        },
        "stage_relative_gain": {
            "planner_baseline_to_draft": planner_gain,
            "repair_draft_to_final": repair_gain,
            "end_to_end_baseline_to_final": total_gain,
        },
        "metric_distance_breakdown": {
            "baseline": baseline_distances,
            "draft": draft_distances,
            "final": final_distances,
        },
        "metric_delta_breakdown": {
            "planner": planner_deltas,
            "repair": repair_deltas,
            "end_to_end": total_deltas,
        },
        "top_metric_changes": {
            "planner_improved": top_metric_names(planner_deltas, positive=True),
            "planner_worsened": top_metric_names(planner_deltas, positive=False),
            "repair_improved": top_metric_names(repair_deltas, positive=True),
            "repair_worsened": top_metric_names(repair_deltas, positive=False),
            "final_improved_vs_baseline": top_metric_names(total_deltas, positive=True),
            "final_worsened_vs_baseline": top_metric_names(total_deltas, positive=False),
        },
        "repair": {
            "attempted": bool(trace.get("repair", {}).get("attempted")),
            "accepted": bool(trace.get("repair", {}).get("accepted")),
            "acceptance": acceptance,
            "draft_total_mismatch_score": actionable_score,
            "draft_blocked_total_mismatch_score": blocked_score,
            "draft_blocked_ratio": blocked_ratio,
            "draft_priority_metrics": list(draft_directives.get("priority_metrics") or []),
            "draft_blocked_priority_metrics": list(draft_directives.get("blocked_priority_metrics") or []),
            "draft_actionable_targets": list(draft_directives.get("actionable_targets") or []),
            "final_total_mismatch_score": safe_float(final_directives.get("total_mismatch_score")),
            "best_round_index": repair_report.get("best_round_index"),
            "repair_round_count": len(repair_report.get("repair_rounds") or []),
        },
        "evaluation": {
            "positive_lift_dimensions": positive_lift_dims,
            "negative_lift_dimensions": negative_lift_dims,
            "lift": lift,
            "winners": {
                key: (value or {}).get("winner")
                for key, value in dimensions.items()
            },
        },
        "diagnosis": {
            "failure_mode": failure_mode,
            "summary": diagnosis,
        },
    }


def build_aggregate(audits: list[dict[str, Any]]) -> dict[str, Any]:
    ok_rows = [row for row in audits if row.get("status") == "ok"]
    failure_mode_counts: dict[str, int] = {}
    for row in ok_rows:
        mode = str(row.get("diagnosis", {}).get("failure_mode") or "unknown")
        failure_mode_counts[mode] = failure_mode_counts.get(mode, 0) + 1

    def avg_for(key: str) -> float:
        values = [safe_float(row["stage_relative_gain"].get(key)) for row in ok_rows]
        return round(mean([value for value in values if value is not None]), 4) if values else 0.0

    return {
        "num_papers": len(audits),
        "num_ok": len(ok_rows),
        "num_missing_inputs": len(audits) - len(ok_rows),
        "failure_mode_counts": failure_mode_counts,
        "average_stage_relative_gain": {
            "planner_baseline_to_draft": avg_for("planner_baseline_to_draft"),
            "repair_draft_to_final": avg_for("repair_draft_to_final"),
            "end_to_end_baseline_to_final": avg_for("end_to_end_baseline_to_final"),
        },
        "papers": audits,
    }


def write_csv(path: Path, audits: list[dict[str, Any]]) -> None:
    fieldnames = [
        "paper_id",
        "author_id",
        "status",
        "profile_source_paper_count",
        "profile_softened_for_sparse",
        "baseline_avg_distance",
        "draft_avg_distance",
        "final_avg_distance",
        "planner_gain",
        "repair_gain",
        "total_gain",
        "repair_attempted",
        "repair_accepted",
        "repair_round_count",
        "best_round_index",
        "draft_total_mismatch_score",
        "draft_blocked_total_mismatch_score",
        "draft_blocked_ratio",
        "positive_lift_dimensions",
        "negative_lift_dimensions",
        "planner_improved",
        "planner_worsened",
        "repair_improved",
        "repair_worsened",
        "failure_mode",
        "diagnosis_summary",
        "missing_inputs",
    ]

    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        for row in audits:
            if row.get("status") != "ok":
                writer.writerow(
                    {
                        "paper_id": row.get("paper_id"),
                        "author_id": row.get("author_id"),
                        "status": row.get("status"),
                        "missing_inputs": ",".join(row.get("missing_inputs") or []),
                    }
                )
                continue
            writer.writerow(
                {
                    "paper_id": row["paper_id"],
                    "author_id": row.get("author_id"),
                    "status": row.get("status"),
                    "profile_source_paper_count": row.get("profile_source_paper_count"),
                    "profile_softened_for_sparse": row.get("profile_softened_for_sparse"),
                    "baseline_avg_distance": row["stage_average_target_distance"]["baseline"],
                    "draft_avg_distance": row["stage_average_target_distance"]["draft"],
                    "final_avg_distance": row["stage_average_target_distance"]["final"],
                    "planner_gain": row["stage_relative_gain"]["planner_baseline_to_draft"],
                    "repair_gain": row["stage_relative_gain"]["repair_draft_to_final"],
                    "total_gain": row["stage_relative_gain"]["end_to_end_baseline_to_final"],
                    "repair_attempted": row["repair"]["attempted"],
                    "repair_accepted": row["repair"]["accepted"],
                    "repair_round_count": row["repair"]["repair_round_count"],
                    "best_round_index": row["repair"]["best_round_index"],
                    "draft_total_mismatch_score": row["repair"]["draft_total_mismatch_score"],
                    "draft_blocked_total_mismatch_score": row["repair"]["draft_blocked_total_mismatch_score"],
                    "draft_blocked_ratio": row["repair"]["draft_blocked_ratio"],
                    "positive_lift_dimensions": ",".join(row["evaluation"]["positive_lift_dimensions"]),
                    "negative_lift_dimensions": ",".join(row["evaluation"]["negative_lift_dimensions"]),
                    "planner_improved": ",".join(row["top_metric_changes"]["planner_improved"]),
                    "planner_worsened": ",".join(row["top_metric_changes"]["planner_worsened"]),
                    "repair_improved": ",".join(row["top_metric_changes"]["repair_improved"]),
                    "repair_worsened": ",".join(row["top_metric_changes"]["repair_worsened"]),
                    "failure_mode": row["diagnosis"]["failure_mode"],
                    "diagnosis_summary": row["diagnosis"]["summary"],
                }
            )


def main() -> None:
    parser = argparse.ArgumentParser(description="Audit where a personalization experiment is struggling.")
    parser.add_argument("--experiment-dir", type=Path, required=True, help="Experiment directory containing summary.json.")
    parser.add_argument(
        "--output-json",
        type=Path,
        default=None,
        help="Optional output JSON path. Defaults to <experiment-dir>/personalization_audit/SUMMARY.json",
    )
    parser.add_argument(
        "--output-csv",
        type=Path,
        default=None,
        help="Optional output CSV path. Defaults to <experiment-dir>/personalization_audit/DETAILS.csv",
    )
    args = parser.parse_args()

    experiment_dir = args.experiment_dir.resolve()
    summary_path = experiment_dir / "summary.json"
    if not summary_path.exists():
        raise SystemExit(f"Missing experiment summary: {summary_path}")

    experiment_summary = load_json(summary_path)
    papers = list(experiment_summary.get("papers") or [])
    model_name_t = str(experiment_summary.get("model_name_t") or "")
    model_name_v = str(experiment_summary.get("model_name_v") or "")
    if not model_name_t or not model_name_v:
        raise SystemExit(f"summary.json is missing model_name_t/model_name_v: {summary_path}")

    alignment_dir = experiment_dir / "personalization_eval"
    output_root = experiment_dir / "personalization_audit"
    output_json = args.output_json.resolve() if args.output_json else output_root / "SUMMARY.json"
    output_csv = args.output_csv.resolve() if args.output_csv else output_root / "DETAILS.csv"

    audits = [
        build_paper_audit(
            row,
            model_name_t=model_name_t,
            model_name_v=model_name_v,
            alignment_dir=alignment_dir,
        )
        for row in papers
    ]
    aggregate = build_aggregate(audits)

    output_json.parent.mkdir(parents=True, exist_ok=True)
    output_json.write_text(json.dumps(aggregate, indent=2, ensure_ascii=False), encoding="utf-8")
    write_csv(output_csv, audits)

    print(
        json.dumps(
            {
                "output_json": str(output_json),
                "output_csv": str(output_csv),
                "num_papers": aggregate["num_papers"],
                "num_ok": aggregate["num_ok"],
                "failure_mode_counts": aggregate["failure_mode_counts"],
            },
            indent=2,
        )
    )


if __name__ == "__main__":
    main()
