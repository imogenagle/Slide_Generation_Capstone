"""Generate the two EDA figures referenced in the Data section of the paper.

Outputs:
  fig_author_history.png  - authorship long-tail distribution
  fig_deck_length.png     - distribution of slide counts per deck
"""
from __future__ import annotations

import csv
from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np

ROOT = Path(__file__).resolve().parents[1]
AUTHORS_CSV = ROOT / "author_tables" / "authors.csv"
PAPERS_CSV = ROOT / "author_tables" / "papers.csv"
OUT = Path(__file__).resolve().parent

plt.rcParams.update({
    "font.family": "serif",
    "font.size": 11,
    "axes.titlesize": 12,
    "axes.labelsize": 11,
    "axes.spines.top": False,
    "axes.spines.right": False,
})


def author_history_figure() -> None:
    with open(AUTHORS_CSV) as f:
        counts = [int(r["paper_count"]) for r in csv.DictReader(f)]
    bins = [1, 2, 3, 4, 5, 10, max(counts) + 1]
    labels = ["1", "2", "3", "4", "5-9", "10+"]
    hist, _ = np.histogram(counts, bins=bins)
    total = sum(hist)

    fig, ax = plt.subplots(figsize=(6.0, 3.4))
    bars = ax.bar(labels, hist, color="#4C72B0", edgecolor="white")
    for bar, n in zip(bars, hist):
        ax.text(
            bar.get_x() + bar.get_width() / 2,
            bar.get_height(),
            f"{n:,}\n({n / total:.1%})",
            ha="center", va="bottom", fontsize=9,
        )
    ax.set_xlabel("Paper-deck pairs per author")
    ax.set_ylabel("Number of authors")
    ax.set_ylim(0, max(hist) * 1.18)
    ax.grid(axis="y", linestyle=":", alpha=0.5)
    fig.tight_layout()
    fig.savefig(OUT / "fig_author_history.png", dpi=200, bbox_inches="tight")
    plt.close(fig)


def deck_length_figure() -> None:
    with open(PAPERS_CSV) as f:
        lengths = [
            int(r["slide_image_count"])
            for r in csv.DictReader(f)
            if r["slide_image_count"].isdigit()
        ]
    p95 = int(np.percentile(lengths, 95))
    capped = [min(L, p95) for L in lengths]

    fig, ax = plt.subplots(figsize=(6.0, 3.4))
    ax.hist(capped, bins=np.arange(0, p95 + 2, 2), color="#55A868", edgecolor="white")
    median = int(np.median(lengths))
    mean = float(np.mean(lengths))
    ax.axvline(median, color="black", linestyle="--", linewidth=1)
    ax.text(median + 0.5, ax.get_ylim()[1] * 0.92, f"median = {median}", fontsize=9)
    ax.axvline(mean, color="black", linestyle=":", linewidth=1)
    ax.text(mean + 0.5, ax.get_ylim()[1] * 0.82, f"mean = {mean:.1f}", fontsize=9)
    ax.set_xlabel(f"Slides per deck (capped at 95th percentile = {p95})")
    ax.set_ylabel("Number of decks")
    ax.grid(axis="y", linestyle=":", alpha=0.5)
    fig.tight_layout()
    fig.savefig(OUT / "fig_deck_length.png", dpi=200, bbox_inches="tight")
    plt.close(fig)


if __name__ == "__main__":
    author_history_figure()
    deck_length_figure()
    print("Wrote:")
    for p in sorted(OUT.glob("fig_*.png")):
        print(" ", p.relative_to(ROOT))
