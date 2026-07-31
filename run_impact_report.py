"""
run_impact_report.py — PT Research Pipeline v2 — Clinical Impact Synthesis

Reads clinical_impact_ledger.csv and generates:
  1. A subclassification synthesis markdown report
  2. Populates two new sheets in the Excel workbook via run_export.py hooks

Usage:
    python3 run_impact_report.py
    python3 run_impact_report.py --output my_impact_report.md
"""

import argparse
import json
import logging
import os
import sys
import yaml
import pandas as pd
import numpy as np
from datetime import datetime
from pathlib import Path

PIPELINE_ROOT = Path(__file__).parent.resolve()
sys.path.insert(0, str(PIPELINE_ROOT))
os.chdir(PIPELINE_ROOT)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] — %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)]
)
log = logging.getLogger(__name__)

STAR_LABELS = {
    (9, 10): "⭐⭐⭐⭐⭐ Overwhelming",
    (7, 8):  "⭐⭐⭐⭐ Strong",
    (5, 6):  "⭐⭐⭐ Moderate",
    (3, 4):  "⭐⭐ Limited",
    (0, 2):  "⭐ Insufficient",
}

# Arms that should be excluded from intervention hierarchy ranking —
# these are comparators/controls, not interventions of interest
CONTROL_ARM_PATTERNS = [
    "control", "sham", "placebo", "standard care", "standard of care",
    "usual care", "waitlist", "wait list", "no treatment", "not_applicable",
    "observation", "conservative", "baseline", "pre-intervention",
]

def is_control_arm(arm_label: str) -> bool:
    """Return True if arm label is a control/sham/comparator arm."""
    if not arm_label or arm_label == "not_reported":
        return True
    label_lower = arm_label.lower().strip()
    return any(pat in label_lower for pat in CONTROL_ARM_PATTERNS)


def load_taxonomy_map(taxonomy_file: str) -> dict:
    """Build a mapping from free-text labels and keywords → canonical taxonomy ID."""
    p = Path(taxonomy_file)
    if not p.exists():
        return {}
    data = yaml.safe_load(p.read_text())
    mapping = {}
    for cls in data.get("classifications", []):
        tid   = cls["id"]
        label = cls["label"].lower().strip()
        mapping[label] = tid
        mapping[tid.lower()] = tid
        for kw in cls.get("keywords", []):
            if len(kw) > 4:  # skip very short keywords to avoid false matches
                mapping[kw.lower()] = tid
    return mapping


def normalize_condition(label: str, taxonomy_map: dict) -> str:
    """Map a free-text condition label to its canonical taxonomy ID if possible."""
    if not label:
        return label
    lower = label.lower().strip()
    # Direct match on ID or label
    if lower in taxonomy_map:
        return taxonomy_map[lower]
    # Partial match — find the taxonomy entry whose keywords appear in the label
    for key, tid in taxonomy_map.items():
        if len(key) > 6 and key in lower:
            return tid
    return label  # return as-is if no match found

def score_to_stars(score) -> str:
    try:
        s = int(score)
    except (TypeError, ValueError):
        return "— Unscored"
    for (lo, hi), label in STAR_LABELS.items():
        if lo <= s <= hi:
            return label
    return "— Unscored"


def safe_mean(series) -> str:
    vals = pd.to_numeric(series, errors="coerce").dropna()
    if len(vals) == 0:
        return "not_reported"
    return f"{vals.mean():.2f}"


def safe_median(series) -> str:
    vals = pd.to_numeric(series, errors="coerce").dropna()
    if len(vals) == 0:
        return "not_reported"
    return f"{vals.median():.2f}"


def build_intervention_hierarchy(group_df: pd.DataFrame) -> pd.DataFrame:
    """Aggregate per-comparison rows into per-intervention summary.
    Control, sham, and comparator arms are excluded from ranking."""
    records = []
    for arm, arm_df in group_df.groupby("winning_arm", dropna=True):
        if not arm or arm == "not_reported":
            continue
        if is_control_arm(arm):
            continue  # exclude control/sham/standard care arms
        scores = pd.to_numeric(arm_df["clinical_leverage_score"], errors="coerce").dropna()
        mcid_met = arm_df["mcid_met"].str.lower().isin(["yes", "borderline"]).sum()
        n_total  = len(arm_df)
        median_score = scores.median() if len(scores) else None

        records.append({
            "winning_arm":              arm,
            "n_comparisons":            n_total,
            "median_leverage_score":    round(median_score, 1) if median_score is not None else None,
            "star_rating":              score_to_stars(median_score),
            "median_delta":             safe_median(arm_df["winning_arm_delta"]),
            "mcid_met_count":           mcid_met,
            "mcid_met_pct":             f"{mcid_met/n_total*100:.0f}%" if n_total else "0%",
            "median_effect_size":       safe_median(arm_df["effect_size"]),
            "typical_protocol":         arm_df["winning_arm_protocol"].mode()[0]
                                        if len(arm_df["winning_arm_protocol"].dropna()) else "not_reported",
            "outcome_measures":         ", ".join(arm_df["outcome_measure"].dropna().unique()),
        })

    if not records:
        return pd.DataFrame()
    df = pd.DataFrame(records)
    df = df.sort_values("median_leverage_score", ascending=False, na_position="last")
    return df


def build_report(df: pd.DataFrame, cfg: dict) -> str:
    topic = cfg.get("topic", {}).get("short_name", "the intervention")
    rq    = cfg.get("research_question", "").strip()
    today = datetime.now().strftime("%Y-%m-%d")
    ci_cfg = cfg.get("clinical_impact", {})
    threshold = ci_cfg.get("leverage_threshold", 7)

    # Load taxonomy and normalize condition labels → canonical IDs
    taxonomy_file = ci_cfg.get("taxonomy_file", "condition_taxonomy.yaml")
    if taxonomy_file:
        taxonomy_map = load_taxonomy_map(taxonomy_file)
    else:
        taxonomy_map = {}
        log.info("No taxonomy file — using dynamic condition labels")
    if taxonomy_map:
        df = df.copy()
        df["condition_classification"] = df["condition_classification"].apply(
            lambda x: normalize_condition(str(x), taxonomy_map) if pd.notna(x) else x
        )
        log.info(f"Condition labels after normalization: {df['condition_classification'].nunique()} unique")

    lines = [
        f"# Clinical Impact Report: {topic.title()}",
        f"",
        f"**Generated:** {today}  ",
        f"**Pipeline:** PT Research Pipeline v2 — Stage 5b Clinical Impact  ",
        f"**Models:** {cfg.get('model', {}).get('layer2', '')}  ",
        f"**Leverage threshold for high-impact flag:** ≥ {threshold}/10",
        f"",
        f"---",
        f"",
        f"## Research Question",
        f"",
        f"> {rq}",
        f"",
        f"---",
        f"",
    ]

    # Executive summary
    total_articles = df["pmid"].nunique() if "pmid" in df.columns else 0
    max_impact = df[df.get("maximal_impact_flag", pd.Series(dtype=bool)) == True]["pmid"].nunique() \
        if "maximal_impact_flag" in df.columns else \
        df[df["clinical_leverage_score"].apply(
            lambda x: int(x) >= threshold if str(x).isdigit() else False
        )]["pmid"].nunique()
    n_measures = df["outcome_measure"].nunique() if "outcome_measure" in df.columns else 0
    conditions = df["condition_classification"].nunique() if "condition_classification" in df.columns else 0

    lines += [
        f"## Executive Summary",
        f"",
        f"| Metric | Value |",
        f"|--------|-------|",
        f"| Articles with clinical impact data | {total_articles} |",
        f"| Clinical subclassifications identified | {conditions} |",
        f"| Unique outcome measures extracted | {n_measures} |",
        f"| High-leverage articles (score ≥ {threshold}) | {max_impact} ({max_impact*100//total_articles if total_articles else 0}%) |",
        f"",
        f"### Clinical Leverage Score Distribution",
        f"",
    ]

    if "clinical_leverage_score" in df.columns:
        score_counts = pd.to_numeric(df.drop_duplicates("pmid")["clinical_leverage_score"],
                                      errors="coerce").dropna()
        for (lo, hi), label in STAR_LABELS.items():
            n = ((score_counts >= lo) & (score_counts <= hi)).sum()
            lines.append(f"- {label}: {n} article(s)")
    lines += ["", "---", ""]

    # Per-condition sections
    lines += ["## Findings by Clinical Subclassification", ""]

    condition_groups = df.groupby("condition_classification", sort=False)

    for condition, cdf in condition_groups:
        n_articles = cdf["pmid"].nunique()
        hierarchy  = build_intervention_hierarchy(cdf)

        # Median leverage for this condition
        med_score  = safe_median(cdf.drop_duplicates("pmid")["clinical_leverage_score"])
        stars      = score_to_stars(med_score)

        lines += [
            f"### {condition}",
            f"",
            f"**Articles:** {n_articles} &nbsp;|&nbsp; "
            f"**Median Leverage Score:** {med_score}/10 &nbsp;|&nbsp; {stars}",
            f"",
        ]

        # High-impact alerts
        alerts = cdf[
            pd.to_numeric(cdf["clinical_leverage_score"], errors="coerce").fillna(0) >= threshold
        ].drop_duplicates("pmid")
        if len(alerts):
            lines += [f"#### 🔴 Maximal Impact Alerts ({len(alerts)} article(s))", ""]
            for _, row in alerts.iterrows():
                headline = row.get("alert_headline") or ""
                body     = row.get("alert_body") or ""
                pmid     = row.get("pmid", "")
                score    = row.get("clinical_leverage_score", "?")
                if headline:
                    lines += [
                        f"**[PMID {pmid}](https://pubmed.ncbi.nlm.nih.gov/{pmid}/) — Leverage {score}/10**",
                        f"> {headline}",
                        f"> {body}",
                        f"",
                    ]

        # Intervention hierarchy table
        if not hierarchy.empty:
            lines += [
                f"#### Intervention Hierarchy",
                f"",
                f"Ranked by median Clinical Leverage Score across all articles in this subclassification.",
                f"",
                f"| Rank | Intervention | n | Median Δ | MCID Met | Effect Size | Leverage | Rating |",
                f"|------|-------------|---|---------|---------|------------|---------|--------|",
            ]
            for rank, (_, row) in enumerate(hierarchy.iterrows(), 1):
                lines.append(
                    f"| {rank} | {row['winning_arm'][:50]} | {row['n_comparisons']} | "
                    f"{row['median_delta']} | {row['mcid_met_count']}/{row['n_comparisons']} | "
                    f"{row['median_effect_size']} | {row.get('median_leverage_score','?')}/10 | "
                    f"{row['star_rating']} |"
                )
            lines += [""]

            # Top protocol detail
            top = hierarchy.iloc[0]
            if top.get("typical_protocol") and top["typical_protocol"] != "not_reported":
                lines += [
                    f"**Highest-leverage protocol:** {top['winning_arm']}  ",
                    f"*{top['typical_protocol']}*",
                    f"",
                ]

        # Outcome measures summary
        om_summary = cdf.groupby("outcome_measure").agg(
            n=("pmid", "nunique"),
            median_delta=("between_group_delta", lambda x: safe_median(x)),
            mcid_met_n=("mcid_met", lambda x: x.str.lower().isin(["yes", "borderline"]).sum()),
        ).reset_index()

        if not om_summary.empty:
            lines += [
                f"#### Outcome Measures Reported",
                f"",
                f"| Outcome Measure | n Studies | Median Between-Group Δ | MCID Met |",
                f"|----------------|-----------|----------------------|---------|",
            ]
            for _, row in om_summary.iterrows():
                n_total = row["n"]
                lines.append(
                    f"| {row['outcome_measure']} | {n_total} | "
                    f"{row['median_delta']} | {row['mcid_met_n']}/{n_total} |"
                )
            lines += [""]

        lines += ["---", ""]

    # Clinical directive summary
    lines += [
        "## Clinical Directive Summary",
        "",
        "The table below summarizes the highest-leverage intervention per clinical subclassification.",
        "This translates quantified evidence into a directive stronger than CPG obligation language alone.",
        "",
        "| Subclassification | Top Intervention | Leverage | Rating | MCID Met |",
        "|------------------|-----------------|---------|--------|---------|",
    ]

    for condition, cdf in condition_groups:
        hierarchy = build_intervention_hierarchy(cdf)
        if hierarchy.empty:
            lines.append(f"| {condition} | No data | — | — | — |")
        else:
            top = hierarchy.iloc[0]
            lines.append(
                f"| {condition} | {top['winning_arm'][:40]} | "
                f"{top.get('median_leverage_score','?')}/10 | {top['star_rating']} | "
                f"{top['mcid_met_count']}/{top['n_comparisons']} |"
            )

    lines += [
        "",
        "---",
        "",
        "## Methods Note",
        "",
        "Clinical impact data extracted by PT Research Pipeline v2 Stage 5b. "
        "MCID thresholds sourced from `mcid_reference.yaml` (cited per measure). "
        "Clinical Leverage Score is a composite rubric (0-10) weighting MCID attainment, "
        "effect size magnitude, statistical significance, protocol replicability, bias risk, "
        "study design, and absence of spin. "
        "All outputs should be reviewed by a qualified clinician before informing clinical decisions.",
        "",
        f"*Generated {today} — PT Research Pipeline v2*",
    ]

    return "\n".join(lines)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--output", default=None, help="Output markdown filename")
    args = parser.parse_args()

    cfg = yaml.safe_load(Path("config.yaml").read_text())
    ci_cfg = cfg.get("clinical_impact", {})
    if not ci_cfg.get("enabled", False):
        log.info("clinical_impact.enabled is false — nothing to do.")
        sys.exit(0)

    ledger_path = Path("clinical_impact_ledger.csv")
    if not ledger_path.exists():
        log.error("clinical_impact_ledger.csv not found — run run_clinical_impact.py first.")
        sys.exit(1)

    df = pd.read_csv(ledger_path)
    log.info(f"Loaded {len(df)} rows from clinical_impact_ledger.csv")
    log.info(f"PMIDs: {df['pmid'].nunique()} | "
             f"Conditions: {df['condition_classification'].nunique()} | "
             f"Measures: {df['outcome_measure'].nunique()}")

    report = build_report(df, cfg)

    topic = cfg.get("topic", {}).get("short_name", "topic").replace(" ", "_")
    today = datetime.now().strftime("%Y%m%d")
    out_name = args.output or f"{topic}_clinical_impact_report_{today}.md"
    Path(out_name).write_text(report)
    log.info(f"Report written: {out_name}")


if __name__ == "__main__":
    main()
