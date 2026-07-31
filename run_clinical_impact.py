"""
run_clinical_impact.py — PT Research Pipeline v2 — Stage 5b: Clinical Impact

Runs AFTER Layer 2. Reads existing *_layer2.json files and the full article
text, then runs a focused clinical impact extraction:
  - Dynamic outcome measure discovery
  - MCID-aware arm-by-arm data extraction
  - Clinical Leverage Score calculation (0-10)
  - Maximal impact flagging

Writes results to:
  summaries/clinical_impact/{pmid}_impact.json   (per-article)
  clinical_impact_ledger.csv                     (corpus-wide ledger)

The impact ledger is read by run_export.py to populate two new Excel sheets:
  "Clinical Impact"        — per-article per-outcome-measure
  "Intervention Hierarchy" — aggregated per condition subclassification

Usage:
    python3 run_clinical_impact.py
    python3 run_clinical_impact.py --pmid 12345678   # single article
    python3 run_clinical_impact.py --rerun           # reprocess all
"""

import argparse
import json
import logging
import os
import sys
import yaml
import requests
import pandas as pd
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


def load_config():
    cfg = yaml.safe_load(Path("config.yaml").read_text())
    ci_cfg = cfg.get("clinical_impact", {})
    if not ci_cfg.get("enabled", False):
        log.info("clinical_impact.enabled is false in config.yaml — nothing to do.")
        sys.exit(0)
    return cfg, ci_cfg


def load_mcid_reference(mcid_file: str) -> dict:
    p = Path(mcid_file)
    if not p.exists():
        log.error(f"MCID reference file not found: {mcid_file}")
        sys.exit(1)
    data = yaml.safe_load(p.read_text())
    return data.get("outcome_measures", {})


def load_taxonomy(taxonomy_file: str) -> list:
    p = Path(taxonomy_file)
    if not p.exists():
        log.error(f"Condition taxonomy file not found: {taxonomy_file}")
        sys.exit(1)
    data = yaml.safe_load(p.read_text())
    return data.get("classifications", [])


def get_relevant_mcid_entries(mcid_db: dict, topic_keywords: list) -> dict:
    """Filter MCID database to measures relevant for this topic."""
    relevant = {}
    topic_kw = [k.lower() for k in topic_keywords]
    for key, measure in mcid_db.items():
        conditions = [c.lower() for c in measure.get("conditions", [])]
        if any(kw in cond for kw in topic_kw for cond in conditions):
            relevant[key] = measure
        elif "general" in conditions or "musculoskeletal pain" in conditions:
            relevant[key] = measure
    return relevant


def format_mcid_entries_for_prompt(mcid_entries: dict) -> str:
    lines = []
    for abbr, m in mcid_entries.items():
        line = (
            f"{abbr} ({m['full_name']}): "
            f"scale {m.get('scale','?')}, "
            f"MCID absolute={m.get('mcid_absolute','?')}, "
            f"percent={m.get('mcid_percent','?')}%, "
            f"direction={m.get('direction','?')}, "
            f"source={m.get('source','?')}"
        )
        lines.append(line)
    return "\n".join(lines) if lines else "No specific MCID entries for this topic — use training knowledge and flag as estimated."


def format_taxonomy_for_prompt(taxonomy: list) -> str:
    lines = []
    for t in taxonomy:
        kw = ", ".join(t.get("keywords", [])[:8])
        lines.append(f"  {t['id']}: {t['label']} — keywords: {kw}")
    return "\n".join(lines)


def load_full_text(pmid: str, cfg: dict) -> str:
    cache_dir = Path(cfg["paths"].get("full_text_cache", "cache/full_text"))
    txt_path = cache_dir / f"{pmid}.txt"
    if txt_path.exists():
        return txt_path.read_text(encoding="utf-8", errors="replace")
    xml_path = Path("articles") / f"{pmid}.xml"
    if xml_path.exists():
        return xml_path.read_text(encoding="utf-8", errors="replace")
    return ""


def call_ollama(prompt: str, cfg: dict) -> dict:
    model_cfg = cfg.get("model", {})
    opts = model_cfg.get("layer2_options", {})
    payload = {
        "model": model_cfg.get("layer2", "qwen3.6:35b-a3b"),
        "prompt": prompt,
        "format": "json",
        "stream": False,
        "options": {
            "temperature": opts.get("temperature", 0.1),
            "num_predict": opts.get("num_predict", 2500),
            "num_ctx": opts.get("num_ctx", 40000),
        },
    }
    if "think" in opts:
        payload["think"] = opts["think"]

    base_url = model_cfg.get("base_url", "http://localhost:11434")
    try:
        resp = requests.post(
            f"{base_url}/api/generate",
            json=payload,
            timeout=300,
        )
        resp.raise_for_status()
        raw = resp.json().get("response", "").strip()
        raw = raw.lstrip("```json").lstrip("```").rstrip("```").strip()
        return json.loads(raw)
    except Exception as e:
        log.error(f"Ollama call failed: {e}")
        return {}


def run_impact_extraction(pmid: str, layer2: dict, full_text: str,
                           prompt_template: str, mcid_entries: dict,
                           taxonomy: list, cfg: dict) -> dict:
    s1 = layer2.get("stage_outputs", {}).get("s1", {})
    s2 = layer2.get("stage_outputs", {}).get("s2", {})
    s3 = layer2.get("stage_outputs", {}).get("s3", {})
    s5 = layer2.get("stage_outputs", {}).get("s5", {})

    # Build context
    topic_cfg = cfg.get("topic", {})
    topic_keywords = topic_cfg.get("relevance_criterion", "").split()[:20]
    relevant_mcid = get_relevant_mcid_entries(mcid_entries, topic_keywords)
    mcid_text = format_mcid_entries_for_prompt(relevant_mcid)
    taxonomy_text = format_taxonomy_for_prompt(taxonomy)

    # Truncate full text for context window
    layer2_opts = cfg.get("model", {}).get("layer2_options", {})
    num_ctx = layer2_opts.get("num_ctx", 40000)
    char_limit = int(num_ctx * 2.8)
    text_chunk = full_text[:char_limit] if full_text else ""

    prompt = prompt_template
    replacements = {
        "{study_design}":        s1.get("study_design", layer2.get("study_design", "")),
        "{sample_size}":         s1.get("sample_size", layer2.get("sample_size", "")),
        "{intervention}":        s1.get("intervention", layer2.get("intervention", "")),
        "{comparator}":          s1.get("comparator", layer2.get("comparator", "")),
        "{blinding}":            s2.get("blinding", layer2.get("blinding", "")),
        "{bias_risk}":           s2.get("bias_risk_structured", layer2.get("bias_risk", "")),
        "{statistical_significance}": s5.get("statistical_significance",
                                             layer2.get("statistical_significance", "")),
        "{spin_detected}":       s5.get("spin_detected", layer2.get("spin_detected", "")),
        "{oxford_level}":        s3.get("oxford_roman", layer2.get("oxford_roman", "")),
        "{condition_taxonomy}":  taxonomy_text,
        "{mcid_entries}":        mcid_text,
        "{full_text}":           text_chunk,
    }
    for key, val in replacements.items():
        prompt = prompt.replace(key, str(val))

    result = call_ollama(prompt, cfg)
    if not result:
        return {}

    result["pmid"] = pmid
    result["title"] = layer2.get("title", "")
    result["run_date"] = datetime.now().strftime("%Y-%m-%d")
    result["model_used"] = cfg.get("model", {}).get("layer2", "")
    return result


def flatten_for_ledger(impact: dict) -> list:
    """Flatten per-article impact JSON into rows for the CSV ledger."""
    rows = []
    base = {
        "pmid":                        impact.get("pmid", ""),
        "title":                       impact.get("title", ""),
        "condition_classification":    impact.get("condition_classification", ""),
        "condition_taxonomy_id":       impact.get("condition_taxonomy_id", ""),
        "taxonomy_match":              impact.get("taxonomy_match", ""),
        "clinical_leverage_score":     impact.get("clinical_leverage_score", ""),
        "maximal_impact_flag":         impact.get("maximal_impact_flag", ""),
        "alert_headline":              impact.get("alert_headline", ""),
        "alert_body":                  impact.get("alert_body", ""),
        "primary_outcome_measure":     impact.get("primary_outcome_measure", ""),
        "leverage_mcid_met":           impact.get("leverage_score_breakdown", {}).get("mcid_met_primary", ""),
        "leverage_large_effect":       impact.get("leverage_score_breakdown", {}).get("large_effect_size", ""),
        "leverage_p_sig":              impact.get("leverage_score_breakdown", {}).get("p_significant", ""),
        "leverage_replicable":         impact.get("leverage_score_breakdown", {}).get("protocol_replicable", ""),
        "leverage_low_bias":           impact.get("leverage_score_breakdown", {}).get("low_moderate_bias", ""),
        "leverage_rct":                impact.get("leverage_score_breakdown", {}).get("rct_or_higher", ""),
        "leverage_no_spin":            impact.get("leverage_score_breakdown", {}).get("no_spin", ""),
        "data_quality_notes":          impact.get("data_quality_notes", ""),
        "run_date":                    impact.get("run_date", ""),
        "model_used":                  impact.get("model_used", ""),
    }

    for om in impact.get("outcome_measures", []):
        for comp in om.get("between_group_comparisons", []):
            row = dict(base)
            row.update({
                "outcome_measure":          om.get("measure_name", ""),
                "outcome_abbreviation":     om.get("measure_abbreviation", ""),
                "is_primary_outcome":       om.get("is_primary_outcome", ""),
                "mcid_absolute":            om.get("mcid_absolute", ""),
                "mcid_percent":             om.get("mcid_percent", ""),
                "mcid_source":              om.get("mcid_source", ""),
                "winning_arm":              om.get("winning_arm", ""),
                "winning_arm_delta":        om.get("winning_arm_delta", ""),
                "winning_arm_protocol":     om.get("winning_arm_protocol", ""),
                "arm_a":                    comp.get("arm_a", ""),
                "arm_b":                    comp.get("arm_b", ""),
                "between_group_delta":      comp.get("between_group_delta", ""),
                "ci_95":                    str(comp.get("ci_95", "")),
                "between_group_p":          comp.get("between_group_p", ""),
                "effect_size":              comp.get("effect_size", ""),
                "effect_size_type":         comp.get("effect_size_type", ""),
                "effect_size_magnitude":    comp.get("effect_size_magnitude", ""),
                "mcid_met":                 comp.get("mcid_met", ""),
                "mcid_threshold_used":      comp.get("mcid_threshold_used", ""),
            })
            rows.append(row)

    # If no outcome measures extracted, still write a summary row
    if not rows:
        rows.append(base)

    return rows


def main():
    parser = argparse.ArgumentParser(description="Stage 5b: Clinical Impact Extraction")
    parser.add_argument("--pmid", default=None, help="Process a single PMID only")
    parser.add_argument("--rerun", action="store_true",
                        help="Reprocess all articles even if already done")
    args = parser.parse_args()

    cfg, ci_cfg = load_config()

    mcid_file    = ci_cfg.get("mcid_file", "mcid_reference.yaml")
    taxonomy_file = ci_cfg.get("taxonomy_file") or None
    output_dir   = Path("summaries/clinical_impact")
    ledger_path  = Path("clinical_impact_ledger.csv")
    prompt_file  = Path("prompts/stages/stage5b_clinical_impact.txt")
    layer2_dir   = Path(cfg["paths"].get("summaries_layer2", "summaries/deep"))

    output_dir.mkdir(parents=True, exist_ok=True)

    mcid_db   = load_mcid_reference(mcid_file)
    if taxonomy_file:
        taxonomy = load_taxonomy(taxonomy_file)
    else:
        taxonomy = []
        log.info("No taxonomy file configured — using dynamic subclassification")

    if not prompt_file.exists():
        log.error(f"Stage 5b prompt not found: {prompt_file}")
        sys.exit(1)
    prompt_template = prompt_file.read_text()

    # Find Layer 2 JSON files
    layer2_files = sorted(layer2_dir.glob("*_layer2.json"))
    if args.pmid:
        layer2_files = [f for f in layer2_files if args.pmid in f.name]

    if not layer2_files:
        log.error("No Layer 2 JSON files found — run Layer 2 first.")
        sys.exit(1)

    log.info("=" * 60)
    log.info("PT Research Pipeline v2 — Stage 5b: Clinical Impact")
    log.info(f"Topic:     {cfg.get('topic', {}).get('short_name', '')}")
    log.info(f"Articles:  {len(layer2_files)}")
    log.info(f"MCID file: {mcid_file} ({len(mcid_db)} measures)")
    log.info(f"Taxonomy:  {taxonomy_file} ({len(taxonomy)} classifications)")
    log.info("=" * 60)

    all_rows = []
    completed = 0
    failed = 0
    skipped = 0

    for layer2_path in layer2_files:
        pmid = layer2_path.stem.replace("_layer2", "")
        out_path = output_dir / f"{pmid}_impact.json"

        if out_path.exists() and not args.rerun:
            log.info(f"PMID {pmid}: already done — skipping")
            # Still load for ledger aggregation
            try:
                existing = json.loads(out_path.read_text())
                all_rows.extend(flatten_for_ledger(existing))
                skipped += 1
            except Exception:
                pass
            continue

        try:
            layer2 = json.loads(layer2_path.read_text())
        except Exception as e:
            log.error(f"PMID {pmid}: failed to load layer2 JSON: {e}")
            failed += 1
            continue

        full_text = load_full_text(pmid, cfg)
        if not full_text:
            log.warning(f"PMID {pmid}: no full text — skipping clinical impact")
            skipped += 1
            continue

        log.info(f"PMID {pmid}: running Stage 5b...")
        result = run_impact_extraction(
            pmid, layer2, full_text, prompt_template,
            mcid_db, taxonomy, cfg
        )

        if not result:
            log.error(f"PMID {pmid}: Stage 5b returned empty result")
            failed += 1
            continue

        out_path.write_text(json.dumps(result, indent=2))
        rows = flatten_for_ledger(result)
        all_rows.extend(rows)

        score = result.get("clinical_leverage_score", "?")
        flag = "🔴 MAXIMAL IMPACT" if result.get("maximal_impact_flag") else ""
        n_measures = len(result.get("outcome_measures", []))
        log.info(f"PMID {pmid}: ✓ leverage={score}/10 | measures={n_measures} | {flag}")
        if result.get("alert_headline"):
            log.info(f"  → {result['alert_headline']}")

        completed += 1

    # Write ledger
    if all_rows:
        df = pd.DataFrame(all_rows)
        df.to_csv(ledger_path, index=False)
        log.info(f"\nLedger saved: {ledger_path} ({len(df)} rows)")

    log.info("\n" + "=" * 60)
    log.info("Stage 5b complete")
    log.info(f"  Completed: {completed}")
    log.info(f"  Skipped:   {skipped}")
    log.info(f"  Failed:    {failed}")
    log.info(f"  Max impact articles: {sum(1 for r in all_rows if r.get('maximal_impact_flag') is True)}")
    log.info("=" * 60)


if __name__ == "__main__":
    main()
