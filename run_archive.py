"""
run_archive.py — PT Research Pipeline v2 — Step 9: Archive & reset

Moves every run-specific output into archive/{short_name}_{YYYYMMDD}/ and
recreates the empty working directories (cache/, summaries/, logs/), so the
pipeline folder is back to a clean baseline and ready for a new topic.

A copy of config.yaml is preserved in the archive folder for provenance
(the working config.yaml is left in place, untouched, for the next run).

Usage:
    python3 run_archive.py              # archive + reset
    python3 run_archive.py --dry-run    # show what would happen, change nothing
    python3 run_archive.py --keep-config-copy=false   # skip copying config.yaml
"""

import argparse
import glob
import json
import logging
import os
import shutil
import sys
from datetime import datetime
from pathlib import Path

import yaml

PIPELINE_ROOT = Path(__file__).parent.resolve()
sys.path.insert(0, str(PIPELINE_ROOT))
os.chdir(PIPELINE_ROOT)

logging.basicConfig(level=logging.INFO,
    format="%(asctime)s [%(levelname)s] — %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)])
log = logging.getLogger(__name__)


def slugify(text: str) -> str:
    return "_".join(text.strip().lower().split())


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--dry-run", action="store_true",
                         help="Show what would be archived/reset without changing anything")
    parser.add_argument("--keep-config-copy", default="true",
                         choices=["true", "false"],
                         help="Copy config.yaml into the archive folder (default: true)")
    args = parser.parse_args()
    dry_run = args.dry_run

    with open("config.yaml") as f:
        cfg = yaml.safe_load(f)

    short_name = slugify(cfg.get("topic", {}).get("short_name", "topic"))
    today = datetime.now().strftime("%Y%m%d")
    archive_dir = Path("archive") / f"{short_name}_{today}"

    paths = cfg.get("paths", {})
    full_text_cache   = Path(paths.get("full_text_cache", "cache/full_text"))
    articles_dir      = Path(paths.get("articles", "articles"))
    summaries_layer0  = Path(paths.get("summaries_layer0", "summaries/layer0"))
    summaries_layer2  = Path(paths.get("summaries_layer2", "summaries/deep"))
    summaries_layer3  = Path(paths.get("summaries_layer3", "summaries/layer3"))
    ledger            = Path(paths.get("ledger", "ledger.csv"))
    layer0_ledger     = Path(paths.get("layer0_ledger", "layer0_ledger.csv"))
    layer2_ledger     = Path(paths.get("layer2_ledger", "layer2_ledger.csv"))
    layer3_ledger     = Path(paths.get("layer3_ledger", "layer3_ledger.csv"))

    cache_root      = full_text_cache.parent if full_text_cache.name else full_text_cache
    summaries_root  = Path("summaries")
    logs_root       = Path("logs")

    # ── Files / dirs to MOVE into the archive (run-specific outputs) ────────
    move_targets = []

    for p in [ledger, layer0_ledger, layer2_ledger, layer3_ledger,
              Path("layer3_clusters.json"), Path("layer3_clusters.csv"),
              Path("layer3_cluster_definitions.json"),
              Path("clinical_impact_ledger.csv"),
              Path(".pipeline_checkpoint.json")]:
        if p.exists():
            move_targets.append(p)

    for pattern in ["*_evidence_base_*.xlsx", "*_synthesis_report_*.md",
                    "*_clinical_impact_report_*.md"]:
        for f in glob.glob(pattern):
            move_targets.append(Path(f))

    for d in [cache_root, articles_dir, summaries_root, logs_root]:
        if d.is_symlink():
            log.warning(f"  SKIPPING {d} — it is a symlink, not a real directory. "
                        f"Remove symlink and ensure real data is in place before archiving.")
        elif d.exists() and any(d.iterdir()):
            move_targets.append(d)

    # ── Report plan ──────────────────────────────────────────────────────────
    log.info("="*60)
    log.info("PT Research Pipeline v2 — Archive & Reset")
    log.info(f"Topic:   {cfg.get('topic', {}).get('short_name', '')}")
    log.info(f"Archive: {archive_dir}")
    log.info("="*60)

    if not move_targets:
        log.info("Nothing to archive — working tree already at baseline.")
        return

    log.info("Will move into archive:")
    for p in move_targets:
        kind = "dir " if p.is_dir() else "file"
        log.info(f"  [{kind}] {p}")

    if args.keep_config_copy == "true":
        log.info("Will COPY (not move) config.yaml into archive for provenance.")

    if dry_run:
        log.info("\n--dry-run: no files were changed.")
        return

    # ── Execute ──────────────────────────────────────────────────────────────
    archive_dir.mkdir(parents=True, exist_ok=True)

    for p in move_targets:
        dest = archive_dir / p.name
        if p.is_symlink():
            log.warning(f"  SKIPPING {p} — symlink detected, not real data. "
                        f"Clean up symlinks before running archive.")
            continue
        if dest.exists():
            log.warning(f"  {dest} already exists in archive — skipping {p}")
            continue
        shutil.move(str(p), str(dest))
        log.info(f"  moved {p} -> {dest}")

    if args.keep_config_copy == "true" and Path("config.yaml").exists():
        shutil.copy2("config.yaml", archive_dir / "config.yaml")
        log.info(f"  copied config.yaml -> {archive_dir / 'config.yaml'}")

    # ── Recreate empty baseline directories ─────────────────────────────────
    summaries_clinical_impact = Path("summaries/clinical_impact")
    for d in [full_text_cache, articles_dir, summaries_layer0, summaries_layer2,
              summaries_layer3, summaries_clinical_impact, logs_root]:
        d.mkdir(parents=True, exist_ok=True)
        log.info(f"  recreated empty: {d}")

    log.info("\n" + "="*60)
    log.info(f"Archive complete: {archive_dir}")
    log.info("Working tree reset to baseline — ready for a new topic.")
    log.info("Edit config.yaml for the next topic and re-run the pipeline.")
    log.info("="*60)


if __name__ == "__main__":
    main()
