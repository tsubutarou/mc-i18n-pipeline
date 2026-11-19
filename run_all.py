# -*- coding: utf-8 -*-
"""全ステップ、または指定した1ステップだけを実行するランチャー。

使い方:
  py -3 -u run_all.py --config config.json               # 全ステップ実行
  py -3 -u run_all.py --config config.json --step extract # jar展開だけ
  py -3 -u run_all.py --config config.json --step xlsx    # xlsx生成だけ
  py -3 -u run_all.py --config config.json --step check   # 衝突チェックだけ
  py -3 -u run_all.py --config config.json --step pack    # パック生成だけ
  py -3 -u run_all.py --config config.json --step sp_scan  # 単体パック群のスキャン
  py -3 -u run_all.py --config config.json --step sp_apply # 単体パック群への反映

config.json に mods_jar_dir が設定されている場合、--step all 実行時は
自動的に extract (jar展開) が先頭に追加される。
sp_scan / sp_apply は独立した別ワークフロー(単体の互換パック等向け)なので、
--step all には含まれない。single_packs_dir を設定した上で個別に実行する。
"""
from __future__ import annotations

import argparse

from common import load_config
import importlib


STEP_MODULES = {
    "extract": "00_extract_mods",
    "xlsx": "01_build_xlsx",
    "check": "02_check_conflicts",
    "pack": "03_build_pack",
    "diagnose": "04_diagnose",
    "sp_scan": "05_scan_single_packs",
    "sp_apply": "06_apply_single_packs",
}


def main() -> None:
    ap = argparse.ArgumentParser(description="mc_i18n_pipeline ランチャー")
    ap.add_argument("--config", default="config.json")
    ap.add_argument(
        "--step",
        choices=["extract", "xlsx", "check", "pack", "diagnose", "sp_scan", "sp_apply", "all"],
        default="all",
    )
    ap.add_argument(
        "--force", action="store_true",
        help="extractステップで既に展開済みのjarも強制的に再展開する",
    )
    args = ap.parse_args()

    cfg = load_config(args.config)

    steps = ["xlsx", "check", "pack"] if args.step == "all" else [args.step]
    if args.step == "all" and cfg.get("mods_jar_dir"):
        steps = ["extract"] + steps

    for step in steps:
        print(f"\n===== [{step}] =====")
        mod = importlib.import_module(STEP_MODULES[step])
        if step == "extract":
            mod.run(cfg, force=args.force)
        else:
            mod.run(cfg)


if __name__ == "__main__":
    main()
