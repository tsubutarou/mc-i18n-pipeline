# -*- coding: utf-8 -*-
"""
Step 2 (任意/確認用): 翻訳済みxlsx(lang用/Patchouli用)を読み込み、
  - 全mod横断でキーが衝突していないか(同じkeyに違う訳が入っていないか)
  - 訳が空のまま残っている行がどれだけあるか
をチェックし、merged_ja.json / merged_conflicts.json (+ patchouli版) として出力する。

このステップはリソースパック生成(03)には使われない独立した確認用ツール。
スキップしても03は実行できる。

旧 xlsx_to_merged_json.py / _v2.py の統合版。
"""
from __future__ import annotations

from pathlib import Path
from collections import defaultdict

from openpyxl import load_workbook

from common import (
    load_config, print_progress, safe_write_json, norm,
    build_header_index, pick_ja, ptr_get,
)

LANG_COLS = {"modid": "modid", "key": "key", "en": "en_us", "ja_old": "ja_jp(existing)",
             "ja_auto": "ja_auto_google", "ja_new": "ja_jp(new)"}

PATCH_COLS = {"modid": "modid", "book_id": "book_id", "relpath_en": "relpath_en",
              "ptr": "json_pointer", "en": "en_text", "ja_old": "ja_text(existing)",
              "ja_auto": "ja_auto_google", "ja_new": "ja_text(new)"}


def check_lang(ws) -> tuple[dict, dict, dict]:
    idx = build_header_index(ws, LANG_COLS)
    total = max(ws.max_row - 1, 0)
    merged: dict[str, str] = {}
    conflicts: dict[str, list] = defaultdict(list)
    skipped_no_ja = 0

    for i, r in enumerate(range(2, ws.max_row + 1), start=1):
        if i % 200 == 0 or i == total:
            print_progress(i, total, "diff_all 確認中")

        key = norm(ws.cell(r, idx["key"]).value).strip()
        if not key:
            continue
        ja_final = pick_ja(ws, r, idx)
        if not ja_final:
            skipped_no_ja += 1
            continue
        if key in merged and merged[key] != ja_final:
            conflicts[key].append({
                "modid": norm(ws.cell(r, idx["modid"]).value).strip(),
                "ja_final": ja_final, "previous": merged[key], "row": r,
            })
            continue
        merged[key] = ja_final

    stats = {"used": len(merged), "skipped_no_ja": skipped_no_ja, "conflicts": len(conflicts)}
    return merged, conflicts, stats


def check_patchouli(ws) -> tuple[dict, dict]:
    idx = build_header_index(ws, PATCH_COLS)
    total = max(ws.max_row - 1, 0)
    merged: dict[str, str] = {}
    conflicts: dict[str, list] = defaultdict(list)
    skipped_no_ja = 0

    for i, r in enumerate(range(2, ws.max_row + 1), start=1):
        if i % 200 == 0 or i == total:
            print_progress(i, total, "patchouli_diff 確認中")

        rel = norm(ws.cell(r, idx["relpath_en"]).value).strip()
        ptr = norm(ws.cell(r, idx["ptr"]).value).strip()
        if not rel or not ptr:
            continue
        composite_key = f"{rel}{ptr}"
        ja_final = pick_ja(ws, r, idx)
        if not ja_final:
            skipped_no_ja += 1
            continue
        if composite_key in merged and merged[composite_key] != ja_final:
            conflicts[composite_key].append({
                "modid": norm(ws.cell(r, idx["modid"]).value).strip(),
                "ja_final": ja_final, "previous": merged[composite_key], "row": r,
            })
            continue
        merged[composite_key] = ja_final

    stats = {"used": len(merged), "skipped_no_ja": skipped_no_ja, "conflicts": len(conflicts)}
    return merged, conflicts, stats  # type: ignore[return-value]


def run(cfg: dict) -> None:
    out_dir = Path(cfg["out_dir"])

    xlsx_lang = Path(cfg["xlsx_lang_path"])
    if not xlsx_lang.exists():
        raise SystemExit(f"[ERROR] xlsx (lang) が見つかりません: {xlsx_lang}")

    print(f"[INFO] xlsx 読み込み(lang): {xlsx_lang}")
    wb = load_workbook(xlsx_lang, data_only=True)

    if "diff_all" not in wb.sheetnames:
        raise SystemExit("シート 'diff_all' が見つかりません(lang xlsx)")
    ws = wb["diff_all"]
    merged, conflicts, stats = check_lang(ws)
    print(f"[INFO] lang: 採用={stats['used']} 未翻訳スキップ={stats['skipped_no_ja']} 衝突={stats['conflicts']}")
    safe_write_json(out_dir / "merged_ja.json", merged)
    safe_write_json(out_dir / "merged_conflicts.json", conflicts)

    if cfg.get("include_patchouli", True):
        xlsx_patchouli = Path(cfg["xlsx_patchouli_path"])
        if not xlsx_patchouli.exists():
            print(f"[WARN] xlsx (patchouli) が見つかりません: {xlsx_patchouli} — patchouliのチェックはスキップします")
        else:
            print(f"[INFO] xlsx 読み込み(patchouli): {xlsx_patchouli}")
            wb2 = load_workbook(xlsx_patchouli, data_only=True)
            if "patchouli_diff" in wb2.sheetnames:
                ws2 = wb2["patchouli_diff"]
                p_merged, p_conflicts, p_stats = check_patchouli(ws2)
                print(f"[INFO] patchouli: 採用={p_stats['used']} 未翻訳スキップ={p_stats['skipped_no_ja']} 衝突={p_stats['conflicts']}")
                safe_write_json(out_dir / "patchouli_conflicts.json", p_conflicts)

    if conflicts:
        print(f"[WARN] lang に {len(conflicts)} 件のキー衝突があります → merged_conflicts.json を確認してください")
    print(f"[OK] 確認用ファイルを出力しました: {out_dir}")


if __name__ == "__main__":
    import argparse
    ap = argparse.ArgumentParser(description="翻訳済みxlsxのキー衝突・未翻訳チェック")
    ap.add_argument("--config", default="config.json")
    args = ap.parse_args()
    run(load_config(args.config))
