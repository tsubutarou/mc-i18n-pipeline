# -*- coding: utf-8 -*-
"""
診断ツール: 翻訳済みxlsxが「壊れていないか」を機械的にチェックする。

検出するもの:
  1. mods_root には存在するのに、xlsxの diff_all / patchouli_diff に
     1行も出てこない modid (= 01実行後にmodフォルダが増えた、または
     何らかの理由でその場で丸ごと抜け落ちた)
  2. xlsxの en_us 列の値が、実際の mods_root 上のファイルの中身と一致しない行
     (= Googleスプレッドシート上で一部の列だけソート/フィルタして
     行がズレた場合に典型的に発生する。この場合、翻訳がキー違いの場所に
     入ってしまっているので、そのまま03でパックを作ると
     「どこにも正しく反映されない」状態になる)

出力: <out_dir>/diagnose_report.json
"""
from __future__ import annotations

from pathlib import Path
from collections import defaultdict

from openpyxl import load_workbook

from common import (
    load_config, print_progress, safe_write_json, norm,
    build_header_index, load_json_dict, load_json_any, iter_patchouli_strings,
)

LANG_COLS = {"modid": "modid", "key": "key", "en": "en_us"}
PATCH_COLS = {"modid": "modid", "relpath_en": "relpath_en", "ptr": "json_pointer", "en": "en_text"}


def scan_actual_modids(root: Path) -> set[str]:
    modids = set()
    for en_path in root.rglob("assets/*/lang/en_us.json"):
        modids.add(en_path.parent.parent.name)
    for en_path in root.rglob("assets/*/patchouli_books/*/en_us/**/*.json"):
        parts = en_path.parts
        try:
            aidx = parts.index("assets")
            modids.add(parts[aidx + 1])
        except (ValueError, IndexError):
            continue
    return modids


def check_lang_alignment(ws, root: Path) -> tuple[set[str], list[dict]]:
    idx = build_header_index(ws, LANG_COLS)
    total = max(ws.max_row - 1, 0)
    seen_modids: set[str] = set()
    mismatches: list[dict] = []
    en_cache: dict[str, dict] = {}

    for i, r in enumerate(range(2, ws.max_row + 1), start=1):
        if i % 200 == 0 or i == total:
            print_progress(i, total, "diff_all 整合性チェック")

        mod_folder = norm(ws.cell(r, build_header_index(ws, {"mf": "mod_folder"})["mf"]).value).strip()
        modid = norm(ws.cell(r, idx["modid"]).value).strip()
        key = norm(ws.cell(r, idx["key"]).value).strip()
        en_in_sheet = norm(ws.cell(r, idx["en"]).value)
        if not modid or not key:
            continue
        seen_modids.add(modid)

        cache_key = f"{mod_folder}::{modid}"
        if cache_key not in en_cache:
            en_path = Path(mod_folder) / "assets" / modid / "lang" / "en_us.json"
            en_cache[cache_key] = load_json_dict(en_path)
        actual_en = en_cache[cache_key]

        if key in actual_en and actual_en[key] != en_in_sheet:
            mismatches.append({
                "row": r, "modid": modid, "key": key,
                "en_in_sheet": en_in_sheet, "en_on_disk": actual_en[key],
            })

    return seen_modids, mismatches


def check_patchouli_alignment(ws, root: Path) -> tuple[set[str], list[dict]]:
    idx = build_header_index(ws, PATCH_COLS)
    total = max(ws.max_row - 1, 0)
    seen_modids: set[str] = set()
    mismatches: list[dict] = []
    en_cache: dict[str, dict] = {}

    for i, r in enumerate(range(2, ws.max_row + 1), start=1):
        if i % 200 == 0 or i == total:
            print_progress(i, total, "patchouli_diff 整合性チェック")

        modid = norm(ws.cell(r, idx["modid"]).value).strip()
        rel_en = norm(ws.cell(r, idx["relpath_en"]).value).strip()
        ptr = norm(ws.cell(r, idx["ptr"]).value).strip()
        en_in_sheet = norm(ws.cell(r, idx["en"]).value)
        if not modid or not rel_en or not ptr:
            continue
        seen_modids.add(modid)

        if rel_en not in en_cache:
            en_cache[rel_en] = load_json_any(root / rel_en) or {}
        obj = en_cache[rel_en]
        actual_lookup = dict(iter_patchouli_strings(obj))
        actual_en = actual_lookup.get(ptr)

        if actual_en is not None and actual_en != en_in_sheet:
            mismatches.append({
                "row": r, "modid": modid, "relpath_en": rel_en, "json_pointer": ptr,
                "en_in_sheet": en_in_sheet, "en_on_disk": actual_en,
            })

    return seen_modids, mismatches


def run(cfg: dict) -> None:
    if not cfg.get("mods_root"):
        raise SystemExit("[ERROR] config.json に mods_root を指定してください(mod用の設定です)。")
    root = Path(cfg["mods_root"])
    out_dir = Path(cfg["out_dir"])

    xlsx_lang = Path(cfg["xlsx_lang_path"])
    if not xlsx_lang.exists():
        raise SystemExit(f"[ERROR] xlsx (lang) が見つかりません: {xlsx_lang}")

    print(f"[INFO] mods_root 走査中...")
    actual_modids = scan_actual_modids(root)
    print(f"[INFO] mods_root 上のmodid数: {len(actual_modids)}")

    report: dict = {}

    wb = load_workbook(xlsx_lang, data_only=True)
    if "diff_all" in wb.sheetnames:
        seen, mismatches = check_lang_alignment(wb["diff_all"], root)
        missing = sorted(actual_modids - seen)
        report["lang"] = {
            "modids_in_sheet": len(seen),
            "modids_missing_from_sheet": missing,
            "row_mismatches": mismatches,
        }
        print(f"[INFO] lang: シートに無いmodid={len(missing)}件, en_us不一致行={len(mismatches)}件")
        if missing:
            print(f"       -> {missing[:20]}{' ...' if len(missing) > 20 else ''}")

    if cfg.get("include_patchouli", True):
        xlsx_patchouli = Path(cfg["xlsx_patchouli_path"])
        if not xlsx_patchouli.exists():
            print(f"[WARN] xlsx (patchouli) が見つかりません: {xlsx_patchouli} — patchouliの診断はスキップします")
        else:
            wb2 = load_workbook(xlsx_patchouli, data_only=True)
            if "patchouli_diff" in wb2.sheetnames:
                seen_p, mismatches_p = check_patchouli_alignment(wb2["patchouli_diff"], root)
                report["patchouli"] = {
                    "modids_in_sheet": len(seen_p),
                    "row_mismatches": mismatches_p,
                }
                print(f"[INFO] patchouli: en_text不一致行={len(mismatches_p)}件")

    safe_write_json(out_dir / "diagnose_report.json", report)
    print(f"[OK] 診断結果を出力しました: {out_dir / 'diagnose_report.json'}")

    total_mismatch = len(report.get("lang", {}).get("row_mismatches", [])) + \
        len(report.get("patchouli", {}).get("row_mismatches", []))
    if total_mismatch:
        print(f"[WARN] {total_mismatch} 行で en_us の内容が元ファイルと不一致です。")
        print("       Googleスプレッドシート上で一部の列だけソート/フィルタして行がズレた")
        print("       可能性が高いです。該当行は元のxlsx(01の出力直後のもの)から作り直すことを推奨します。")
    if report.get("lang", {}).get("modids_missing_from_sheet"):
        print("[WARN] 上記のmodidはxlsxに1行も存在しません。01実行時点でmodフォルダが"
              "無かった/読み取れなかった可能性があります。")


if __name__ == "__main__":
    import argparse
    ap = argparse.ArgumentParser(description="翻訳済みxlsxの整合性診断(行ズレ・欠落modid検出)")
    ap.add_argument("--config", default="config.json")
    args = ap.parse_args()
    run(load_config(args.config))
