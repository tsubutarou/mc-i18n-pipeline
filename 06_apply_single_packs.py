# -*- coding: utf-8 -*-
"""
Step (single packs, 任意): 05_scan_single_packs.py で生成した single_packs_diff.xlsx
の翻訳を、各パックの「そのまま差し替えられるコピー」に反映する。

modpack本体(03_build_pack.py)は差分だけのoverlayリソースパックを作るが、単体の
リソースパック/データパックはそれ自体で完結しているため、overlayではなく
元パックを丸ごとコピーした上に ja_jp.json / patchouli本文 / pack.mcmetaの
descriptionを書き込む(= 元パックの日本語版そのもの)。

  <out_dir>/_single_packs_out/<pack_name>_JA/   翻訳反映後のパック本体
  <out_dir>/<pack_name>_JA.zip                   zip化(config.zip_output=true時)

生成された<pack_name>_JA.zipは、元のパックの代わりにresourcepacks/datapacksへ
そのまま入れて使う想定(overlayではなく置き換え)。
"""
from __future__ import annotations

import shutil
import zipfile
from pathlib import Path
from collections import defaultdict

from openpyxl import load_workbook

from common import (
    load_config, print_progress, safe_write_json, load_json_any, norm,
    build_header_index, pick_ja, sanitize_name, ptr_set, resolve_single_pack_roots,
)

DESC_COLS = {"pack_name": "pack_name", "ja_new": "ja_description(new)",
             "ja_auto": "ja_auto_google", "status": "status"}
LANG_COLS = {"pack_name": "pack_name", "modid": "modid", "key": "key", "en": "en_us",
             "ja_old": "ja_jp(existing)", "ja_auto": "ja_auto_google", "ja_new": "ja_jp(new)"}
PATCH_COLS = {"pack_name": "pack_name", "modid": "modid", "book_id": "book_id",
              "relpath_en": "relpath_en", "ptr": "json_pointer", "en": "en_text",
              "ja_old": "ja_text(existing)", "ja_auto": "ja_auto_google", "ja_new": "ja_text(new)"}


def group_rows(ws, cols: dict[str, str]) -> tuple[dict[str, list[int]], dict[str, int]]:
    idx = build_header_index(ws, cols)
    groups: dict[str, list[int]] = defaultdict(list)
    for r in range(2, ws.max_row + 1):
        pack_name = norm(ws.cell(r, idx["pack_name"]).value).strip()
        if not pack_name:
            continue
        groups[pack_name].append(r)
    return groups, idx


def apply_description(ws, idx, rows: list[int], dest_root: Path) -> str:
    """戻り値: 'applied' | 'skip_complex' | 'no_translation' | 'no_row'"""
    if not rows:
        return "no_row"
    r = rows[0]
    if norm(ws.cell(r, idx["status"]).value).strip() != "PLAIN":
        return "skip_complex"
    ja_new = norm(ws.cell(r, idx["ja_new"]).value).strip()
    ja_final = ja_new or norm(ws.cell(r, idx["ja_auto"]).value).strip()
    if not ja_final:
        return "no_translation"

    mcmeta_path = dest_root / "pack.mcmeta"
    obj = load_json_any(mcmeta_path)
    if not isinstance(obj, dict):
        obj = {}
    obj.setdefault("pack", {})["description"] = ja_final
    safe_write_json(mcmeta_path, obj)
    return "applied"


def apply_lang(ws, idx, rows: list[int], dest_root: Path, lang_dir_name: str) -> int:
    per_modid: dict[str, dict[str, str]] = defaultdict(dict)
    for r in rows:
        modid = norm(ws.cell(r, idx["modid"]).value).strip()
        key = norm(ws.cell(r, idx["key"]).value).strip()
        if not modid or not key:
            continue
        ja_final = pick_ja(ws, r, idx)
        if ja_final:
            per_modid[modid][key] = ja_final

    written = 0
    for modid, kv in per_modid.items():
        if not kv:
            continue
        out_path = dest_root / "assets" / sanitize_name(modid) / lang_dir_name / "ja_jp.json"
        existing = load_json_any(out_path) if out_path.exists() else {}
        if not isinstance(existing, dict):
            existing = {}
        existing.update(kv)
        safe_write_json(out_path, dict(sorted(existing.items())))
        written += len(kv)
    return written


def apply_patchouli(ws, idx, rows: list[int], dest_root: Path) -> int:
    groups: dict[str, list[int]] = defaultdict(list)
    for r in rows:
        rel_en = norm(ws.cell(r, idx["relpath_en"]).value).strip()
        ptr = norm(ws.cell(r, idx["ptr"]).value).strip()
        if rel_en and ptr:
            groups[rel_en].append(r)

    written = 0
    for rel_en, group_rs in groups.items():
        parts = rel_en.split("/")
        try:
            aidx = parts.index("assets")
        except ValueError:
            continue
        asset_parts = parts[aidx:]
        if len(asset_parts) < 6 or asset_parts[4] != "en_us":
            continue
        ja_asset_parts = asset_parts[:]
        ja_asset_parts[4] = "ja_jp"
        target_rel = "/".join(ja_asset_parts)

        en_obj = load_json_any(dest_root / rel_en)
        if en_obj is None:
            continue
        ja_path = dest_root / target_rel
        base_obj = load_json_any(ja_path) if ja_path.exists() else None
        if base_obj is None:
            base_obj = en_obj

        applied = set()
        for r in group_rs:
            ptr = norm(ws.cell(r, idx["ptr"]).value).strip()
            if ptr in applied:
                continue
            ja_final = pick_ja(ws, r, idx)
            if not ja_final:
                continue
            try:
                ptr_set(base_obj, ptr, ja_final)
                applied.add(ptr)
            except Exception:
                continue
        if applied:
            safe_write_json(ja_path, base_obj)
            written += len(applied)
    return written


def zip_pack(pack_root: Path, out_dir: Path, pack_name: str) -> Path:
    zip_path = out_dir / f"{pack_name}.zip"
    if zip_path.exists():
        zip_path.unlink()
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for f in pack_root.rglob("*"):
            if f.is_file():
                zf.write(f, f.relative_to(pack_root))
    return zip_path


def run(cfg: dict) -> None:
    xlsx = Path(cfg["single_packs_xlsx_path"])
    if not xlsx.exists():
        raise SystemExit(f"[ERROR] xlsx が見つかりません: {xlsx} (先に 05_scan_single_packs.py を実行してください)")

    pack_roots = dict(resolve_single_pack_roots(cfg))
    lang_dir_name = cfg.get("lang_dir_name", "lang")
    out_dir = Path(cfg["out_dir"])
    out_root = out_dir / "_single_packs_out"

    print(f"[INFO] xlsx 読み込み: {xlsx}")
    wb = load_workbook(xlsx, data_only=True)

    for name in ("description_diff", "lang_diff", "patchouli_diff"):
        if name not in wb.sheetnames:
            raise SystemExit(f"シート '{name}' が見つかりません(05_scan_single_packs.pyの出力を使ってください)")

    ws_desc = wb["description_diff"]
    ws_lang = wb["lang_diff"]
    ws_patch = wb["patchouli_diff"]

    desc_groups, desc_idx = group_rows(ws_desc, DESC_COLS)
    lang_groups, lang_idx = group_rows(ws_lang, LANG_COLS)
    patch_groups, patch_idx = group_rows(ws_patch, PATCH_COLS)

    all_pack_names = sorted(set(desc_groups) | set(lang_groups) | set(patch_groups))
    if not all_pack_names:
        print("[WARN] xlsxに反映対象の行がありません。")
        return

    total = len(all_pack_names)
    for i, pack_name in enumerate(all_pack_names, start=1):
        print_progress(i, total, pack_name)
        src_root = pack_roots.get(pack_name)
        if src_root is None:
            print(f"\n[WARN] {pack_name}: single_packs_dir 内に元パックが見つからないためスキップします")
            continue

        dest = out_root / f"{pack_name}_JA"
        if dest.exists():
            shutil.rmtree(dest)
        shutil.copytree(src_root, dest, ignore=shutil.ignore_patterns(".extracted_from.txt"))

        desc_status = apply_description(ws_desc, desc_idx, desc_groups.get(pack_name, []), dest)
        lang_written = apply_lang(ws_lang, lang_idx, lang_groups.get(pack_name, []), dest, lang_dir_name)
        patch_written = apply_patchouli(ws_patch, patch_idx, patch_groups.get(pack_name, []), dest)

        if cfg.get("zip_output", True):
            zip_pack(dest, out_dir, f"{pack_name}_JA")

        note = []
        if desc_status == "applied":
            note.append("description反映")
        elif desc_status == "skip_complex":
            note.append("description=複雑な形式のため要手動対応")
        if lang_written:
            note.append(f"lang {lang_written}件")
        if patch_written:
            note.append(f"patchouli {patch_written}件")
        if note:
            print(f"\n[OK] {pack_name}: " + " / ".join(note))

    print(f"\n[OK] 完了。出力先: {out_root}")
    if cfg.get("zip_output", True):
        print(f"[OK] zipも生成しました: {out_dir}")
    print("[NEXT] 生成された <pack_name>_JA.zip を、元のパックの代わりに"
          "resourcepacks(またはdatapacks)フォルダに入れてください。")


if __name__ == "__main__":
    import argparse
    ap = argparse.ArgumentParser(description="single_packs_diff.xlsxの翻訳を各パックのコピーに反映する")
    ap.add_argument("--config", default="config.json")
    args = ap.parse_args()
    run(load_config(args.config))
