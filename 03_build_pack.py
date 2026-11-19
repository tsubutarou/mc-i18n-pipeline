# -*- coding: utf-8 -*-
"""
Step 3: 翻訳済みxlsx(lang用/Patchouli用)を読み込み、それぞれ別々のリソースパックを生成する。

  <out_dir>/<pack_name_lang>/
    pack.mcmeta
    assets/<modid>/<lang_dir_name>/ja_jp.json         (lang差分のみ)

  <out_dir>/<pack_name_patchouli>/          (config.include_patchouli=trueの場合のみ)
    pack.mcmeta
    assets/<modid>/patchouli_books/<book_id>/ja_jp/... (Patchouli差分のみ)

lang本文とPatchouli本文を別パックに分けているのは、Patchouli側だけ有効/無効を
切り替えて不具合の切り分けをしやすくするため。conflicts.json もパックごとに
別々に出力する(<pack_root>/_conflicts.json)。
config.zip_output=true なら各パックの<pack_name>.zip も生成する
(そのままresourcepacksフォルダに入れられる形)。

旧 xlsx_to_modpacks_v1.py / v2.py + xlsx_to_merged_json_v2.py のPatchouli書き戻し部分を統合。
"""
from __future__ import annotations

import shutil
import zipfile
from pathlib import Path
from collections import defaultdict

from openpyxl import load_workbook

from common import (
    load_config, print_progress, safe_write_json, load_json_any, norm,
    build_header_index, pick_ja, sanitize_name, ptr_set,
)

LANG_COLS = {"modid": "modid", "key": "key", "en": "en_us", "ja_old": "ja_jp(existing)",
             "ja_auto": "ja_auto_google", "ja_new": "ja_jp(new)"}

PATCH_COLS = {"modid": "modid", "book_id": "book_id", "relpath_en": "relpath_en",
              "ptr": "json_pointer", "en": "en_text", "ja_old": "ja_text(existing)",
              "ja_auto": "ja_auto_google", "ja_new": "ja_text(new)"}


def build_lang_files(ws, pack_root: Path, lang_dir_name: str, exclude_modids: set[str]) -> tuple[int, dict]:
    idx = build_header_index(ws, LANG_COLS)
    per_mod: dict[str, dict[str, str]] = defaultdict(dict)
    conflicts: dict[str, list] = defaultdict(list)

    used = skipped_no_ja = skipped_excluded = 0
    total = max(ws.max_row - 1, 0)

    for i, r in enumerate(range(2, ws.max_row + 1), start=1):
        if i % 200 == 0 or i == total:
            print_progress(i, total, "lang 書き出し")

        modid = norm(ws.cell(r, idx["modid"]).value).strip()
        key = norm(ws.cell(r, idx["key"]).value).strip()
        if not modid or not key:
            continue
        if modid.lower() in exclude_modids:
            skipped_excluded += 1
            continue

        ja_final = pick_ja(ws, r, idx)
        if not ja_final:
            skipped_no_ja += 1
            continue

        prev = per_mod[modid].get(key)
        if prev is not None and prev != ja_final:
            conflicts[f"{modid}:{key}"].append({"previous": prev, "ja_final": ja_final, "row": r})
            continue

        per_mod[modid][key] = ja_final
        used += 1

    for modid, kv in per_mod.items():
        out_path = pack_root / "assets" / sanitize_name(modid) / lang_dir_name / "ja_jp.json"
        safe_write_json(out_path, dict(sorted(kv.items())))

    print(f"[OK] lang: {len(per_mod)} mod分書き出し "
          f"(採用={used}, 未翻訳スキップ={skipped_no_ja}, 除外modidスキップ={skipped_excluded}, 衝突={len(conflicts)})")
    return used, conflicts


def build_patchouli_files(ws, mods_root: Path, pack_root: Path, exclude_modids: set[str]) -> tuple[int, dict]:
    idx = build_header_index(ws, PATCH_COLS)
    conflicts: dict[str, list] = defaultdict(list)

    # 出力先ファイルごとに行をグループ化
    groups: dict[str, list[int]] = defaultdict(list)
    skipped_excluded = 0
    for r in range(2, ws.max_row + 1):
        modid = norm(ws.cell(r, idx["modid"]).value).strip()
        rel_en = norm(ws.cell(r, idx["relpath_en"]).value).strip()
        ptr = norm(ws.cell(r, idx["ptr"]).value).strip()
        if not rel_en or not ptr:
            continue
        if modid.lower() in exclude_modids:
            skipped_excluded += 1
            continue
        groups[rel_en].append(r)

    total = len(groups)
    files_written = 0
    rows_used = 0

    for gi, (rel_en, rows) in enumerate(groups.items(), start=1):
        if gi % 20 == 0 or gi == total:
            print_progress(gi, total, "patchouli 書き出し")

        # rel_en: "<modフォルダ>/assets/<modid>/patchouli_books/<bookid>/en_us/....json"
        parts = rel_en.split("/")
        try:
            aidx = parts.index("assets")
        except ValueError:
            continue
        asset_parts = parts[aidx:]          # ["assets", modid, "patchouli_books", bookid, "en_us", ...]
        if len(asset_parts) < 6 or asset_parts[4] != "en_us":
            continue
        ja_asset_parts = asset_parts[:]
        ja_asset_parts[4] = "ja_jp"
        target_rel = "/".join(ja_asset_parts)  # pack内での出力先(assets/から)

        en_path = mods_root / rel_en
        ja_src_path = mods_root / "/".join(parts[:aidx] + ja_asset_parts)

        en_obj = load_json_any(en_path)
        if en_obj is None:
            continue
        # 既存のja本があればそれを土台に、無ければenを土台に上書きしていく
        base_obj = load_json_any(ja_src_path) or en_obj

        applied = set()
        for r in rows:
            ptr = norm(ws.cell(r, idx["ptr"]).value).strip()
            ja_final = pick_ja(ws, r, idx)
            if not ja_final:
                continue
            if ptr in applied:
                continue
            try:
                ptr_set(base_obj, ptr, ja_final)
                applied.add(ptr)
                rows_used += 1
            except Exception as e:
                conflicts[target_rel].append({"json_pointer": ptr, "row": r, "error": str(e)})

        if applied:
            out_path = pack_root / target_rel
            safe_write_json(out_path, base_obj)
            files_written += 1

    print(f"[OK] patchouli: {files_written} ファイル書き出し "
          f"(置換件数={rows_used}, 除外modidスキップ={skipped_excluded}, 衝突={len(conflicts)})")
    return rows_used, conflicts


def write_pack_mcmeta(pack_root: Path, cfg: dict, kind_label: str) -> None:
    base_desc = cfg.get("pack_description", "Japanese translation pack")
    safe_write_json(pack_root / "pack.mcmeta", {
        "pack": {
            "pack_format": cfg.get("pack_format", 15),
            "description": f"{base_desc} ({kind_label})",
        }
    })


def zip_pack(pack_root: Path, out_dir: Path, pack_name: str) -> Path:
    zip_path = out_dir / f"{pack_name}.zip"
    if zip_path.exists():
        zip_path.unlink()
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for f in pack_root.rglob("*"):
            if f.is_file():
                zf.write(f, f.relative_to(pack_root))
    return zip_path


def build_one_pack(kind: str, xlsx: Path, sheet_name: str, out_dir: Path, pack_name: str,
                    cfg: dict, mods_root: Path, exclude_modids: set[str]) -> None:
    if not xlsx.exists():
        raise SystemExit(
            f"[ERROR] xlsx ({kind}) が見つかりません: {xlsx} (先に 01_build_xlsx.py を実行してください)"
        )

    pack_root = out_dir / pack_name
    if pack_root.exists():
        shutil.rmtree(pack_root)

    print(f"[INFO] xlsx 読み込み({kind}): {xlsx}")
    wb = load_workbook(xlsx, data_only=True)
    if sheet_name not in wb.sheetnames:
        raise SystemExit(f"シート '{sheet_name}' が見つかりません(xlsx: {xlsx})")

    if kind == "lang":
        lang_dir_name = cfg.get("lang_dir_name", "lang")
        _, conflicts = build_lang_files(wb[sheet_name], pack_root, lang_dir_name, exclude_modids)
        kind_label = "Lang"
    else:
        _, conflicts = build_patchouli_files(wb[sheet_name], mods_root, pack_root, exclude_modids)
        kind_label = "Patchouli"

    write_pack_mcmeta(pack_root, cfg, kind_label)

    if conflicts:
        safe_write_json(pack_root / "_conflicts.json", {kind: conflicts})
        print(f"[WARN] 衝突あり({kind})。詳細: {pack_root / '_conflicts.json'}")

    print(f"[OK] リソースパック生成({kind}): {pack_root}")

    if cfg.get("zip_output", True):
        zip_path = zip_pack(pack_root, out_dir, pack_name)
        print(f"[OK] zip化しました({kind}): {zip_path}")


def run(cfg: dict) -> None:
    if not cfg.get("mods_root"):
        raise SystemExit("[ERROR] config.json に mods_root を指定してください(mod用の設定です)。")
    mods_root = Path(cfg["mods_root"])
    out_dir = Path(cfg["out_dir"])
    exclude_modids = {m.strip().lower() for m in cfg.get("exclude_modids", []) if str(m).strip()}
    if exclude_modids:
        print(f"[INFO] パック生成から除外するmodid: {sorted(exclude_modids)}")

    pack_name_lang = sanitize_name(cfg.get("pack_name_lang", "JP_LangTranslations"))
    build_one_pack(
        "lang", Path(cfg["xlsx_lang_path"]), "diff_all",
        out_dir, pack_name_lang, cfg, mods_root, exclude_modids,
    )

    if cfg.get("include_patchouli", True):
        pack_name_patchouli = sanitize_name(cfg.get("pack_name_patchouli", "JP_PatchouliTranslations"))
        build_one_pack(
            "patchouli", Path(cfg["xlsx_patchouli_path"]), "patchouli_diff",
            out_dir, pack_name_patchouli, cfg, mods_root, exclude_modids,
        )

    if cfg.get("zip_output", True):
        print("[NEXT] 生成されたzipを .minecraft/resourcepacks に入れて、ゲーム内で有効化してください。"
              "lang/Patchouliを別々に有効・無効にできるので、不具合の切り分けにも使えます。")


if __name__ == "__main__":
    import argparse
    ap = argparse.ArgumentParser(description="翻訳済みxlsxからlang用/Patchouli用のリソースパックをそれぞれ生成する")
    ap.add_argument("--config", default="config.json")
    args = ap.parse_args()
    run(load_config(args.config))
