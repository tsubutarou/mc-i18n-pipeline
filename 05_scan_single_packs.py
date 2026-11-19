# -*- coding: utf-8 -*-
"""
Step (single packs, 任意): 互換パック等、単体で完結したリソースパック/データパックを
まとめて日本語化するための最初のステップ。

modpack本体(00〜04)は「mods_root直下に複数modフォルダが並ぶ」構造を前提にしているが、
単体パックは普通 pack.mcmeta と assets/ (または data/) が直接ルートにある1個のパック。
このステップでは config.json の single_packs_dir 直下にある各アイテム(フォルダ or .zip)
を1パックとして扱い、まとめて1つのxlsxに翻訳対象を書き出す。

拾う翻訳対象:
  - pack.mcmeta の description (通常の文字列の場合のみ。テキストコンポーネント
    形式(オブジェクト/配列)の場合は参考表示のみで、06での自動反映は行わない)
  - assets/*/lang/en_us.json (あれば)
  - assets/*/patchouli_books/*/en_us/**/*.json (あれば)

翻訳対象が1つも無いパック(構造物データのみのcompatパック等)も、summaryシートに
「翻訳対象なし」として一覧で分かるようにする。

出力: <out_dir>/single_packs_diff.xlsx
  - description_diff / lang_diff / patchouli_diff / summary / _meta の各シート
"""
from __future__ import annotations

import json
from pathlib import Path
from collections import Counter

from openpyxl import Workbook

from common import (
    load_config, load_json_dict, load_json_any, print_progress,
    style_header, autosize_columns, detect_placeholders, detect_patchouli_tokens,
    iter_patchouli_strings, parse_patchouli_en_path, append_row_safe,
    resolve_single_pack_roots, auto_translate_texts,
)


def scan_description(root: Path, cfg: dict) -> dict | None:
    mcmeta = load_json_any(root / "pack.mcmeta")
    if not isinstance(mcmeta, dict):
        return None
    desc = mcmeta.get("pack", {}).get("description")
    if desc is None:
        return None
    if isinstance(desc, str):
        if not desc.strip():
            return None
        result = {"text": desc, "status": "PLAIN", "auto": ""}
        if cfg.get("auto_translate", False):
            translated = auto_translate_texts([desc], label="description自動翻訳")
            result["auto"] = translated.get(desc.strip(), "")
        return result
    # テキストコンポーネント形式(オブジェクト/配列)。参考表示のみ。
    return {"text": json.dumps(desc, ensure_ascii=False), "status": "COMPLEX_FORMAT", "auto": ""}


def scan_pack_lang(ws, pack_name: str, root: Path, cfg: dict) -> int:
    rows_data: list[dict] = []
    for en_path in sorted(root.rglob("assets/*/lang/en_us.json")):
        rel = en_path.relative_to(root)
        parts = rel.parts
        aidx = parts.index("assets")
        modid = parts[aidx + 1]

        en = load_json_dict(en_path)
        ja = load_json_dict(en_path.with_name("ja_jp.json"))
        keys = sorted(set(en.keys()) | set(ja.keys()))

        for key in keys:
            en_val = en.get(key, "")
            ja_val = ja.get(key, "")
            if key in en and key in ja:
                same = str(en_val).strip() and str(en_val).strip() == str(ja_val).strip()
                status = "SAME_AS_EN" if same else "OK"
            elif key in en:
                status = "MISSING_JA"
            else:
                status = "EXTRA_JA"
            needs = status in ("MISSING_JA", "SAME_AS_EN")
            ph_mismatch = set(detect_placeholders(str(en_val))) != set(detect_placeholders(str(ja_val))) and (en_val or ja_val)

            rows_data.append({
                "modid": modid, "key": key, "en_val": en_val, "ja_val": ja_val,
                "status": status, "needs": needs, "ph_mismatch": ph_mismatch,
            })

    translations: dict[str, str] = {}
    if cfg.get("auto_translate", False):
        need_texts = [d["en_val"] for d in rows_data if d["needs"] and str(d["en_val"]).strip()]
        translations = auto_translate_texts(need_texts, label=f"[{pack_name}] lang自動翻訳")

    for d in rows_data:
        r = append_row_safe(ws, [
            pack_name, d["modid"], d["key"],
            d["en_val"], d["ja_val"], "", "",
            d["status"], "TRUE" if d["needs"] else "", "TRUE" if d["ph_mismatch"] else "",
        ])
        if d["needs"] and str(d["en_val"]).strip():
            auto = translations.get(str(d["en_val"]).strip(), "")
            if auto:
                ws.cell(r, 6).value = auto
            else:
                ws.cell(r, 6).value = f'=IFERROR(GOOGLETRANSLATE(D{r},"en","ja"),"")'
    return len(rows_data)


def scan_pack_patchouli(ws, pack_name: str, root: Path, cfg: dict) -> int:
    rows_data: list[dict] = []
    for en_path in sorted(root.rglob("assets/*/patchouli_books/*/en_us/**/*.json")):
        parsed = parse_patchouli_en_path(root, en_path)
        if not parsed:
            continue
        modid, bookid, relpath_en, rel_ja = parsed

        en_obj = load_json_any(en_path)
        if en_obj is None:
            continue
        ja_path = root / rel_ja
        ja_obj = load_json_any(ja_path) if ja_path.exists() else None
        ja_lookup = dict(iter_patchouli_strings(ja_obj)) if ja_obj is not None else {}

        for ptr, en_text in iter_patchouli_strings(en_obj):
            ja_text = ja_lookup.get(ptr, "")
            has_ja = ja_obj is not None and ja_text
            if has_ja and str(en_text).strip() and str(en_text).strip() == str(ja_text).strip():
                status = "SAME_AS_EN"
            elif has_ja:
                status = "OK"
            else:
                status = "MISSING_JA"
            needs = status in ("MISSING_JA", "SAME_AS_EN")
            ph_mismatch = set(detect_placeholders(str(en_text))) != set(detect_placeholders(str(ja_text))) and (en_text or ja_text)
            tok_mismatch = Counter(detect_patchouli_tokens(str(en_text))) != Counter(detect_patchouli_tokens(str(ja_text))) and (en_text or ja_text)

            rows_data.append({
                "modid": modid, "bookid": bookid, "relpath_en": relpath_en, "ptr": ptr,
                "en_text": en_text, "ja_text": ja_text, "status": status, "needs": needs,
                "ph_mismatch": ph_mismatch, "tok_mismatch": tok_mismatch,
            })

    translations: dict[str, str] = {}
    if cfg.get("auto_translate", False):
        need_texts = [d["en_text"] for d in rows_data if d["needs"] and str(d["en_text"]).strip()]
        translations = auto_translate_texts(need_texts, label=f"[{pack_name}] patchouli自動翻訳")

    for d in rows_data:
        r = append_row_safe(ws, [
            pack_name, d["modid"], d["bookid"],
            d["relpath_en"], d["ptr"],
            d["en_text"], d["ja_text"], "", "",
            d["status"], "TRUE" if d["needs"] else "",
            "TRUE" if d["ph_mismatch"] else "", "TRUE" if d["tok_mismatch"] else "",
        ])
        if d["needs"] and str(d["en_text"]).strip():
            auto = translations.get(str(d["en_text"]).strip(), "")
            if auto:
                ws.cell(r, 8).value = auto
            else:
                ws.cell(r, 8).value = f'=IFERROR(GOOGLETRANSLATE(F{r},"en","ja"),"")'
    return len(rows_data)


def run(cfg: dict) -> None:
    pack_roots = resolve_single_pack_roots(cfg)
    if not pack_roots:
        print("[WARN] single_packs_dir 内にパック(フォルダ/.zip)が見つかりませんでした。")
        return
    print(f"[INFO] 単体パック数: {len(pack_roots)}")

    wb = Workbook()
    wb.remove(wb.active)

    meta = wb.create_sheet("_meta")
    meta.append(["pack_name", "source_root"])
    for name, root in pack_roots:
        meta.append([name, str(root.resolve())])

    ws_desc = wb.create_sheet("description_diff")
    ws_desc.append(["pack_name", "en_description", "ja_auto_google", "ja_description(new)", "status"])
    style_header(ws_desc, 5)

    ws_lang = wb.create_sheet("lang_diff")
    ws_lang.append(["pack_name", "modid", "key", "en_us", "ja_jp(existing)",
                     "ja_auto_google", "ja_jp(new)", "status", "needs_translation", "placeholder_mismatch"])
    style_header(ws_lang, 10)

    ws_patch = wb.create_sheet("patchouli_diff")
    ws_patch.append(["pack_name", "modid", "book_id", "relpath_en", "json_pointer",
                      "en_text", "ja_text(existing)", "ja_auto_google", "ja_text(new)",
                      "status", "needs_translation", "placeholder_mismatch", "patchouli_markup_mismatch"])
    style_header(ws_patch, 13)

    pack_summary: list[tuple[str, bool, int, int]] = []

    for i, (pack_name, root) in enumerate(pack_roots, start=1):
        print_progress(i, len(pack_roots), pack_name)

        desc_row = scan_description(root, cfg)
        if desc_row:
            r = append_row_safe(ws_desc, [pack_name, desc_row["text"], "", "", desc_row["status"]])
            if desc_row["status"] == "PLAIN":
                if desc_row["auto"]:
                    ws_desc.cell(r, 3).value = desc_row["auto"]
                else:
                    ws_desc.cell(r, 3).value = f'=IFERROR(GOOGLETRANSLATE(B{r},"en","ja"),"")'

        lang_count = scan_pack_lang(ws_lang, pack_name, root, cfg)
        patchouli_count = scan_pack_patchouli(ws_patch, pack_name, root, cfg)
        pack_summary.append((pack_name, desc_row is not None, lang_count, patchouli_count))

    autosize_columns(ws_desc)
    autosize_columns(ws_lang)
    autosize_columns(ws_patch)

    ws_sum = wb.create_sheet("summary")
    ws_sum.append(["pack_name", "description", "lang_rows", "patchouli_rows", "note"])
    no_content: list[str] = []
    for pack_name, has_desc, lang_count, patchouli_count in pack_summary:
        has_any = has_desc or lang_count or patchouli_count
        if not has_any:
            no_content.append(pack_name)
        ws_sum.append([
            pack_name, "TRUE" if has_desc else "", lang_count, patchouli_count,
            "" if has_any else "翻訳対象なし",
        ])
    style_header(ws_sum, 5)
    autosize_columns(ws_sum)

    out = Path(cfg["single_packs_xlsx_path"])
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)
    print(f"[OK] 保存しました: {out}")

    if no_content:
        print(f"[INFO] 翻訳対象が見つからなかったパック({len(no_content)}件): {', '.join(no_content)}")
    print("[NEXT] xlsxをGoogleドライブにアップロードして翻訳したら、"
          ".\\run_rp.ps1 -Step apply で各パックのコピーに反映してください。")


if __name__ == "__main__":
    import argparse
    ap = argparse.ArgumentParser(description="単体パック群(single_packs_dir)から翻訳対象xlsxを生成する")
    ap.add_argument("--config", default="config.json")
    args = ap.parse_args()
    run(load_config(args.config))
