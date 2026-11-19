# -*- coding: utf-8 -*-
"""
Step 1: 展開済みmodフォルダ群を走査して、翻訳用xlsxを2つ生成する。

  ALL_mods_lang_diff.xlsx（lang用）
    - diff_all シート      : assets/*/lang/en_us.json vs ja_jp.json の差分
    - summary シート        : lang件数サマリ
    - _meta シート          : mods_root を記録(後続ステップが参照)

  ALL_mods_patchouli_diff.xlsx（Patchouli用、config.include_patchouli=false で生成しない）
    - patchouli_diff シート : assets/*/patchouli_books/*/en_us/**/*.json の地の文差分
    - summary シート        : patchouli件数サマリ
    - _meta シート          : mods_root を記録(後続ステップが参照)

lang本文とPatchouli本文は不具合の出方が異なることが多いため、xlsxファイル自体を
分けている(03_build_pack.pyもこの2ファイルからそれぞれ別のリソースパックを作る)。

旧 mods_lang_to_one_xlsx.py / _v2.py の統合版。
"""
from __future__ import annotations

import sys
from pathlib import Path
from collections import Counter

from openpyxl import Workbook

from common import (
    load_config, load_json_dict, load_json_any, print_progress,
    style_header, autosize_columns, detect_placeholders, detect_patchouli_tokens,
    iter_patchouli_strings, parse_patchouli_en_path, append_row_safe, auto_translate_texts,
)


def make_lang_sheet(wb: Workbook, root: Path, cfg: dict) -> dict:
    ws = wb.create_sheet("diff_all")
    headers = [
        "mod_folder", "modid", "key",
        "en_us", "ja_jp(existing)", "ja_auto_google", "ja_jp(new)",
        "status", "needs_translation", "placeholder_mismatch",
    ]
    ws.append(headers)
    style_header(ws, len(headers))

    en_files = sorted(root.rglob("assets/*/lang/en_us.json"))
    total = len(en_files)
    print(f"[INFO] en_us.json 件数: {total}", flush=True)

    status_counter: Counter = Counter()
    placeholder_mismatch_count = 0
    missing_by_mod: Counter = Counter()
    rows_data: list[dict] = []

    for i, en_path in enumerate(en_files, start=1):
        print_progress(i, total, f"[lang] {en_path.parents[3].name}")

        try:
            rel = en_path.relative_to(root)
            parts = rel.parts
            aidx = parts.index("assets")
            mod_folder = str(root.joinpath(*parts[:aidx]))
            modid = parts[aidx + 1]
        except Exception:
            mod_folder = str(en_path.parent.parent.parent)
            modid = ""

        ja_path = en_path.with_name("ja_jp.json")
        en = load_json_dict(en_path)
        ja = load_json_dict(ja_path)

        keys = sorted(set(en.keys()) | set(ja.keys()))
        for key in keys:
            en_val = en.get(key, "")
            ja_val = ja.get(key, "")

            if key in en and key in ja:
                if str(en_val).strip() and str(en_val).strip() == str(ja_val).strip():
                    # modが未翻訳のまま en_us をコピーしただけの ja_jp.json を
                    # 同梱しているケース。訳が入っているように見えるが実質未翻訳。
                    status = "SAME_AS_EN"
                else:
                    status = "OK"
            elif key in en and key not in ja:
                status = "MISSING_JA"
                missing_by_mod[mod_folder] += 1
            else:
                status = "EXTRA_JA"

            needs_translation = status in ("MISSING_JA", "SAME_AS_EN")

            ph_en = set(detect_placeholders(str(en_val)))
            ph_ja = set(detect_placeholders(str(ja_val)))
            ph_mismatch = ph_en != ph_ja and (en_val or ja_val)
            if ph_mismatch:
                placeholder_mismatch_count += 1

            rows_data.append({
                "mod_folder": mod_folder, "modid": modid, "key": key,
                "en_val": en_val, "ja_val": ja_val,
                "status": status, "needs": needs_translation, "ph_mismatch": ph_mismatch,
            })
            status_counter[status] += 1

    translations: dict[str, str] = {}
    if cfg.get("auto_translate", False):
        need_texts = [d["en_val"] for d in rows_data if d["needs"] and str(d["en_val"]).strip()]
        translations = auto_translate_texts(need_texts, label="[lang] 自動翻訳")

    for d in rows_data:
        r = append_row_safe(ws, [
            d["mod_folder"], d["modid"], d["key"],
            d["en_val"], d["ja_val"], "", "",
            d["status"],
            "TRUE" if d["needs"] else "",
            "TRUE" if d["ph_mismatch"] else "",
        ])

        # 既訳が無い、または既存訳がen_usのコピーでしかない場合に自動翻訳を入れる。
        # auto_translateが有効かつ翻訳に成功した場合はテキストをそのまま書き込み、
        # それ以外(無効/失敗)はGoogle Sheets側で評価される数式にフォールバックする。
        if d["needs"] and str(d["en_val"]).strip():
            auto = translations.get(str(d["en_val"]).strip(), "")
            if auto:
                ws.cell(r, 6).value = auto
            else:
                ws.cell(r, 6).value = f'=IFERROR(GOOGLETRANSLATE(D{r},"en","ja"),"")'

    autosize_columns(ws)
    return {
        "lang_status": status_counter,
        "lang_placeholder_mismatch": placeholder_mismatch_count,
        "lang_missing_by_mod": missing_by_mod,
    }


def make_patchouli_sheet(wb: Workbook, root: Path, cfg: dict) -> dict:
    ws = wb.create_sheet("patchouli_diff")
    headers = [
        "mod_folder", "modid", "book_id",
        "relpath_en", "json_pointer",
        "en_text", "ja_text(existing)", "ja_auto_google", "ja_text(new)",
        "status", "needs_translation",
        "placeholder_mismatch", "patchouli_markup_mismatch",
    ]
    ws.append(headers)
    style_header(ws, len(headers))

    en_files = sorted(root.rglob("assets/*/patchouli_books/*/en_us/**/*.json"))
    total = len(en_files)
    print(f"[INFO] Patchouli en_us json 件数: {total}", flush=True)

    status_counter: Counter = Counter()
    ph_mismatch_count = 0
    token_mismatch_count = 0
    missing_by_mod: Counter = Counter()
    rows_data: list[dict] = []

    for i, en_path in enumerate(en_files, start=1):
        print_progress(i, total, f"[patchouli] {en_path.name}")

        parsed = parse_patchouli_en_path(root, en_path)
        if not parsed:
            continue
        modid, bookid, relpath_en, rel_ja = parsed

        rel = en_path.relative_to(root)
        parts = rel.parts
        aidx = parts.index("assets")
        mod_folder = str(root.joinpath(*parts[:aidx]))

        en_obj = load_json_any(en_path)
        if en_obj is None:
            continue
        ja_path = root / rel_ja
        ja_obj = load_json_any(ja_path) if ja_path.exists() else None

        ja_lookup = {}
        if ja_obj is not None:
            for ptr, s in iter_patchouli_strings(ja_obj):
                ja_lookup[ptr] = s

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
            if needs:
                missing_by_mod[mod_folder] += 1

            ph_mismatch = set(detect_placeholders(str(en_text))) != set(detect_placeholders(str(ja_text))) and (en_text or ja_text)
            tok_mismatch = Counter(detect_patchouli_tokens(str(en_text))) != Counter(detect_patchouli_tokens(str(ja_text))) and (en_text or ja_text)
            if ph_mismatch:
                ph_mismatch_count += 1
            if tok_mismatch:
                token_mismatch_count += 1

            rows_data.append({
                "mod_folder": mod_folder, "modid": modid, "bookid": bookid,
                "relpath_en": relpath_en, "ptr": ptr,
                "en_text": en_text, "ja_text": ja_text,
                "status": status, "needs": needs,
                "ph_mismatch": ph_mismatch, "tok_mismatch": tok_mismatch,
            })
            status_counter[status] += 1

    translations: dict[str, str] = {}
    if cfg.get("auto_translate", False):
        need_texts = [d["en_text"] for d in rows_data if d["needs"] and str(d["en_text"]).strip()]
        translations = auto_translate_texts(need_texts, label="[patchouli] 自動翻訳")

    for d in rows_data:
        r = append_row_safe(ws, [
            d["mod_folder"], d["modid"], d["bookid"],
            d["relpath_en"], d["ptr"],
            d["en_text"], d["ja_text"], "", "",
            d["status"],
            "TRUE" if d["needs"] else "",
            "TRUE" if d["ph_mismatch"] else "",
            "TRUE" if d["tok_mismatch"] else "",
        ])

        if d["needs"] and str(d["en_text"]).strip():
            auto = translations.get(str(d["en_text"]).strip(), "")
            if auto:
                ws.cell(r, 8).value = auto
            else:
                ws.cell(r, 8).value = f'=IFERROR(GOOGLETRANSLATE(F{r},"en","ja"),"")'

    autosize_columns(ws)
    return {
        "patchouli_status": status_counter,
        "patchouli_placeholder_mismatch": ph_mismatch_count,
        "patchouli_markup_mismatch": token_mismatch_count,
        "patchouli_missing_by_mod": missing_by_mod,
    }


def make_summary_sheet(wb: Workbook, stats: dict, kind: str) -> None:
    """kind: 'lang' または 'patchouli'。該当する集計だけをsummaryシートに書く。"""
    ws = wb.create_sheet("summary")
    ws.append(["metric", "value"])

    def add_counter(title: str, c: Counter, top: int | None = None):
        ws.append([title, ""])
        items = c.most_common(top) if top else c.most_common()
        for k, v in items:
            ws.append([f"  {k}", v])

    if kind == "lang":
        add_counter("lang_status", stats.get("lang_status", Counter()))
        ws.append(["lang_placeholder_mismatch", stats.get("lang_placeholder_mismatch", 0)])
        add_counter("lang_missing_by_mod (top30)", stats.get("lang_missing_by_mod", Counter()), 30)
    else:
        add_counter("patchouli_status", stats.get("patchouli_status", Counter()))
        ws.append(["patchouli_placeholder_mismatch", stats.get("patchouli_placeholder_mismatch", 0)])
        ws.append(["patchouli_markup_mismatch", stats.get("patchouli_markup_mismatch", 0)])
        add_counter("patchouli_missing_by_mod (top30)", stats.get("patchouli_missing_by_mod", Counter()), 30)

    style_header(ws, 2)
    autosize_columns(ws, max_width=60)


def make_meta_sheet(wb: Workbook, root: Path) -> None:
    meta = wb.create_sheet("_meta")
    meta.append(["key", "value"])
    meta.append(["mods_root", str(root.resolve())])
    meta.append(["note", "このシートは編集しないでください(後続ステップが参照します)"])


def run(cfg: dict) -> None:
    if not cfg.get("mods_root"):
        raise SystemExit("[ERROR] config.json に mods_root を指定してください(mod用の設定です)。")
    root = Path(cfg["mods_root"])
    if not root.exists():
        raise SystemExit(f"[ERROR] mods_root が存在しません: {root}")

    print(f"[INFO] mods_root: {root}")

    # --- lang xlsx ---
    out_lang = Path(cfg["xlsx_lang_path"])
    print(f"[INFO] out xlsx (lang): {out_lang}")

    wb_lang = Workbook()
    wb_lang.remove(wb_lang.active)
    make_meta_sheet(wb_lang, root)
    lang_stats = make_lang_sheet(wb_lang, root, cfg)
    make_summary_sheet(wb_lang, lang_stats, "lang")

    out_lang.parent.mkdir(parents=True, exist_ok=True)
    print("[INFO] xlsx (lang) 保存中...", flush=True)
    wb_lang.save(out_lang)
    print(f"[OK] 保存しました: {out_lang}", flush=True)

    # --- patchouli xlsx ---
    if cfg.get("include_patchouli", True):
        out_patchouli = Path(cfg["xlsx_patchouli_path"])
        print(f"[INFO] out xlsx (patchouli): {out_patchouli}")

        wb_patchouli = Workbook()
        wb_patchouli.remove(wb_patchouli.active)
        make_meta_sheet(wb_patchouli, root)
        patchouli_stats = make_patchouli_sheet(wb_patchouli, root, cfg)
        make_summary_sheet(wb_patchouli, patchouli_stats, "patchouli")

        out_patchouli.parent.mkdir(parents=True, exist_ok=True)
        print("[INFO] xlsx (patchouli) 保存中...", flush=True)
        wb_patchouli.save(out_patchouli)
        print(f"[OK] 保存しました: {out_patchouli}", flush=True)

    print("[NEXT] 2つのxlsxをそれぞれGoogleドライブにアップロードしてスプレッドシートとして開き、"
          "ja_jp(new) 列(または空欄なら ja_auto_google の結果)を確認・修正してください。")


if __name__ == "__main__":
    import argparse
    ap = argparse.ArgumentParser(description="mod lang / Patchouli 差分xlsxを生成する")
    ap.add_argument("--config", default="config.json")
    args = ap.parse_args()
    run(load_config(args.config))
