# -*- coding: utf-8 -*-
"""
Step 0（新規・任意）: mods_jar_dir 内の .jar を mods_root 以下に自動展開する。

これまで「modフォルダ群を手動で1つずつ右クリック→展開...してmods_rootに置く」
という作業が必要だったが、それを自動化するステップ。

  - 対象: mods_jar_dir 直下の *.jar（サブフォルダは見ない）
  - 出力先: mods_root/<jarのファイル名(拡張子抜き)>/assets/...
  - 高速化のため、jar内の assets/ 配下だけを展開する
    （01_build_xlsx.py が見るのは assets/*/lang/*.json と
    assets/*/patchouli_books/... だけなので、class等の他ファイルは不要）
  - 既に展開済み（出力フォルダがあり、jarより新しい）ならスキップする。
    jarを更新した後などに再展開したい場合は --force を付ける。
"""
from __future__ import annotations

import zipfile
from pathlib import Path

from common import load_config, print_progress, is_safe_zip_member


def extract_one_jar(jar_path: Path, dest_root: Path, force: bool) -> tuple[str, int]:
    """1つのjarを展開する。戻り値: (status, 展開ファイル数)

    status: "ok" | "skip"(既展開でスキップ) | "no_assets"(assets/を含まない=
            ライブラリ系modの可能性) | "bad_zip"(壊れている/jarでない)
    """
    dest = dest_root / jar_path.stem
    marker = dest / ".extracted_from.txt"

    if dest.exists() and not force:
        try:
            done_mtime = marker.stat().st_mtime if marker.exists() else dest.stat().st_mtime
            if done_mtime >= jar_path.stat().st_mtime:
                return "skip", 0
        except OSError:
            pass

    try:
        with zipfile.ZipFile(jar_path) as zf:
            members = [
                m for m in zf.infolist()
                if m.filename.startswith("assets/")
                and not m.filename.endswith("/")
                and is_safe_zip_member(m.filename)
            ]
            if not members:
                return "no_assets", 0
            dest.mkdir(parents=True, exist_ok=True)
            for m in members:
                zf.extract(m, dest)
            marker.write_text(str(jar_path.resolve()), encoding="utf-8")
            return "ok", len(members)
    except zipfile.BadZipFile:
        return "bad_zip", 0


def run(cfg: dict, force: bool = False) -> None:
    jar_dir = cfg.get("mods_jar_dir")
    if not jar_dir:
        raise SystemExit(
            "[ERROR] config.json に mods_jar_dir (.jarが入っているフォルダ) を指定してください。"
        )
    if not cfg.get("mods_root"):
        raise SystemExit("[ERROR] config.json に mods_root を指定してください(mod用の設定です)。")
    jar_dir = Path(jar_dir)
    if not jar_dir.exists():
        raise SystemExit(f"[ERROR] mods_jar_dir が存在しません: {jar_dir}")

    dest_root = Path(cfg["mods_root"])
    dest_root.mkdir(parents=True, exist_ok=True)

    jars = sorted(jar_dir.glob("*.jar"))
    total = len(jars)
    print(f"[INFO] jar 件数    : {total}")
    print(f"[INFO] mods_jar_dir: {jar_dir}")
    print(f"[INFO] mods_root   : {dest_root}")
    if total == 0:
        print("[WARN] .jar が見つかりませんでした。mods_jar_dir の指定を確認してください。")
        return

    counts = {"ok": 0, "skip": 0, "no_assets": 0, "bad_zip": 0}
    bad_zips: list[str] = []
    no_asset_jars: list[str] = []

    for i, jar_path in enumerate(jars, start=1):
        print_progress(i, total, jar_path.name)
        status, _n = extract_one_jar(jar_path, dest_root, force)
        counts[status] += 1
        if status == "bad_zip":
            bad_zips.append(jar_path.name)
        elif status == "no_assets":
            no_asset_jars.append(jar_path.name)

    print(
        f"[OK] 展開: {counts['ok']} 件 / スキップ(既展開): {counts['skip']} 件 / "
        f"assetsなし: {counts['no_assets']} 件 / 壊れたzip: {counts['bad_zip']} 件"
    )
    if bad_zips:
        shown = ", ".join(bad_zips[:20]) + (" ..." if len(bad_zips) > 20 else "")
        print(f"[WARN] 展開できなかったjar({len(bad_zips)}件): {shown}")
    if no_asset_jars:
        shown = ", ".join(no_asset_jars[:20]) + (" ..." if len(no_asset_jars) > 20 else "")
        print(f"[INFO] assets/ を含まないjar({len(no_asset_jars)}件、ライブラリ系modの可能性): {shown}")

    print("[NEXT] 続けて .\\run.ps1 -Step xlsx を実行してください。")


if __name__ == "__main__":
    import argparse

    ap = argparse.ArgumentParser(description="mods_jar_dir 内の.jarをmods_root以下に展開する")
    ap.add_argument("--config", default="config.json")
    ap.add_argument("--force", action="store_true", help="既に展開済みでも強制的に再展開する")
    args = ap.parse_args()
    run(load_config(args.config), force=args.force)
