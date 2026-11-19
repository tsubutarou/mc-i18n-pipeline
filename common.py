# -*- coding: utf-8 -*-
"""
mc_i18n_pipeline 共通ユーティリティ

各ステップで共通して使う処理をここに集約。
- JSON読み書き
- 進捗バー表示
- placeholder(%s, %1$s, {0}...) / Patchouli markup($(...)) 検出
- Patchouli本文(JSON)の再帰抽出 / JSON Pointer 操作
- openpyxl のヘッダースタイル・列幅自動調整・ヘッダー→列インデックス解決
- config.json の読み込みと検証
"""
from __future__ import annotations

import json
import re
import secrets
import shutil
import sys
import time
import zipfile
from pathlib import Path
from typing import Any, Iterable, Optional

from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter

# --------------------------------------------------------------------------
# Placeholder / Patchouli markup detection
# --------------------------------------------------------------------------
PH_RE = re.compile(r"%(?:\d+\$)?[sdfoxX]|\{[0-9]+\}")
PATCHOULI_TOKEN_RE = re.compile(r"\$\([^)]+\)")
LANGKEY_LIKE_RE = re.compile(r"^[a-z0-9_.:-]{3,80}$")

# Patchouli JSON内で「翻訳対象の地の文」ではないキー(構造・ID・数値メタ等)
PATCHOULI_EXCLUDE_KEYS = {
    "type", "id", "icon", "item", "block", "entity", "recipe", "advancement",
    "anchor", "link", "url", "category", "entry", "model", "texture",
    "trigger", "flag", "flags", "conditions", "condition", "criteria",
    "lore", "sound", "chapter", "chapters",
    "x", "y", "z", "u", "v", "w", "h", "width", "height", "scale",
    "timer", "time", "delay", "speed",
}


def detect_placeholders(s: str) -> list[str]:
    return PH_RE.findall(s or "")


def detect_patchouli_tokens(s: str) -> list[str]:
    return PATCHOULI_TOKEN_RE.findall(s or "")


def looks_like_lang_key(s: str) -> bool:
    """Patchouli本文が翻訳キー文字列そのものを指している場合(地の文ではない)を除外する"""
    if not s:
        return False
    ss = s.strip()
    if " " in ss or "\n" in ss or "\t" in ss:
        return False
    if not LANGKEY_LIKE_RE.match(ss):
        return False
    return ("." in ss) and any("a" <= c <= "z" for c in ss.lower())


# --------------------------------------------------------------------------
# JSON I/O
# --------------------------------------------------------------------------
def load_json_dict(p: Path) -> dict:
    if not p.exists():
        return {}
    try:
        text = p.read_text(encoding="utf-8-sig")
    except Exception as e:
        print(f"[WARN] JSON読み込み失敗: {p} ({e})", file=sys.stderr)
        return {}
    data = _parse_json_lenient(text, p)
    if not isinstance(data, dict):
        return {}
    return {str(k): ("" if v is None else str(v)) for k, v in data.items()}


def load_json_any(p: Path) -> Any:
    try:
        text = p.read_text(encoding="utf-8-sig")
    except Exception:
        return None
    return _parse_json_lenient(text, p)


_TRAILING_COMMA_RE = re.compile(r",(\s*[}\]])")


def _parse_json_lenient(text: str, p: Path) -> Any:
    """
    通常のjson.loadsに失敗した場合、末尾カンマ(よくあるmod側の軽微な構文ミス)
    だけ補正して再トライする。それでも失敗する場合(コメント混入、生の制御文字、
    エスケープ不正など)は諦めてNoneを返し、元のエラーを警告として出す。
    """
    try:
        return json.loads(text)
    except Exception as e:
        repaired = _TRAILING_COMMA_RE.sub(r"\1", text)
        if repaired != text:
            try:
                data = json.loads(repaired)
                print(f"[INFO] 末尾カンマを補正して読み込みました: {p}", file=sys.stderr)
                return data
            except Exception:
                pass
        print(f"[WARN] JSON読み込み失敗: {p} ({e})", file=sys.stderr)
        return None


def safe_write_json(p: Path, obj: Any) -> None:
    p.parent.mkdir(parents=True, exist_ok=True)
    with p.open("w", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False, indent=2)


# --------------------------------------------------------------------------
# JSON Pointer (RFC 6901) - Patchouli本文の書き戻しに使用
# --------------------------------------------------------------------------
def _ptr_escape(token: str) -> str:
    return token.replace("~", "~0").replace("/", "~1")


def _ptr_unescape(token: str) -> str:
    return token.replace("~1", "/").replace("~0", "~")


def ptr_split(ptr: str) -> list[str]:
    if not ptr or ptr == "/":
        return []
    if not ptr.startswith("/"):
        raise ValueError(f"json_pointer must start with '/': {ptr}")
    return [_ptr_unescape(p) for p in ptr.split("/")[1:]]


def ptr_get(obj: Any, ptr: str) -> Any:
    cur = obj
    for t in ptr_split(ptr):
        if isinstance(cur, list):
            cur = cur[int(t)]
        elif isinstance(cur, dict):
            cur = cur.get(t)
        else:
            return None
    return cur


def ptr_set(obj: Any, ptr: str, value: Any) -> None:
    tokens = ptr_split(ptr)
    if not tokens:
        raise ValueError("cannot set root pointer")
    cur = obj
    for t in tokens[:-1]:
        if isinstance(cur, list):
            cur = cur[int(t)]
        else:
            if t not in cur or not isinstance(cur[t], (dict, list)):
                cur[t] = {}
            cur = cur[t]
    last = tokens[-1]
    if isinstance(cur, list):
        cur[int(last)] = value
    else:
        cur[last] = value


def iter_patchouli_strings(obj: Any, base_ptr: str = ""):
    """Patchouli本のJSONを再帰的に走査し、(json_pointer, 文字列) を列挙する"""
    if isinstance(obj, dict):
        for k, v in obj.items():
            if isinstance(k, str) and k in PATCHOULI_EXCLUDE_KEYS:
                continue
            ptr = base_ptr + "/" + _ptr_escape(str(k))
            yield from iter_patchouli_strings(v, ptr)
    elif isinstance(obj, list):
        for i, v in enumerate(obj):
            yield from iter_patchouli_strings(v, base_ptr + f"/{i}")
    elif isinstance(obj, str):
        s = obj.strip()
        if not s or looks_like_lang_key(s):
            return
        yield (base_ptr, obj)


def parse_patchouli_en_path(root: Path, p: Path):
    """
    p: .../assets/<modid>/patchouli_books/<bookid>/en_us/... 形式のパス。
    戻り値: (modid, bookid, relpath(en基準, posix), relpath(ja基準))、パターン不一致ならNone
    """
    try:
        rel = p.relative_to(root)
    except ValueError:
        return None
    parts = list(rel.parts)
    try:
        aidx = parts.index("assets")
        modid = parts[aidx + 1]
        if parts[aidx + 2] != "patchouli_books":
            return None
        bookid = parts[aidx + 3]
        if parts[aidx + 4] != "en_us":
            return None
    except (ValueError, IndexError):
        return None

    parts_ja = parts[:]
    parts_ja[aidx + 4] = "ja_jp"
    rel_ja = Path(*parts_ja).as_posix()
    return modid, bookid, rel.as_posix(), rel_ja


# --------------------------------------------------------------------------
# 進捗表示
# --------------------------------------------------------------------------
def print_progress(i: int, total: int, label: str = "") -> None:
    pct = int(100 * i / total) if total else 100
    bar_len = 28
    fill = int(bar_len * pct / 100)
    bar = "#" * fill + " " * (bar_len - fill)
    sys.stdout.write(f"\r[{bar}] {pct:3d}% {i}/{total} {label}")
    sys.stdout.flush()
    if i >= total:
        print()


# --------------------------------------------------------------------------
# openpyxl ヘルパー
# --------------------------------------------------------------------------
def style_header(ws, col_count: int) -> None:
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="4F81BD")
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for c in range(1, col_count + 1):
        cell = ws.cell(1, c)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align
    ws.freeze_panes = "A2"


def autosize_columns(ws, max_width: int = 80, scan_cap: int = 5000) -> None:
    for col in range(1, ws.max_column + 1):
        letter = get_column_letter(col)
        max_len = 10
        for row in range(1, min(ws.max_row, scan_cap) + 1):
            v = ws.cell(row, col).value
            if v is None:
                continue
            max_len = max(max_len, len(str(v)) + 2)
        ws.column_dimensions[letter].width = min(max_width, max_len)


def build_header_index(ws, cols: dict[str, str]) -> dict[str, int]:
    header = [norm(c.value).strip() for c in ws[1]]
    idx = {}
    for key, col_name in cols.items():
        if col_name not in header:
            raise SystemExit(f"列が見つかりません: {col_name} (sheet={ws.title})")
        idx[key] = header.index(col_name) + 1
    return idx


def norm(v: Any) -> str:
    return "" if v is None else str(v)


def pick_ja(ws, r: int, idx: dict[str, int]) -> str:
    """
    採用ルール: ja_jp(new) > ja_jp(existing) > ja_auto_google

    ただし ja_jp(existing) が en_us と完全一致する場合(=modが未翻訳のまま
    en_us をコピーしただけのプレースホルダー ja_jp.json を同梱しているケース)は
    「訳ではない」とみなしてスキップし、ja_auto_google側にフォールバックする。
    idx に "en" キーが無い場合(呼び出し側が未対応)はこのチェックをスキップする。
    """
    ja_new = norm(ws.cell(r, idx["ja_new"]).value).strip()
    if ja_new:
        return ja_new

    ja_old = norm(ws.cell(r, idx["ja_old"]).value).strip()
    if ja_old:
        en_col = idx.get("en")
        en_val = norm(ws.cell(r, en_col).value).strip() if en_col else ""
        if not (en_val and ja_old == en_val):
            return ja_old

    ja_auto = norm(ws.cell(r, idx["ja_auto"]).value).strip()
    return ja_auto


FORMULA_TRIGGER_CHARS = ("=", "+", "-", "@")


def guard_formula_injection(cell) -> None:
    """
    openpyxlは文字列セルの値が "=" で始まると自動的に数式(formula)として保存してしまい、
    実際にExcel/Google Sheetsで一度も計算されていないと、再度openpyxlで読み込んだ時に
    値が None になって消えてしまう(silent data loss)。
    また "+" "-" "@" で始まる文字列もGoogle Sheets側の自動解釈で数式扱いされることがある。
    mod の en_us.json / 既存の ja_jp.json には、たまたまこれらの文字で始まる文言が
    含まれることがあるため、該当するセルは明示的に文字列型として保存し直す。
    ※ 意図的に数式を入れているセル(GOOGLETRANSLATE式など)には適用しないこと。
    """
    v = cell.value
    if isinstance(v, str) and v.startswith(FORMULA_TRIGGER_CHARS):
        cell.data_type = "s"


def append_row_safe(ws, values: list) -> int:
    """
    ws.append(values) した上で、文字列セルに対して guard_formula_injection を適用する。
    GOOGLETRANSLATE式などの意図的な数式は、この呼び出しの後で別途セットすること。
    戻り値は追加された行番号。
    """
    ws.append(values)
    r = ws.max_row
    for c, v in enumerate(values, start=1):
        if isinstance(v, str):
            guard_formula_injection(ws.cell(r, c))
    return r


def sanitize_name(name: str) -> str:
    name = (name or "").strip()
    name = re.sub(r'[<>:"/\\|?*\x00-\x1F]', "_", name)
    return name or "UNKNOWN"


# --------------------------------------------------------------------------
# 単体パック(単一のリソースパック/データパック。zipまたは展開済みフォルダ)の解決
# 05_scan_single_packs.py / 06_apply_single_packs.py で使用。
# --------------------------------------------------------------------------
def is_safe_zip_member(name: str) -> bool:
    """zip slip対策: 絶対パス・親ディレクトリ参照を含むメンバーは無視する"""
    if name.startswith("/") or name.startswith("\\"):
        return False
    if ".." in Path(name).parts:
        return False
    return True


def resolve_single_pack_roots(cfg: dict) -> list[tuple[str, Path]]:
    """
    cfg["single_packs_dir"] 直下の各アイテムを1つの独立したパックとして解決する。

    - サブフォルダ  : 既に展開済みのパックとしてそのまま使う
    - .zip ファイル : out_dir/_single_packs_src/<pack_name>/ に展開してから使う
                       (zipより新しく再展開が必要な場合のみ再展開、それ以外はスキップ)

    戻り値: [(pack_name, pack_root), ...] を名前順で返す。pack_root直下に
    pack.mcmeta や assets/ data/ がある想定(modpack一式のmods_rootとは別物)。
    """
    src_dir = Path(cfg.get("single_packs_dir") or "")
    if not cfg.get("single_packs_dir"):
        raise SystemExit("[ERROR] config.json に single_packs_dir (単体パック群のフォルダ) を指定してください。")
    if not src_dir.exists():
        raise SystemExit(f"[ERROR] single_packs_dir が存在しません: {src_dir}")

    cache_dir = Path(cfg["out_dir"]) / "_single_packs_src"
    cache_dir.mkdir(parents=True, exist_ok=True)

    # out_dir が single_packs_dir の中にある場合(例: single_packs\_out)、それ自体はパックではないので除外する
    out_resolved = Path(cfg["out_dir"]).resolve()
    items = sorted(i for i in src_dir.iterdir() if i.resolve() != out_resolved)
    print(f"[INFO] single_packs_dir 走査中: {src_dir} ({len(items)}件)", flush=True)

    roots: list[tuple[str, Path]] = []
    for item in items:
        if item.is_dir():
            roots.append((sanitize_name(item.name), item))
            continue
        if item.is_file() and item.suffix.lower() == ".zip":
            pack_name = sanitize_name(item.stem)
            dest = cache_dir / pack_name
            marker = dest / ".extracted_from.txt"
            need_extract = True
            if dest.exists() and marker.exists():
                try:
                    if marker.stat().st_mtime >= item.stat().st_mtime:
                        need_extract = False
                except OSError:
                    pass
            if not need_extract:
                print(f"[INFO] {item.name} は展開済みのためスキップ", flush=True)
                roots.append((pack_name, dest))
                continue

            if dest.exists():
                shutil.rmtree(dest)
            dest.mkdir(parents=True, exist_ok=True)
            try:
                with zipfile.ZipFile(item) as zf:
                    members = [m for m in zf.infolist() if not m.is_dir() and is_safe_zip_member(m.filename)]
                    total = len(members)
                    size_mb = item.stat().st_size / (1024 * 1024)
                    print(f"[INFO] {item.name} を展開中... ({total}ファイル, {size_mb:.1f}MB)", flush=True)
                    for i, m in enumerate(members, start=1):
                        print_progress(i, total, item.name)
                        zf.extract(m, dest)
                marker.write_text(str(item.resolve()), encoding="utf-8")
            except zipfile.BadZipFile:
                print(f"[WARN] 壊れたzipのためスキップ: {item.name}", file=sys.stderr)
                continue
            roots.append((pack_name, dest))
    return roots


# --------------------------------------------------------------------------
# config.json
# --------------------------------------------------------------------------
DEFAULT_CONFIG = {
    "mods_jar_dir": "",
    "mods_root": "",
    "out_dir": "",
    "xlsx_lang_name": "ALL_mods_lang_diff.xlsx",
    "xlsx_patchouli_name": "ALL_mods_patchouli_diff.xlsx",
    "include_patchouli": True,
    "pack_name_lang": "JP_LangTranslations",
    "pack_name_patchouli": "JP_PatchouliTranslations",
    "lang_dir_name": "lang",
    "pack_format": 15,
    "pack_description": "Auto-generated Japanese translation pack",
    "zip_output": True,
    "exclude_modids": ["minecraft"],
    "single_packs_dir": "",
    "single_packs_xlsx_name": "single_packs_diff.xlsx",
    "auto_translate": False,
}


# --------------------------------------------------------------------------
# 自動翻訳(deep_translator経由、Google翻訳のWeb版を利用。APIキー不要)
# config.jsonの auto_translate=false、またはdeep_translatorが未インストールの
# 場合は何もせず空dictを返す。呼び出し側はGOOGLETRANSLATE数式へのフォールバックを
# 自前で用意すること(01_build_xlsx.py / 05_scan_single_packs.py 参照)。
# --------------------------------------------------------------------------
_TRANSLATE_CACHE: dict[str, str] = {}
_TRANSLATOR = None
_TRANSLATOR_UNAVAILABLE = False
_AUTO_TRANSLATE_DISABLED = False  # 繰り返し失敗した場合、以降のプロセス全体で自動翻訳自体を諦める
_NEXT_ALLOWED_TS = 0.0          # プロセス全体で共有するリクエスト間隔管理(呼び出しをまたいで有効)
_CONSECUTIVE_RATE_LIMIT_ITEMS = 0  # 直近で何バッチ連続してレート制限に完全敗北したか(cool-down判定用)
_COOLDOWN_CYCLES = 0             # cool-downを挟んでもなお失敗が続いた回数(諦め判定用)
_BATCH_SIZE = 25                 # 1リクエストにまとめる文字列数の上限
_BATCH_MAX_CHARS = 2500          # 1リクエストの合計文字数の上限(超えたら別バッチに分ける)


def _get_translator():
    global _TRANSLATOR, _TRANSLATOR_UNAVAILABLE
    if _TRANSLATOR is not None:
        return _TRANSLATOR
    if _TRANSLATOR_UNAVAILABLE:
        return None
    try:
        from deep_translator import GoogleTranslator
        _TRANSLATOR = GoogleTranslator(source="en", target="ja")
        return _TRANSLATOR
    except Exception as e:
        print(f"[WARN] deep_translator が使えません({e})。自動翻訳をスキップし、"
              f"GOOGLETRANSLATE数式にフォールバックします。"
              f"(pip install deep-translator でインストールできます)", file=sys.stderr)
        _TRANSLATOR_UNAVAILABLE = True
        return None


def _throttle(min_interval: float) -> None:
    """
    直前の翻訳リクエストから min_interval 秒空くまで待つ。単体パックを大量に
    処理する場合など、auto_translate_texts が何度も呼ばれるケースでも、
    プロセス全体で見て一定間隔以上あける(呼び出し単位ではなくグローバルに管理)。
    """
    global _NEXT_ALLOWED_TS
    now = time.monotonic()
    if now < _NEXT_ALLOWED_TS:
        time.sleep(_NEXT_ALLOWED_TS - now)
        now = time.monotonic()
    _NEXT_ALLOWED_TS = now + min_interval


def _make_batches(texts: list[str], batch_size: int, max_chars: int) -> list[list[str]]:
    """翻訳対象を「1リクエストにまとめて送る単位」に、件数と合計文字数の両方の上限で分割する。"""
    batches: list[list[str]] = []
    cur: list[str] = []
    cur_chars = 0
    for t in texts:
        t_len = len(t)
        if cur and (len(cur) >= batch_size or cur_chars + t_len > max_chars):
            batches.append(cur)
            cur = []
            cur_chars = 0
        cur.append(t)
        cur_chars += t_len
    if cur:
        batches.append(cur)
    return batches


def _translate_batch_once(translator, texts: list[str]) -> Optional[list[str]]:
    """
    texts をまとめて1回のリクエストで翻訳する。1件なら単純にtranslate()し、
    2件以上ならランダムな区切りトークンで1つの文字列に結合して送り、返ってきた
    訳文を同じトークンで分割し直す。結合前後で件数が一致しない(=結合/分割が
    どこかで壊れた)場合はNoneを返し、呼び出し側で1件ずつのフォールバックに
    切り替えさせる(データが壊れたまま採用されることはない)。
    """
    if len(texts) == 1:
        return [translator.translate(texts[0]) or ""]

    sep = f"@@{secrets.token_hex(4)}@@"
    joined = f"\n{sep}\n".join(texts)
    translated = translator.translate(joined) or ""
    parts = [p.strip() for p in re.split(re.escape(sep), translated, flags=re.IGNORECASE)]
    if len(parts) != len(texts):
        return None
    return parts


def auto_translate_texts(texts: Iterable[str], label: str = "自動翻訳", delay: float = 0.5) -> dict[str, str]:
    """
    重複を除いたテキスト群をまとめてEN→JAに自動翻訳して {原文: 訳文} を返す。
    失敗した/翻訳できなかった原文は結果dictに含めない(呼び出し側でフォールバック判定)。
    プロセス内キャッシュを持つため、複数回呼んでも同じ文字列は1回しか翻訳しない。

    数万件規模のmodpackでも現実的な時間で終えられるよう、複数の文字列を
    区切りトークンで結合して1リクエストにまとめて送る(_make_batches /
    _translate_batch_once)。結合・分割がうまくいかなかったバッチだけ、
    その場で1件ずつのリクエストにフォールバックする。

    Google翻訳(Web版)は5リクエスト/秒程度でレート制限されるため、リクエスト間隔を
    プロセス全体で管理しつつ(_throttle)、"too many requests" を検知した場合は
    その場で指数的に待ってリトライする。さらに、cool-downを挟んでもなお失敗が続く
    場合(IP自体がブロックされている可能性が高い)は、無駄にリトライし続けて
    処理が事実上止まってしまうのを避けるため、それ以降はこのプロセス内での
    自動翻訳自体を諦め、即座にGOOGLETRANSLATE数式フォールバックへ切り替える
    (_AUTO_TRANSLATE_DISABLED)。
    """
    global _CONSECUTIVE_RATE_LIMIT_ITEMS, _COOLDOWN_CYCLES, _AUTO_TRANSLATE_DISABLED

    uniq = sorted({t.strip() for t in texts if t and str(t).strip()})
    if not uniq:
        return {}

    if _AUTO_TRANSLATE_DISABLED:
        return {}

    translator = _get_translator()
    if translator is None:
        return {}

    result: dict[str, str] = {}
    todo = [t for t in uniq if t not in _TRANSLATE_CACHE]
    for t in uniq:
        if t in _TRANSLATE_CACHE and _TRANSLATE_CACHE[t]:
            result[t] = _TRANSLATE_CACHE[t]

    total = len(todo)
    if total == 0:
        return result

    batches = _make_batches(todo, _BATCH_SIZE, _BATCH_MAX_CHARS)
    print(f"[INFO] 自動翻訳対象(ユニーク文字列, キャッシュ済み除く): {total}件 → {len(batches)}バッチ", flush=True)

    done = 0
    for batch in batches:
        if _AUTO_TRANSLATE_DISABLED:
            break

        if _CONSECUTIVE_RATE_LIMIT_ITEMS >= 3:
            _COOLDOWN_CYCLES += 1
            if _COOLDOWN_CYCLES >= 2:
                print(
                    "\n[WARN] cool-downを挟んでも自動翻訳が失敗し続けるため、"
                    "自動翻訳先(Google翻訳)がブロックされている可能性が高いです。"
                    "これ以降は今回の実行中、自動翻訳を諦めてGOOGLETRANSLATE数式に"
                    "フォールバックします。config.jsonのauto_translateをfalseにして"
                    "Googleスプレッドシート運用に切り替えるか、しばらく時間を置いてから"
                    "再実行することをおすすめします。",
                    file=sys.stderr,
                )
                _AUTO_TRANSLATE_DISABLED = True
                break
            cooldown = 60
            print(f"\n[WARN] レート制限が続いています。{cooldown}秒間クールダウンします...", file=sys.stderr)
            time.sleep(cooldown)
            _CONSECUTIVE_RATE_LIMIT_ITEMS = 0

        batch_result: Optional[list[str]] = None
        batch_rate_limited = False
        for attempt in range(4):
            _throttle(delay)
            try:
                batch_result = _translate_batch_once(translator, batch)
                break
            except Exception as e:
                msg = str(e).lower()
                is_rate_limit = "too many requests" in msg or " 429" in msg or msg.startswith("429")
                batch_rate_limited = batch_rate_limited or is_rate_limit
                if is_rate_limit and attempt < 3:
                    wait = min(4 * (attempt + 1), 30)
                    print(f"\n[WARN] レート制限を検知、{wait:.0f}秒待機してリトライします...", file=sys.stderr)
                    time.sleep(wait)
                    continue
                if not is_rate_limit and attempt < 1:
                    time.sleep(1.0)
                    continue
                print(f"\n[WARN] バッチ翻訳失敗({len(batch)}件): {e}", file=sys.stderr)
                break

        if batch_result is None and not batch_rate_limited:
            # 結合/分割が壊れたケース。このバッチだけ1件ずつ翻訳し直す
            # (レート制限で丸ごと失敗した場合は1件ずつやっても結果は同じなので、
            # 無駄にリクエストを増やさないためここではやらない)。
            print(f"\n[INFO] バッチの結合/分割が一致しなかったため、{len(batch)}件を1件ずつ翻訳し直します", file=sys.stderr)
            batch_result = []
            for t in batch:
                translated_one = ""
                for attempt in range(2):
                    _throttle(delay)
                    try:
                        translated_one = translator.translate(t) or ""
                        break
                    except Exception:
                        if attempt == 0:
                            time.sleep(1.0)
                batch_result.append(translated_one)

        if batch_result is None:
            batch_result = ["" for _ in batch]

        any_success = False
        for t, tr in zip(batch, batch_result):
            _TRANSLATE_CACHE[t] = tr
            if tr:
                result[t] = tr
                any_success = True

        if any_success:
            _CONSECUTIVE_RATE_LIMIT_ITEMS = 0
            _COOLDOWN_CYCLES = 0
        elif batch_rate_limited:
            _CONSECUTIVE_RATE_LIMIT_ITEMS += 1

        done += len(batch)
        print_progress(done, total, label)

    return result


def _fix_windows_backslashes(raw: str) -> str:
    r"""
    JSON文字列リテラル内で、Windowsパスの "\" がエスケープされずに書かれていても
    読めるように補正する。既に正しくエスケープされている "\\" はそのまま扱い、
    有効なJSONエスケープ(\" \\ \/ \b \f \n \r \t \uXXXX)以外の単独 "\" だけを
    "\\\\" に補正する。(regexの使い捨て置換だと既存の "\\" を壊すことがあるため、
    1文字ずつ走査して処理する)
    """
    out = []
    i = 0
    n = len(raw)
    in_string = False
    # \b \f \n \r \t は config の値(パス・名前)ではまず使われず、逆に "C:\Users\taro"
    # の "\t" のように Windows パスに普通に現れるので、エスケープとしては扱わない。
    # \u も直後が16進4桁のときだけエスケープとみなす("C:\users" 等を壊さないため)。
    valid_escapes = set('\\/"')
    hexdigits = set("0123456789abcdefABCDEF")
    while i < n:
        c = raw[i]
        if in_string:
            if c == "\\":
                nxt = raw[i + 1] if i + 1 < n else ""
                is_unicode_escape = nxt == "u" and len(raw) >= i + 6 and all(h in hexdigits for h in raw[i + 2:i + 6])
                if nxt in valid_escapes or is_unicode_escape:
                    out.append(c)
                    out.append(nxt)
                    i += 2
                    continue
                out.append("\\\\")
                i += 1
                continue
            if c == '"':
                in_string = False
            out.append(c)
            i += 1
            continue
        if c == '"':
            in_string = True
        out.append(c)
        i += 1
    return "".join(out)


def load_config(path: str) -> dict:
    p = Path(path)
    if not p.exists():
        raise SystemExit(
            f"[ERROR] config.json が見つかりません: {p}\n"
            f"        config.example.json をコピーして作成してください。"
        )
    raw = p.read_text(encoding="utf-8-sig")
    raw = _fix_windows_backslashes(raw)
    try:
        user_cfg = json.loads(raw)
    except json.JSONDecodeError as e:
        raise SystemExit(f"[ERROR] config.json の書式が壊れています: {e}")

    cfg = dict(DEFAULT_CONFIG)
    cfg.update(user_cfg)

    if not cfg.get("out_dir"):
        if cfg.get("mods_root"):
            cfg["out_dir"] = str(Path(cfg["mods_root"]) / "_out")
        elif cfg.get("single_packs_dir"):
            cfg["out_dir"] = str(Path(cfg["single_packs_dir"]) / "_out")
        else:
            raise SystemExit(
                "[ERROR] config.json に out_dir (またはmods_root/single_packs_dirのいずれか) を指定してください。"
            )

    if cfg.get("mods_root"):
        cfg["mods_root"] = str(Path(cfg["mods_root"]))
    cfg["out_dir"] = str(Path(cfg["out_dir"]))
    cfg["xlsx_lang_path"] = str(Path(cfg["out_dir"]) / cfg["xlsx_lang_name"])
    cfg["xlsx_patchouli_path"] = str(Path(cfg["out_dir"]) / cfg["xlsx_patchouli_name"])
    cfg["single_packs_xlsx_path"] = str(Path(cfg["out_dir"]) / cfg["single_packs_xlsx_name"])
    return cfg
