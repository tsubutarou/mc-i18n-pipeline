# 変更履歴

## 1.0.0 (2026-11-19)

最初の公開版。

- modpack本体: `.jar` の自動展開 → 未翻訳差分の xlsx 生成 → 衝突チェック → 日本語化リソースパック生成
- lang 本文と Patchouli 本文で xlsx・リソースパックを分けて生成
- 単体パック（互換パック等）を丸ごと日本語版にする `run_rp` ワークフロー
- 自動翻訳（`auto_translate`）: deep-translator 経由でまとめて翻訳。失敗時は `GOOGLETRANSLATE` 数式にフォールバック
- `ja_jp.json` が実質 `en_us.json` のコピーになっているケースを `SAME_AS_EN` として検出
- `=` 等で始まる文言が xlsx で数式と誤認されないよう、文字列として保存
- config の Windows パスの `\` がエスケープされていなくても読めるように
- 診断ステップ（xlsx の行ズレ・欠落 modid の検出）
- ダブルクリックで使える `setup.bat` / `run.bat` / `run_rp.bat`
