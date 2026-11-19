# mc-i18n-pipeline

Minecraft の mod / リソースパックを**まとめて日本語化**するツールです（Windows 向け）。

modpack の `.jar` をまとめて読んで「まだ日本語訳が無い文」だけを Excel（xlsx）に一覧にし、
自動翻訳を入れ、必要なところだけ手直しして、**日本語化リソースパック**として書き出します。

- 数百 mod 規模の modpack でも、未翻訳の文だけを 1 つの xlsx に集約
- 自動翻訳（Google 翻訳）入り。気になる行だけ手直しすれば OK
- mod 内の説明書（Patchouli の本）の本文にも対応
- 互換パックなど「単体で完結したリソースパック / データパック」の日本語化にも対応
- `%s` や `$(...)` などの書式記号が訳文で崩れていないかをチェック

## 必要なもの

- Windows 10 / 11
- [Python](https://www.python.org/) 3.10 以上（インストール時に「Add python.exe to PATH」にチェック推奨）

## はじめかた

1. このページの **Code → Download ZIP** でダウンロードして、好きな場所に展開する
2. `setup.bat` をダブルクリック（必要なパッケージのインストールと、設定ファイル `config.json` の作成をします）
3. `config.json` をメモ帳で開き、フォルダのパスを自分の環境に書き換える
   - `mods_jar_dir` … modpack の `mods` フォルダ（`.jar` が入っているところ）
   - `mods_root` / `out_dir` … 作業用・出力用のフォルダ（無ければ自動で作られます）
4. `run.bat` をダブルクリック
5. `out_dir` にできた `JP_LangTranslations.zip`（と `JP_PatchouliTranslations.zip`）を
   Minecraft の `resourcepacks` フォルダに入れて有効化する

> パスの `\` は `\\` と書くのが JSON の正式な書き方ですが、`\` 1 つのままでも読めるようにしてあります。

訳を手直ししたい場合は、下の「modpack本体」の手順のように `xlsx` と `pack` を分けて実行してください。

---

modpack本体（複数modをまとめて日本語化）と単体パック（互換パック等、
それ単体で完結したリソースパック/データパック）は置き場所も処理も別物なので、
コマンドを分けています:

| 対象 | config | コマンド |
|---|---|---|
| modpack本体（複数modフォルダ） | `config.json` | `.\run.ps1`（または `run.bat`） |
| 単体パック（互換パック等） | `config.resourcepack.json` | `.\run_rp.ps1`（または `run_rp.bat`） |

どちらも「差分xlsx生成 → 翻訳 → リソースパック生成」という流れは同じです。

## modpack本体（`.\run.ps1`）

複数modの`.jar`をまとめて展開 → 差分xlsx生成 → リソースパック生成、を自動化します。
lang本文とPatchouli本文は不具合の出方が違うため、xlsx・リソースパックとも別ファイルです。

### config.json の項目

| キー | 説明 |
|---|---|
| `mods_jar_dir` | `.jar`が入っているフォルダ（`.minecraft\mods`など） |
| `mods_root` | 展開済みmodフォルダの置き場所（`mods_jar_dir`から自動展開される） |
| `out_dir` | 出力先（省略時は `mods_root/_out`） |
| `include_patchouli` | Patchouli本文も対象にするか（falseならPatchouli関連は一切生成しない） |
| `pack_name_lang` / `pack_name_patchouli` | 生成するリソースパック名 |
| `pack_format` | `pack.mcmeta` に書く値。既定の `15` は Minecraft 1.20.1 向け。他のバージョンは [Minecraft Wiki](https://minecraft.wiki/w/Pack_format) を参照 |
| `exclude_modids` | パックから除外するmodid（既定`["minecraft"]`） |
| `lang_dir_name` | 言語ファイルのフォルダ名（通常`lang`のまま） |
| `zip_output` | パックをzip化するか |
| `auto_translate` | 実行時にPythonから自動翻訳するか（既定`false`=Googleスプレッドシート運用。下記参照） |

### 使い方

| Step | コマンド | 内容 |
|---|---|---|
| 0 | `.\run.ps1 -Step extract` | `mods_jar_dir`の`.jar`を`mods_root`に展開（差分のみ、スキップあり） |
| 1 | `.\run.ps1 -Step xlsx` | `ALL_mods_lang_diff.xlsx` / `ALL_mods_patchouli_diff.xlsx` を生成（自動翻訳込み） |
| - | (任意) | 訳の質を上げたい行だけ、xlsxを開いて`ja_jp(new)` / `ja_text(new)`列を手直しする |
| 2 | `.\run.ps1 -Step check` | 任意。訳語衝突・未翻訳を`merged_conflicts.json`等に出力 |
| 3 | `.\run.ps1 -Step pack` | 翻訳済みxlsxからリソースパック(lang用/Patchouli用)を生成・zip化 |
| - | `.\run.ps1` | 上記を通しで実行 |

採用の優先順位は **`(new)` > `(existing)` > `ja_auto_google`**。手直ししなければ
自動翻訳（またはmod既存訳）がそのまま採用されます。

> PowerShell で「スクリプトの実行が無効」と言われる場合は、`.bat` の方を使ってください。

### 翻訳が反映されない・衝突の原因が分からないとき

```powershell
.\run.ps1 -Step diagnose
```

`_out\diagnose_report.json`に、mods_rootにはあるのにxlsxに1行も無いmodid
(`modids_missing_from_sheet`)と、xlsxの`en_us`列が実際のmodファイルと
一致しない行(`row_mismatches`)を出力します。後者はGoogleスプレッドシート上で
一部の列だけソート/フィルタして行がズレた場合に典型的に発生するので、
該当行は元のxlsxから翻訳を入れ直すか、該当modだけStep1からやり直すのが安全です。

## 単体パック（互換パック等、`.\run_rp.ps1`）

互換パックなど、それ単体で完結したリソースパック/データパックをまとめて
日本語化します。modpack本体とは完全に別のconfig/コマンドで、`mods_root`は不要です。

### セットアップ

`config.resourcepack.json`（`setup.bat` で作られます）の `single_packs_dir` を書き換え、
その直下に翻訳したい単体パックのフォルダまたは`.zip`をまとめて置きます。

### 使い方

```powershell
.\run_rp.ps1 -Step scan     # single_packs_dir をスキャンして single_packs_diff.xlsx を生成(自動翻訳込み)
.\run_rp.ps1 -Step apply    # 翻訳済みxlsxを各パックのコピーに反映してzip化
```

`scan`はパックごとに以下を拾います（1つも無ければ`summary`シートに「翻訳対象なし」と出ます。
構造物データのみのcompatパックなど、そもそも翻訳対象が無いパックも珍しくありません）:

- `description_diff` : `pack.mcmeta`の`description`（プレーンな文字列の場合のみ。
  テキストコンポーネント形式は参考表示のみで自動反映されない）
- `lang_diff` : `assets/*/lang/en_us.json`（あれば）
- `patchouli_diff` : Patchouli本文（あれば）

`apply`は`_out\_single_packs_out\<pack_name>_JA\`（元パックの丸ごとコピー＋翻訳反映）
と`_out\<pack_name>_JA.zip`を生成します。差分だけの別パックではなく
**元パックの日本語版そのもの**なので、resourcepacks/datapacksフォルダの元パックと
入れ替えて使います。

## 自動翻訳について（既定は無効・Googleスプレッドシート運用）

既定(`auto_translate: false`)では、`ja_auto_google`列に
`GOOGLETRANSLATE`数式を書き込むだけです。xlsxをGoogleドライブにアップロードして
スプレッドシートとして開き、計算されるのを待ってから`(new)`列を埋める運用になります。

`auto_translate: true`にすると、xlsx生成時に未翻訳の文字列をまとめてPython側
（[deep-translator](https://pypi.org/project/deep-translator/)経由、Google翻訳のWeb版・APIキー不要）で自動翻訳し、
`ja_auto_google`列に**その場で翻訳済みテキスト**を書き込みます
(Googleスプレッドシートを開かずに済みます)。

- `deep-translator` は `setup.bat` で一緒にインストールされます。未インストール/通信エラー時は
  自動的に `GOOGLETRANSLATE` 数式を書き込む動作にフォールバックします
- 未翻訳の文字列は最大25件・2500文字程度ずつまとめて1リクエストで翻訳します。
  結合・分割がうまくいかなかったバッチだけ、その場で1件ずつに切り替えます
- 同じ文字列は1回しか翻訳しません（プロセス内キャッシュ）
- レート制限を検知した場合は自動で待機・リトライします。それでも失敗が続く場合は、
  その回の実行中は自動翻訳を諦めて`GOOGLETRANSLATE`数式フォールバックに切り替わります
- modが数百規模で未翻訳の文字列が数万件を超えるような大きいmodpackでは、
  無料のGoogle翻訳(Web版)の許容量を超えやすいです。そういう場合は
  `auto_translate: false`(既定)のままGoogleスプレッドシート運用にする方が現実的です。
  まずは小さめのmodpackか単体パックで試すのがおすすめです
- 訳の質はGoogle翻訳相当なので、固有名詞やゲーム内表現は`(new)`列で手直しすることを推奨します

## ご注意

- **生成した日本語化パックの再配布について**: パックには各 mod の文章（とその翻訳）が含まれます。
  配布したい場合は、それぞれの mod のライセンスや作者の方針を確認してください。
  このツールの MIT ライセンスは、生成されたパックの中身には及びません。
- **自動翻訳について**: `auto_translate: true` は Google 翻訳の公式 API ではなく Web 版を利用します。
  大量のリクエストを送るとブロックされることがあります。自己責任でお使いください。
- 元の mod ファイルは書き換えません（展開・コピーしたものだけを扱います）。

## ファイル構成

```
setup.bat / setup.ps1            初回セットアップ（パッケージのインストールとconfigの作成）
run.bat / run.ps1                modpack本体用ランチャー
run_rp.bat / run_rp.ps1          単体パック用ランチャー
config.example.json              modpack本体用設定のひな形（config.jsonとしてコピーして編集）
config.resourcepack.example.json 単体パック用設定のひな形（config.resourcepack.jsonとしてコピーして編集）
requirements.txt                 必要なPythonパッケージ
common.py                        共通処理（JSON I/O・進捗表示・xlsxヘルパー・自動翻訳等）
00_extract_mods.py               Step0(任意): .jar群 → mods_root以下に展開
01_build_xlsx.py                 Step1: modフォルダ群 → lang/Patchouli差分xlsxを生成
02_check_conflicts.py            Step2(任意): 翻訳済みxlsxの衝突・未翻訳チェック
03_build_pack.py                 Step3: 翻訳済みxlsx → lang用/Patchouli用リソースパック生成
04_diagnose.py                   診断(任意): xlsxの行ズレ・欠落modidを検出
05_scan_single_packs.py          単体パック: single_packs_dir をスキャンしてxlsx生成
06_apply_single_packs.py         単体パック: xlsxの翻訳を各パックのコピーに反映
run_all.py                       各ステップを呼び出すランチャー本体（直接は呼ばない）
```

## ライセンス

[MIT License](LICENSE) © 2026 みかんのつぶ

使っているもの: [openpyxl](https://openpyxl.readthedocs.io/)（MIT）、[deep-translator](https://github.com/nidhaloff/deep-translator)（Apache-2.0）
