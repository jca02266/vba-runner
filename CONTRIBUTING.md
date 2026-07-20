# Contributing to VBA Runner

> 対象: このリポジトリ（エンジン・拡張・配布物）を開発する人
>
> 前提: Node.js / npm が使えること
>
> 次に読む: [CLAUDE.md](CLAUDE.md)、[REFERENCE.md](REFERENCE.md)、利用者向け地図は [docs/README.md](docs/README.md)

VBA Runner本体、npmパッケージ、VS Code拡張機能、Web UIの開発に参加するためのガイドです。

`vba-runner`や`vba-extractor`を利用するだけの場合は、この文書ではなく各パッケージのREADMEから始めてください。

- [vba-runner](build/runner/README.ja.md)
- [vba-extractor](build/extractor/README.ja.md)
- [VS Code拡張機能](build/extension/README.ja.md)
- [VBA Web Runner](build/playground/README.md)
- [ドキュメントガイド](docs/README.md)（目的別の共通ハブ）

---

## セットアップ

```bash
git clone https://github.com/jca02266/vba-runner.git
cd vba-runner
npm install
```

## よく使うコマンド

| 目的 | コマンド |
|---|---|
| 型チェック | `npm run typecheck` |
| Lint | `npm run lint` |
| 全テスト | `npm test` |
| 全成果物をビルド | `npm run build` |
| npmパッケージ版をビルド | `npm run build:runner` |
| VS Code拡張機能をビルド | `npm run build:extension` |
| Web UIをビルド | `npm run build:playground` |
| vba-extractorをビルド | `npm run build:extractor` |

個別の仕様テストは、対象ファイルを直接実行できます。

```bash
npx tsx tests/spec/<test-file>.test.ts
```

## CLI コマンド対応表

ドキュメントと実装で使う CLI の正本は **`vba-runner <subcommand>`** / **`vba-extractor <command>`** です。
読者によって前置だけが変わります。

| やりたいこと | パッケージ利用者（`npm install`） | リポジトリ開発者（この repo を clone） |
|---|---|---|
| 解析 | `vba-runner analyze <path> ...` | `npm run vba-runner -- analyze <path> ...` |
| 整形 | `vba-runner format <path> ...` | `npm run vba-runner -- format <path> ...` |
| 構文チェック | `vba-runner parse-check <path>` | `npm run vba-runner -- parse-check <path>` |
| 実行 | `vba-runner run <path> ...` | `npm run vba-runner -- run <path> ...` |
| 抽出 | `vba-extractor export ...` | `npm run vba-extractor -- export ...` |
| 書き戻し | `vba-extractor import ...` | `npm run vba-extractor -- import ...` |
| テスト用 TS の実行 | `npx tsx path/to/test.ts` | 同じ |

- 利用者向け文書（`docs/`、`FOR_AI.md`、各パッケージ README）の本文は **パッケージ利用者形** で書く
- このリポジトリを clone している場合は、上表の右列に置き換える
- 旧単独コマンド（`vba-analyzer` / `vba-formatter` / `vba-parse-check` / `vba-run`）は npm パッケージに残る互換エイリアス。新規の説明では使わない
- `npx tsx test-libs/vba-analyzer.ts` などの直接実行は非推奨（開発用ドキュメントでも上表の右列を使う）

## ローカル CLI の実行

リポジトリを clone した開発者は、ビルドせずにソースコード版の CLI を実行できます。

```bash
npm run vba-runner -- <subcommand> [options]
npm run vba-extractor -- <command> [options]
```

例:

```bash
# VBAコードを解析する
npm run vba-runner -- analyze path/to/vba --outline

# VBAコードを整形できるか確認する
npm run vba-runner -- format path/to/Module1.bas --check

# ExcelファイルからVBAソースを抽出する
npm run vba-extractor -- export input.xlsm src/vba
```

これらのコマンドでは、相対ファイルパスはリポジトリルートではなく、npm を実行したディレクトリを基準に解決されます。

npm パッケージ利用者向けの CLI 手順は、各パッケージの README を参照してください。

## プロジェクト構成

| 場所 | 内容 |
|---|---|
| `src/engine/` | VBA Lexer / Parser / Evaluator |
| `test-libs/` | テストランナーと開発用CLIエントリーポイント |
| `tests/` | エンジン・CLI・実VBA比較のnpmテスト |
| `sample/` | サンプルVBAのソース、 ワークスペース |
| `tools/extractor/` | vba-extractorの実装 |
| `build/runner/` | npmパッケージvba-runnerの配布設定 |
| `build/extractor/` | npmパッケージvba-extractorの配布設定 |
| `build/extension/` | VS Code拡張機能の配布設定 |
| `build/playground/` | Web UIの配布設定 |
| `docs/` | 利用者向けのチュートリアル・実践ガイド・内部資料 |

## 目的別の参照ドキュメント

| 参照したいもの | 対象 | まず読む文書 |
|---|---|---|
| 利用者向けの目的別地図 | 利用者 | [docs/README.md](docs/README.md) |
| VBA Runner の実装詳細 | 開発者 | [REFERENCE.md](REFERENCE.md) → [docs/internals/](docs/internals/) |
| VS Code 拡張機能の実装詳細 | 開発者 | [LSP.md](LSP.md) |
| 未着手の開発項目、デバッグ手法 | 開発者 | [TODO.md](TODO.md)、[TODO_SPEC.md](TODO_SPEC.md)、[BUG_HUNTING.md](BUG_HUNTING.md)、[EVAL_LOG.md](EVAL_LOG.md) |
| リポジトリ全体の文書台帳 | メンテナ | [INDEX.md](INDEX.md) |

## AIエージェント向けの補足

利用者のVBAを解析・テスト・リファクタリングする場合は、まず FOR_AI.md を参照してください。
VBA Runner本体を開発するAIエージェント向けの詳細ルールは CLAUDE.md にあります。
