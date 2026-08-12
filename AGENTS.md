# Repository guidance

## Commit messages

Use the conventional Git message layout: a concise, type-prefixed summary
line, followed by one blank line, followed by a detailed body when needed.
The summary is the overview of the change; the body explains the reason and
the material implementation details. Select a type that makes the change
category clear, such as `Fix` for bug fixes, `Add` for new functionality,
`Improve` for behavioral improvements, `Docs`, `Test`, `Chore`, or `Verify`.
Make the summary name the affected behavior and outcome; avoid vague messages
that require reading the diff to understand it.

Keep the summary to 50 characters or fewer when practical (never more than 72).
When a body is needed, enter it as real separate lines and separate it from
the summary with exactly one blank line. Wrap body paragraphs at about 72
characters. Do not represent paragraph breaks with escape notation. Omit the
body when the summary alone is sufficient.

Create the message with separate `git commit -m` arguments, for example:

```bash
git commit -m 'Fix: preserve ByRef values' \
  -m 'Keep the reference alive across the call boundary.' \
  -m 'Add a regression test for nested procedure calls.'
```

In Bash or Zsh, ANSI-C quoting is also valid when the complete body is more
convenient to write as one argument:

```bash
git commit -m $'Fix: preserve ByRef values\n\nKeep the reference alive across the call boundary.\nAdd a regression test for nested procedure calls.'
```

Do not pass the two-character sequence `\n` as message text and never use an
ordinary quoted string such as `git commit -m "summary\n\nbody"`. The
sequence is permitted only inside Bash/Zsh `$'...'` quoting, where it becomes
an actual line break. Each `-m`
argument is a paragraph; Git inserts the real blank line between paragraphs.

## テストの配置

`tests/spec/` はVBA言語仕様とランナー実行意味論のテスト専用とする。
評価状態CLI、レポート生成、その他の開発環境・運用ツールのテストは
`tests/tooling/`（または対象ツールに対応する専用ディレクトリ）へ置き、
`tests/spec/`へ追加しない。

## VBAソース文字列のテスト記法

TypeScriptテストに複数行のVBAソースを埋め込む場合は、通常のテンプレート
リテラルではなく `String.raw\`...\`` で囲む。これにより、Windowsパスなどに
含まれるバックスラッシュやVBAソース内の改行をJavaScriptのエスケープ処理で
変形させず、そのままテスト対象へ渡せる。新規テストと既存テストの修正では
この記法を必須とする。

## バグ原因の報告

バグの原因を説明するときは、原因や再現経路だけで終わらせず、現在の
対処状況も必ず併記する。少なくとも、修正済みか未修正か、修正済みなら
対象コード・回帰テスト・検証結果、未修正なら残る影響範囲と次の対処を
明記する。リファクタリングで一部経路だけ解消した場合は、解消範囲と
未検証・未解消の経路を区別して報告する。

## docs の OKF 準拠

`docs/` 配下の知識文書は、Open Knowledge Format（OKF）v0.2 の構造に
段階的に準拠させる。文書を追加・変更するときは、次の規約を守る。

- `docs/index.md` と `docs/internals/index.md` は段階的開示用の索引、
  `docs/log.md` は更新履歴として扱う。これらには concept 用の
  フロントマターを付けない。
- それ以外の `docs/**/*.md` は、先頭に YAML フロントマターを置き、
  少なくとも `type`、`title`、`description`、`tags`、`status` を定義する。
- 文書間の関係は通常の相対 Markdown リンクで表し、新しい文書を追加
  したときは適切な `index.md` に登録する。
- 内容の出典がある場合は `sources`、生成・確認日時が追跡できる場合は
  `generated`、`verified`、`stale_after` を記録する。
- `AGENTS.md`、`CLAUDE.md`、`FOR_AI.md` などの実行規約は OKF 文書へ
  置き換えず、命令の正本として従来どおり維持する。

文書変更後は必ず次を実行する。

```bash
npm run check:docs
```

この検査はフロントマター、相対 `sources.resource`、ローカル Markdown
リンクを確認する。検査スクリプトは
[`scripts/validate-docs-okf.mjs`](./scripts/validate-docs-okf.mjs) である。
