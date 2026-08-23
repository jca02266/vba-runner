---
name: publish
description: Publish the repository artifacts by following the project-owned .claude/commands/publish.md procedure, including changelog/version review, build, authentication, publish, commit, tag, and push.
---

# Project publish workflow

Use `.claude/commands/publish.md` as the single source of truth. Read it in full
before acting and follow its artifact-selection, changelog, version-confirmation,
build, publish, and git release steps. Do not copy or restate that command's
artifact table, version rules, credentials, or commands here.

Before publishing, preserve unrelated user changes and verify the repository is
clean and the required test gate from the command has passed. Ask the user for
the version bump whenever the command requires confirmation; never infer a
publish version. Do not edit `.claude/commands/publish.md` as part of this skill.

## CHANGELOG の記載粒度

各成果物の既存 CHANGELOG を先に読み、過去のリリースと同程度の具体性で
記載する。大きな変更を「various fixes」の一行にまとめず、利用者に影響する
機能追加・互換性修正・診断修正を `Added` / `Fixed` / `Changed` ごとに複数の
短い箇条書きへ分ける。各項目には対象となるVBA挙動または開発者向け機能と、
利用者が得る結果を書く。内部リファクタリングだけの項目やコミット一覧は
記載しないが、同じ修正で複数の利用者向け挙動が改善された場合は挙動ごとに
分ける。

既存の見出し・日付・英語表記・箇条書きの形式を維持し、最新エントリの粒度を
下回らないこと。更新前に `git log` と最新エントリを照合し、今回の変更が
どの成果物に含まれるかを判断する。
