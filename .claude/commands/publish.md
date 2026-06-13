# publish コマンド

前回 publish 以降に大きな変更があった成果物について、CHANGELOG を更新してから publish する。

## 対象成果物と publish コマンド

| 成果物 | CHANGELOG | バージョン管理 | publish コマンド |
|---|---|---|---|
| VS Code 拡張 | `build/extension/CHANGELOG.md` | `build/extension/package.json` | `npm run publish --prefix build/extension` |
| vba-runner (npm) | `build/runner/CHANGELOG.md` | `build/runner/package.json` | `npm publish --prefix build/runner` ※未整備 |
| vba-extractor (npm) | `build/extractor/CHANGELOG.md` | `build/extractor/package.json` | `npm publish --prefix build/extractor` ※未整備 |

> VS Code 拡張の PAT は `build/extension/.env` の `VSCE_PAT` を使う（[[reference_vsce_pat]]）。

## 手順

### Step 1: 前回 publish からの変更を把握する

各成果物の **前回バージョンタグ** から `HEAD` までの git log を確認する。

```bash
# 直近の publish タグ一覧
git tag --sort=-version:refname | head -20

# 拡張機能の前回 publish バージョンを CHANGELOG から取得して比較
# 例: 前回が 0.1.8 なら
git log --oneline v0.1.8..HEAD   # タグがない場合は CHANGELOG の日付で判断
```

git log と各 CHANGELOG の最新エントリを照合し、**未記載の変更がある成果物**を特定する。迷ったらユーザーに確認する。

### Step 2: CHANGELOG を更新する

- 形式: `## [新バージョン] - YYYY-MM-DD`
- 変更の種類: `### Added` / `### Fixed` / `### Changed` / `### Removed`
- **英語で記述する**（[[feedback_changelog_english]]）
- バグ修正・機能追加のみ記載。リファクタリングや内部整備は省略可

### Step 3: バージョンを bump する

```bash
# 拡張機能
# build/extension/package.json の "version" を手動または npm version で更新

# vba-runner / vba-extractor
# build/runner/package.json, build/extractor/package.json の "version" を更新
```

セマンティックバージョニング:
- バグ修正のみ → patch (`0.1.8` → `0.1.9`)
- 機能追加あり → minor (`0.1.8` → `0.2.0`)
- 破壊的変更あり → major

### Step 4: ビルドして publish

**VS Code 拡張:**
```bash
npm run build:extension           # esbuild バンドル
npm run package:extension         # .vsix 生成（確認用）
npm run publish --prefix build/extension   # vsce publish
```

**vba-runner:**
```bash
npm run build:runner
# publish コマンドは要確認
```

**vba-extractor:**
```bash
npm run build:extractor
# publish コマンドは要確認
```

### Step 5: git tag を打つ

```bash
git tag v<拡張バージョン>   # 例: git tag v0.1.9
git push origin v<拡張バージョン>
```

### Step 6: ルート CHANGELOG に記録する

ルートの CHANGELOG がある場合は同様に更新する。

## 注意事項

- **迷ったらユーザーに確認する**（バージョン bump の判断、変更の重要度の見極め）
- CHANGELOG は英語で書く
- publish 前に `npm run typecheck` と `./scripts/bg-test.sh` が通っていることを確認（コミット済みであること）
- vba-runner / vba-extractor の npm publish フローは未整備の可能性があるため、初回は手順を確認してから実行する
