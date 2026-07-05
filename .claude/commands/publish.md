# publish コマンド

前回 publish 以降に大きな変更があった成果物について、CHANGELOG を更新してから publish する。

## 対象成果物と publish コマンド

| 成果物 | CHANGELOG | バージョン管理 | publish コマンド |
|---|---|---|---|
| VS Code 拡張 | `build/extension/CHANGELOG.md` | `build/extension/package.json` | `npm run publish --prefix build/extension` |
| vba-runner (npm) | `build/runner/CHANGELOG.md` | `build/runner/package.json` | 下記参照 |
| vba-extractor (npm) | `build/extractor/CHANGELOG.md` | `build/extractor/package.json` | 下記参照 |

## 認証情報

**VS Code 拡張:** PAT は `build/extension/.env` の `VSCE_PAT`。期限切れは https://aka.ms/vscodepat で再発行して `.env` を更新。

**npm パッケージ:** トークンは `build/runner/.env` の `NPM_TOKEN`。

## 手順

### Step 1: 前回 publish からの変更を把握する

各成果物の CHANGELOG 最新バージョンを起点に git log を確認する。

```bash
# 直近コミット一覧（タグがない場合は CHANGELOG の日付で判断）
git log --oneline --format="%h %ad %s" --date=short | head -30
```

git log と各 CHANGELOG の最新エントリを照合し、**未記載の変更がある成果物**を特定する。迷ったらユーザーに確認する。

### Step 2: CHANGELOG を更新する

- 形式: `## [新バージョン] - YYYY-MM-DD`
- 変更の種類: `### Added` / `### Fixed` / `### Changed` / `### Removed`
- **英語で記述する**（[[feedback_changelog_english]]）
- バグ修正・機能追加のみ記載。リファクタリングや内部整備は省略可

### Step 3: バージョンを bump する

> ⚠️ **バージョン番号は自分で決めずに必ずユーザーに確認してから変更する。**
> patch (`0.1.9` → `0.1.10`) にするか minor (`0.1.9` → `0.2.0`) にするかを AskUserQuestion で聞く。
> 公開後は取り消せないため、確認なしに進めてはいけない。

セマンティックバージョニング（参考）:
- バグ修正のみ → patch (`0.1.9` → `0.1.10`)
- 機能追加あり → minor (`0.1.9` → `0.2.0`)
- 破壊的変更あり → major

**アルファ版のルール（重要）:**
- アルファ版の間はプレリリース番号 `N` のみ上げる: `0.1.1-alpha.1` → `0.1.1-alpha.2`
- パッチバージョンを上げるのは正式版リリース後または意図的な変更時のみ
- `0.1.1-alpha.1` → `0.1.2-alpha.0` は**誤り**（正式版 `0.1.1` を飛ばすことになる）

```bash
# package.json の version を直接編集する（node -e で一括更新が便利）
node -e "
const fs = require('fs');
const p = 'build/runner/package.json';
const j = JSON.parse(fs.readFileSync(p));
j.version = '新バージョン';
fs.writeFileSync(p, JSON.stringify(j, null, 2) + '\n');
"
```

### Step 4: ビルドして publish

**VS Code 拡張:**
```bash
npm run build:extension
npm run publish --prefix build/extension   # build/extension/.env の VSCE_PAT を使用
```

**vba-runner:**
```bash
npm run build:runner
npm run pub --prefix build/runner   # scripts/publish-npm.sh が .npmrc 経由で認証
```

**vba-extractor:**
```bash
npm run build:extractor
npm run pub --prefix build/extractor   # scripts/publish-npm.sh が .npmrc 経由で認証
```

> `scripts/publish-npm.sh` は `.env` から `NPM_TOKEN` を読み取り、一時的な `.npmrc` に書き出してから `npm publish` を実行し、終了後（エラー時も含む）に `.npmrc` を削除する。

### Step 5: git commit → tag → push

```bash
git add build/*/CHANGELOG.md build/*/package.json
git commit -m "Chore: vX.Y.Z (extension) / vA.B.C (runner) — CHANGELOG & version bump"
git tag vX.Y.Z   # 拡張機能バージョンをタグに使う
git push origin vX.Y.Z
```

## 注意事項

- **迷ったらユーザーに確認する**（バージョン bump の判断、変更の重要度の見極め）
- CHANGELOG は英語で書く
- publish 前に全テスト（`./scripts/bg-test.sh`）が通っていること・コミット済みであることを確認
- 変更のない成果物は publish 不要（extractor に固有の変更がなければスキップ）
