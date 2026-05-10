# VBA機能実装スキル

TODO.mdから優先度の高い未実装機能を1件選び、テスト作成・実装・検証・コミットまでを一連で行う。

## 手順

### Step 1: 対象機能の選定

`TODO.md` を読み、「高優先度」セクションの先頭にある `❌` 項目を1件選ぶ。
高優先度がすべて完了していれば「中優先度」から選ぶ。

### Step 2: TODO.mdに実装中マークを記録

選んだ項目の `❌` を `🚧` に変更してTODO.mdを保存する。
（例: `| ❌ | Select Case Statement |` → `| 🚧 | Select Case Statement |`）

### Step 3: 仕様書で該当機能を確認

`spec/MS-VBAL.txt` から該当セクションを読み、構文ルール・動作仕様・エッジケースを把握する。
仕様書の章番号はTODO.mdの「仕様書」列を参照する。

### Step 4: テストを作成

`tests/spec/` 配下に機能名に対応するテストファイルを作成する（例: `tests/spec/select-case.test.ts`）。

#### tests/spec/ のテスト記述ルール

**ファイル冒頭のボイラープレート**（必ずこの形式で書く）:

```typescript
import { Lexer } from '../../src/compiler/lexer';
import { Parser } from '../../src/compiler/parser';
import { Evaluator } from '../../src/compiler/evaluator';
import { assert } from '../ts/test-runner';

function evalVBA(code: string): any {
    const tokens = new Lexer(code).tokenize();
    const ast = new Parser(tokens).parse();
    const ev = new Evaluator(console.log);
    ev.evaluate(ast);
    return ev;
}

function runFunc(code: string, name: string, args: any[] = []): any {
    return evalVBA(code).callProcedure(name, args);
}
```

**テストの書き方**:
- VBAコードはテンプレートリテラルのインライン文字列として書く（`.vba` ファイルは使わない）
- 引数なしでプロシージャを呼ぶ場合も `runFunc(code, 'FuncName')` を使う
- 式や副作用のみ確認する場合は `evalVBA(code)` を使う
- アサーションは `assert.strictEqual(actual, expected, 'テストの説明')` を使う
- 各テストグループの末尾に `console.log('[PASS] テスト名')` を入れる
- ファイル末尾に `console.log('\n✅ <機能名>: 全テスト通過')` を入れる

**テストで網羅する観点**:
- 基本動作（正常系）
- エッジケース（仕様書に明記されているもの）
- エラー系（不正な入力）

> **注意**: `tests/ts/test-runner.ts` の `VBATest` クラスは `sample/tests/ts/` 配下の `.vba` ファイルを読み込むためのもの。`tests/spec/` では使わない。

### Step 5: 実装

実装対象ファイル（通常は以下の1つ以上）を編集する:
- `src/compiler/lexer.ts` — 新しいキーワード・トークンが必要な場合
- `src/compiler/parser.ts` — 新しいASTノード・構文規則が必要な場合
- `src/compiler/evaluator.ts` — 実行時の動作が必要な場合

実装は仕様書の動作定義に忠実に行う。

### Step 6: テスト実行

以下のコマンドでテストを実行し、全件パスすることを確認する:

```bash
./node_modules/.bin/esbuild tests/spec/<テストファイル名>.ts --bundle --outfile=tests/spec/<テストファイル名>.cjs --platform=node && node tests/spec/<テストファイル名>.cjs
```

既存テストのリグレッションも確認する:

```bash
./node_modules/.bin/esbuild sample/tests/ts/TaskScheduler_Core.test.ts --bundle --outfile=sample/tests/ts/TaskScheduler_Core.test.cjs --platform=node && node sample/tests/ts/TaskScheduler_Core.test.cjs
```

> **注意**: `esbuild` は PATH に入っていないため `npx` ではなく `./node_modules/.bin/esbuild` で実行すること。

### Step 7: 仕様書との照合

テストがパスしたら `spec/MS-VBAL.txt` の該当セクションを再読し、実装が仕様を満たしているか確認する。
問題があれば実装を修正してStep 6に戻る。

### Step 8: TODO.mdを更新してコミット

`TODO.md` の対象項目を `🚧` → `✅` に更新する。
部分実装の場合は `⚠️`（備考を括弧内に記載）にする。

以下の形式でコミットする:

```
Feat: Implement <機能名> (<仕様書章番号>)

- <実装内容の箇条書き>
- テスト: <テストファイル名>
```
