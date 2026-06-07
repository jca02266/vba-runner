# VBA機能実装スキル

TODO_SPEC.mdから優先度の高い未実装機能を1件選び、テスト作成・実装・検証・コミットまでを一連で行う。

## 手順

### Step 1: 対象機能の選定

`TODO_SPEC.md` を読み、「高優先度」セクションの先頭にある `❌` 項目を1件選ぶ。
高優先度がすべて完了していれば「中優先度」から選ぶ。

### Step 2: TODO_SPEC.mdに実装中マークを記録

選んだ項目の `❌` を `🚧` に変更してTODO_SPEC.mdを保存する。
（例: `| ❌ | Select Case Statement |` → `| 🚧 | Select Case Statement |`）

### Step 3: 仕様書で該当機能を確認

`spec/MS-VBAL.txt` から該当セクションを読み、構文ルール・動作仕様・エッジケースを把握する。
仕様書の章番号はTODO_SPEC.mdの「仕様書」列を参照する。

**行番号の引き方**: `spec/MS-VBAL-index.txt` に節番号→行番号の対応表がある。
grep で引いた行番号を使って `sed -n 'LINE,+100p' spec/MS-VBAL.txt` で本文を読む。

```bash
# 例: §5.2.4 を読む
grep "^§5\.2\.4 " spec/MS-VBAL-index.txt   # → 行番号を取得
sed -n '1956,+80p' spec/MS-VBAL.txt         # → その行から本文を読む
```

### Step 4: テストを作成

**通常の機能実装の場合**:
`tests/spec/` 配下に機能名に対応するテストファイルを作成する（例: `tests/spec/select-case.test.ts`）。

**VBA ランタイム挙動の場合**:
以下の両方を作成する：
- TypeScript テスト: `tests/spec/<機能名>.test.ts` （基本的な動作確認用）
- VBA ソーステスト: `tests/vba/<機能名>Test.bas` （実際のVBAコードの挙動確認用）
  - ファイル名例: `ByrefArraysTest.bas`, `DefaultPropertyTest.bas`, `AutoInstantiationTest.bas`
  - テストプロシージャは `Test` で始まる名前で記述（`TestBasicBehavior` など）
  - TODO_SPEC.md の「VBA ランタイム挙動」セクションに両テストファイル名を記載

#### tests/spec/ のテスト記述ルール

**ファイル冒頭のボイラープレート**（必ずこの形式で書く）:

```typescript
import { evalVBASingle, evalVBAModules, assert } from '../../test-libs/test-runner';

function runFunc(code: string, name: string, args: any[] = []): any {
    return evalVBASingle(code).callProcedure(name, args);
}
```

`evalVBASingle` と `evalVBAModules` は `test-libs/test-runner.ts` にエクスポートされている共有関数。ローカルに定義しない。

| 関数 | 用途 |
|---|---|
| `evalVBASingle(code)` | 単一モジュール。Pass 1 + Pass 2 を実行して `Evaluator` を返す |
| `evalVBAModules([{code, name}, ...])` | 複数モジュール。全モジュールを Pass 1 でロード後、Pass 2 を1回実行して `Evaluator` を返す |

**複数モジュールの例**（`Module1.A` のようなモジュール修飾アクセスが必要な場合）:

```typescript
const ev = evalVBAModules([
    { name: 'ModA', code: modACode },
    { name: 'ModB', code: modBCode },
]);
ev.callProcedure('TestProc', []);
```

> **2パスの設計について**: `evaluate()` は Pass 1（手続き・変数・クラスの登録のみ）で、モジュールレベルの `Const` は評価しない。定数は Pass 2（`reEvaluateModuleConstsAll`）で依存グラフ順に確定させる。`evalVBASingle` / `evalVBAModules` は両パスをカプセル化しているため、利用者が意識する必要はない。

> **既存テストの移行**: `tests/spec/` 内の古いテストは `evalVBASingle` をローカル定義（`reEvaluateModuleConstsAll` なし）している。これらは現在の定数がない or `setConstants()` 注入のテストであるため即壊れないが、モジュールレベル `Const` を追加した場合は壊れる。既存テストを修正する際は、ローカル定義の `evalVBASingle` を削除して `test-libs/test-runner` からのインポートに置き換えること。

**テストの書き方**:
- VBAコードはテンプレートリテラルのインライン文字列として書く（`.bas` ファイルは使わない）
- 引数なしでプロシージャを呼ぶ場合も `runFunc(code, 'FuncName')` を使う
- 式や副作用のみ確認する場合は `evalVBASingle(code)` を使う
- アサーションは `assert.strictEqual(actual, expected, 'テストの説明')` を使う
- 各テストグループの末尾に `console.log('[PASS] テスト名')` を入れる
- ファイル末尾に `console.log('\n✅ <機能名>: 全テスト通過')` を入れる

**テストで網羅する観点**:
- 基本動作（正常系）
- エッジケース（仕様書に明記されているもの）
- エラー系（不正な入力）
- **VBA ベテランが書きそうな代表コード**（重要）: 単に API 呼び出しを確認するだけでなく、実務でハマりやすい暗黙のランタイム挙動や、ライブラリの組み合わせ利用パターンを必ず含める
  - 例: `Dim As New X` の Auto-Instantiation、Default Property 経由の暗黙呼び出し、`Variant` の型推移、`On Error` ハンドラー内での再帰エラー など
  - 「仕様書のリストにある関数を呼んでみた」だけでは VBA の "癖" を見落とすため、ユーザーの実際の書き方を想定する

#### tests/vba/ の VBA ソーステスト構成ルール

**ファイル命名規則**:

| 種類 | 命名 | 例 |
|---|---|---|
| テストプロシージャを含む `.bas` | `<名前>Test.bas` | `ArgCountTest.bas` |
| テスト専用サポートクラス `.cls` | 短い名前（`Test` サフィックスなし） | `RefA.cls`, `Helper.cls` |
| 共通インフラクラス `.cls` | 用途がわかる名前 | `AssertHelper.cls` |

**モジュール名の長さ制限（31文字）への対応**:
- VBA ではモジュール名（= ファイル名から拡張子を除いたもの）が **31文字以内** でなければならない
- テストに複数のファイル（`.bas` + 複数の `.cls`）が必要な場合は、**サブディレクトリを作成**してファイル名を短くする

```
tests/vba/
├── Circular/               ← 複数ファイルが必要なテストはサブディレクトリに
│   ├── TerminateTest.bas  ← テストプロシージャを含む（"Test"サフィックス）
│   ├── RefA.cls            ← テスト専用クラス（短い名前）
│   ├── RefB.cls
│   └── Helper.cls
├── ArgCountTest.bas        ← 単一ファイルで完結するテスト
└── AssertHelper.cls        ← 全テスト共通のインフラクラス
```

**テストプロシージャの書き方**:
- `Sub Test_XXX(assert)` 形式で定義。ランナーが `test` で始まるプロシージャを自動検出する
- 検証は `AssertHelper` オブジェクト（引数 `assert`）のメソッドを使う:
  - `assert.Assert actual, expected, "メッセージ"` — 値比較、失敗時に Debug.Print して Err.Raise
  - `assert.IsTrue value, "メッセージ"` — 真偽値テスト
  - `assert.IsFalse value, "メッセージ"` — 偽であることのテスト
- `SetUp` / `TearDown` Sub は引数なし。ランナーが各テストの前後に自動呼び出し

```vb
Option Explicit

Sub Test_BasicBehavior(assert)
    Dim result As Long
    result = SomeFunc(1, 2)
    assert.Assert result, 3, "1+2=3"
End Sub

Sub SetUp()
    ' 各テスト前の初期化
End Sub

Sub TearDown()
    ' 各テスト後のクリーンアップ
End Sub
```

> **注意**: `test-libs/test-runner.ts` の `VBARunner` クラスは `sample/tests/ts/` 配下の `.bas` ファイルを読み込むためのもの。`tests/spec/` では使わない。

### Step 5: 実装

実装対象ファイル（通常は以下の1つ以上）を編集する:
- `src/engine/lexer.ts` — 新しいキーワード・トークンが必要な場合
- `src/engine/parser.ts` — 新しいASTノード・構文規則が必要な場合
- `src/engine/evaluator.ts` — 実行時の動作が必要な場合

実装は仕様書の動作定義に忠実に行う。
仕様と異なる実装や制限事項がある場合は、必ず TODO_SPEC.md の該当項目に制限事項としてメモを残すこと。

### Step 6: テスト実行

以下のコマンドでテストを実行し、全件パスすることを確認する:

```bash
npx tsx tests/spec/<テストファイル名>.ts
```

既存テストのリグレッションも確認する:

```bash
npx tsx sample/tests/ts/TaskScheduler_Core.test.ts
```

### Step 7: 仕様書との照合

テストがパスしたら `spec/MS-VBAL.txt` の該当セクションを再読し、実装が仕様を満たしているか確認する。
問題があれば実装を修正してStep 6に戻る。

### Step 8: TODO_SPEC.mdを更新してコミット

`TODO_SPEC.md` の対象項目を `🚧` → `✅` に更新する。
部分実装や仕様上の制限がある場合は `⚠️`（備考を括弧内に記載）にする。
仕様と異なる挙動が残る場合は、該当行に制限事項を明記すること。

**VBA ランタイム挙動の場合は、テスト列に両テストファイル名を記載する：**
```
| ✅ | **機能説明** | テスト: `feature-name.test.ts`, `feature_name_test.bas` |
```

以下の形式でコミットする:

```
Feat: Implement <機能名> (<仕様書章番号>)

- <実装内容の箇条書き>
- テスト: <テストファイル名>（VBA ランタイム挙動の場合は TypeScript と VBA の両方を記載）
```
