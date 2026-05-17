# リファクタリング・テスト手法カタログ

各ドキュメントで繰り返し登場する手法を名前付きでまとめた参照用カタログです。詳細・実装例は各リンク先を参照してください。

---

## リファクタリング手法

<a id="r-01"></a>
### R-01: Domain Logic / I/O 分離

Excel オブジェクト依存（`Range`, `Cells`, `Sheets` など）を含む処理からビジネスロジックを `Function` に切り出し、I/O を `Sub` に集約する手法。テスト可能性の基礎となる最重要パターン。

- [TESTING_STRATEGY.md — 原則1](TESTING_STRATEGY.md#原則1-domain-logic-と-excel-io-の徹底的分離)
- [REFACTORING_GUIDE.md — 原則1](REFACTORING_GUIDE.md#原則-1-職責の分離separation-of-concerns)
- [MOCK_GUIDE.md — Step 0](MOCK_GUIDE.md#step-0-モックが必要かどうかを確認する)

---

<a id="r-02"></a>
### R-02: ビジネスロジック抽出

巨大な `Sub` から計算・判定ロジックを独立した `Function` として切り出す手法。関数名がビジネスルールの名前になり、可読性・保守性が向上する。

- [REFACTORING_GUIDE.md — パターン1〜3](REFACTORING_GUIDE.md#パターン-1-単純な集計ロジック)
- [REFACTORING_EXAMPLE.md](REFACTORING_EXAMPLE.md)

---

<a id="r-03"></a>
### R-03: UDT パラメータ集約

ドメインとして一体のパラメータ群を `Type` 宣言にまとめて関数に渡す手法。関数シグネチャが簡潔になり、関連しない変数が紛れ込みにくくなる。関係のないパラメータを無理にまとめることは避ける。

- [TESTING_STRATEGY.md — §3](TESTING_STRATEGY.md#3-関連パラメータは-type-にまとめる)
- [REFACTORING_GUIDE.md — パターン3](REFACTORING_GUIDE.md#パターン-3-複雑なビジネスロジックudt-を使用)
- [REFACTORING_EXAMPLE.md — 改善1](REFACTORING_EXAMPLE.md#改善-1-関心の分離と構造化udt-導入)

---

<a id="r-04"></a>
### R-04: クラスモデル化

状態保持・バリデーション・複数の操作が必要になった場合に `Class` を導入する手法。VBA のクラスはコンストラクタへの引数渡し・継承・オーバーロードがなく制限が多いため、`Type` + モジュール関数で足りる間はクラス化を急がない。

- [TESTING_STRATEGY.md — §4](TESTING_STRATEGY.md#4-振る舞いを持たせたい場合はクラスを検討する)
- [REFACTORING_GUIDE.md — パターン4](REFACTORING_GUIDE.md#パターン-4-状態を持つ処理オブジェクト指向版)

---

<a id="r-05"></a>
### R-05: データ Double Buffering

Excel シートのデータをメモリ上の配列に一括読み込みし、配列だけで処理して最後に一括書き戻す手法。シートへのアクセス回数を最小化してパフォーマンスを向上し、ロジック部分を Excel 依存から切り離してテスト可能にする。

- [REFACTORING_EXAMPLE.md — フェーズ2 データ読み込み](REFACTORING_EXAMPLE.md#フェーズ-2-データ読み込み行-101-122)

---

<a id="r-06"></a>
### R-06: 定数グループ化

バラバラに定義された定数を意味のまとまりで `Type` にまとめ、初期化関数で一括セットアップする手法。変更箇所が一ヶ所に集中し、関連する定数の一貫性が保ちやすくなる。

- [REFACTORING_EXAMPLE.md — 改善1](REFACTORING_EXAMPLE.md#改善-1-関心の分離と構造化udt-導入)

---

<a id="r-07"></a>
### R-07: ホットスポット優先戦略

すべてのコードを一度にリファクタリングするのではなく、変更頻度が高い・複雑度が高い・テストが切実に必要な箇所から着手する戦略。投資対効果を最大化する。

- [REFACTORING_GUIDE.md — フェーズ1](REFACTORING_GUIDE.md#フェーズ-1-ホットスポットを特定)

---

<a id="r-08"></a>
### R-08: 段階的リファクタリング

1つの `Function` を抽出 → テスト追加 → 元の `Sub` から呼び出しに差し替え、という小さなサイクルを繰り返す手法。一度に大きく変えずリグレッションリスクを低く保つ。

- [REFACTORING_GUIDE.md — フェーズ2〜3](REFACTORING_GUIDE.md#フェーズ-2-小さく始める)

---

<a id="r-09"></a>
### R-09: セル位置の設定シート管理

列番号・行番号の即値をコードに埋め込まず、Excel の設定シート（テーブル）で管理して `GetSetting()` などの関数で動的に参照する手法。シートレイアウト変更時のコード修正を不要にする。

- [REFACTORING_EXAMPLE.md — セル位置の即値を避ける](REFACTORING_EXAMPLE.md#さらなるリファクタリングセル位置の即値を避ける)

---

## テスト手法

<a id="t-01"></a>
### T-01: 純粋関数テスト

副作用なし・Excel 依存なしの `Function` を直接呼び出して入出力を検証するテストパターン。最もシンプルで安定しており、テスト設計の出発点となる。

- [TESTING_STRATEGY.md — パターン1](TESTING_STRATEGY.md#パターン1-純粋関数テスト最適)
- [TEST_FRAMEWORK_GUIDE.md — パターン1](TEST_FRAMEWORK_GUIDE.md#パターン1-基本的な単体テスト)

---

<a id="t-02"></a>
### T-02: パラメータ化テスト（forEach）

同じ検証ロジックに対して複数の入力・期待値ペアをまとめて `forEach` で実行するパターン。テストケースの追加が容易で、網羅性を高めやすい。

```typescript
const cases: [InputType, number, string][] = [
    [{ ... }, 120, '通常'],
    [{ ... },  10, 'MinStock に丸め'],
];
cases.forEach(([params, expected, label]) => {
    assert.strictEqual(calcFn(params), expected, label);
});
```

- [TESTING_STRATEGY.md — パターン2](TESTING_STRATEGY.md#パターン2-パラメータ化テスト複数ケース)
- [TEST_FRAMEWORK_GUIDE.md — パターン2](TEST_FRAMEWORK_GUIDE.md#パターン2-パラメータ化テスト)

---

<a id="t-03"></a>
### T-03: Partial + デフォルト値テストヘルパー

`Partial<T>` とデフォルト値のスプレッドを組み合わせたヘルパー関数を用意し、テストごとに変化するフィールドだけを書く手法。テストケースが多い場合にコードを大幅に短縮できる（TypeScript 上級者向け）。

```typescript
type Params = Parameters<typeof calcFn>[0];
const base: Params = { A: 0, B: 0, C: 10, D: 200 };
const calcWith = (overrides: Partial<Params>) => calcFn({ ...base, ...overrides });

assert.strictEqual(calcWith({ A: 100, B: 30 }), 120);
```

- [TESTING_STRATEGY.md — §3](TESTING_STRATEGY.md#3-関連パラメータは-type-にまとめる)

---

<a id="t-04"></a>
### T-04: 状態変更テスト

モジュール変数やクラスの状態を初期化してから複数回呼び出し、状態遷移を順番に検証するパターン。`beforeEach` で状態をリセットし、テスト間の干渉を防ぐ。

- [TESTING_STRATEGY.md — パターン3](TESTING_STRATEGY.md#パターン3-状態変更を伴うロジックcontroller)
- [TEST_FRAMEWORK_GUIDE.md — beforeEach/afterEach](TEST_FRAMEWORK_GUIDE.md#1-beforeeach--aftereach-でセットアップクリーンアップ)

---

<a id="t-05"></a>
### T-05: エラーハンドリングテスト

意図的に異常系の入力を与え、期待通りのエラーが発生することを検証するパターン。正常系と同じ粒度でエラー系のケースを網羅する。

- [TESTING_STRATEGY.md — パターン4](TESTING_STRATEGY.md#パターン4-エラーハンドリング)
- [TEST_FRAMEWORK_GUIDE.md — パターン4](TEST_FRAMEWORK_GUIDE.md#パターン4-エラーハンドリングのテスト)

---

<a id="t-06"></a>
### T-06: 仮想ファイルシステムテスト（VFS）

`MemoryFileSystem` を使い、実際のファイルを作成せずにファイル I/O を含む VBA コードをテストする手法。サンドボックスルートを設定してパステストも可能。

- [TEST_FRAMEWORK_GUIDE.md — パターン6](TEST_FRAMEWORK_GUIDE.md#パターン6-ファイルシステム操作vfs-使用)

---

<a id="t-07"></a>
### T-07: 日時モック（Time Mocking）

`vbaRunner.mockDate()` で `Now` / `Date` / `Time` / `Timer` が返す値を固定し、日時依存のロジックを決定論的にテストする手法。

```typescript
vbaRunner.mockDate('2024-12-31T09:00:00');
// ...テスト実行...
vbaRunner.mockDate(null); // 解除
```

- [TEST_FRAMEWORK_GUIDE.md — パターン7](TEST_FRAMEWORK_GUIDE.md#パターン7-日時依存テストtime-mocking)

---

<a id="t-08"></a>
### T-08: スパイ（Spy）

`vbaRunner.spy()` で VBA 関数をラップし、呼び出し回数・引数・戻り値を記録して検証する手法。`MsgBox` などの副作用のある関数の検証やモック化に使う。

```typescript
const spy = vbaRunner.spy('MsgBox', () => 6); // vbYes を返すモック
vbaRunner.run('MyProc', []);
assert.strictEqual(spy.callCount, 1);
assert.ok(spy.calledWith('確認しますか？'));
```

- [TEST_FRAMEWORK_GUIDE.md — spy()](TEST_FRAMEWORK_GUIDE.md#spyname-returnfn--spyrecord)

---

<a id="t-09"></a>
### T-09: 外部オブジェクトスタブ

`vbaRunner.registerExternalObject()` で `CreateObject(progId)` が返すオブジェクトをテスト用スタブに差し替える手法。`VBScript.RegExp` などのランタイム依存を排除できる。

- [TEST_FRAMEWORK_GUIDE.md — registerExternalObject()](TEST_FRAMEWORK_GUIDE.md#registerexternalobjectprogid-factory--void)
- [MOCK_GUIDE.md](MOCK_GUIDE.md)

---

<a id="t-10"></a>
### T-10: MockWorksheet 統合テスト

`MockApplication` / `MockWorksheet` を使い、Excel I/O を含む複数関数の相互作用を検証する統合テストパターン。単体テストでは確認できない関数間の連携を検証できる。

- [MOCK_GUIDE.md — Part 1](MOCK_GUIDE.md#1-クイックスタート)
- [INTEGRATION_TEST_EXAMPLE.md](INTEGRATION_TEST_EXAMPLE.md#なぜ統合テストが必要か)

---

<a id="t-11"></a>
### T-11: 手動モック実装（4パターン）

既存の `MockWorksheet` を使わず、テスト対象に合わせた最小限のモックオブジェクトを自作する手法。軽量オブジェクト・クラス・ビルダー・JSON ベースの 4 パターンがある。

- [MOCK_GUIDE.md — §7](MOCK_GUIDE.md#7-自作モックのパターン4-つ)

---

<a id="t-12"></a>
### T-12: VBA IDE 手動テスト

Excel I/O や UI 操作を伴う `Sub` は VBA IDE 内で実行し、シートの状態を目視確認する手法。自動テストの対象外となる処理の検証方法。

- [TESTING_STRATEGY.md — 原則3](TESTING_STRATEGY.md#原則3-excel-io-テストは-vba-ide-で行う)

---

## テスト戦略

<a id="s-01"></a>
### S-01: テストピラミッド戦略

Unit Tests（70〜80%）→ Integration Tests（15〜25%）→ E2E / VBA IDE 手動テスト（5〜10%）の比率でテストを構成する戦略。下層ほど高速・安定で、上層ほど現実に近いが遅く壊れやすい。

- [TESTING_STRATEGY.md](TESTING_STRATEGY.md)
- [MOCK_GUIDE.md — §12](MOCK_GUIDE.md#12-テスト戦略段階的なテスト化)
