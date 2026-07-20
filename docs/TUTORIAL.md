# TUTORIAL.md — Excel VBA を抽出し、AI と一緒にテスト・機能追加するチュートリアル

> 対象: Excel VBA を AI と安全に改善したい利用者
>
> 前提: [vba-extractor](../build/extractor/README.ja.md) と [vba-runner](../build/runner/README.ja.md)（または本リポジトリの clone）
>
> 次に読む: [FOR_AI.md](../FOR_AI.md)、[MOCK_GUIDE.md](MOCK_GUIDE.md)、目的別ハブは [README.md](README.md)
>
> コマンド表記: 本文はパッケージ利用者向け（`vba-runner` / `vba-extractor`）。clone している場合は [CONTRIBUTING.md](../CONTRIBUTING.md#cli-コマンド対応表) の右列に置き換える

## 対象読者

Excel で作成された VBA マクロ（`.xlsm`）を持っていて、

- Excel を開かずに VS Code でソースを読み書きしたい
- AI（Claude Code など）にリファクタリングや機能追加を手伝ってもらいたい
- 変更が既存の動作を壊していないことをテストで保証したい

という人向けのチュートリアルです。最終的に変更結果を `.xlsm` に書き戻すところまでをひととおり体験します。

> 手元に `.xlsm` がなく `.bas`/`.cls` ファイルしかない場合は、まず
> [vba-extractor README](../build/extractor/README.ja.md#新規-xlsm-をゼロから作る-windows--excel)
> の手順で最初の `.xlsm` を作成してから本チュートリアルを始めてください
> (Windows + ローカルにインストールされたExcelが必要です)。

---

## 全体の流れ

```
.xlsm
  │  ① vba-extractor export
  ▼
.bas / .cls （テキストファイル）
  │  ② VS Code で参照（VBA Runner 拡張機能）
  ▼
AI に FOR_AI.md / MOCK_GUIDE.md を読ませる
  │  ③ 解析 → 計画 → 提案 → 承認
  ▼
④ テストで既存動作を記録（GREEN）
  │
⑤ 機能追加（RED → 実装 → GREEN）
  │
  ▼
.bas / .cls （修正済み）
  │  ⑥ vba-extractor import
  ▼
.xlsm （更新済み）
```

このチュートリアルでは具体例として、本リポジトリに収録されているサンプル `sample/excel/test.xlsm` を使います。手元の `.xlsm` がある場合はファイル名を読み替えてください。

---

## 0. 事前準備

```bash
# vba-extractor（VBA ソースの抽出・書き戻し CLI）
npm install -g vba-extractor

# VBA Runner（VS Code 拡張機能。LSP によるホバー・定義ジャンプ・コードレンズ等）
code --install-extension jca02266.vba-runner
```

リファクタリング・テストには `vba-runner` の実行エンジンが必要です。npm パッケージとして使う場合は [build/runner/README.md](../build/runner/README.md) を参照してください。本リポジトリを clone して作業する場合は追加のパッケージインストールは不要で、CLI は [CONTRIBUTING.md](../CONTRIBUTING.md#cli-コマンド対応表) の右列（`npm run vba-runner --` 等）を使います。

---

## ① vba-extractor で VBA ソースを抽出する

```bash
vba-extractor export sample/excel/test.xlsm src/vba
```

```
Modules  : 5
Encoding : cp932
  → ThisWorkbook.cls (368 chars)
  → Sheet1.cls (372 chars)
  → Module1.bas (6742 chars)
  → Class1.cls (460 chars)
  → UserForm1.cls (498 chars)
Done.
```

標準モジュールは `.bas`、クラスモジュール／フォーム／シートオブジェクトは `.cls` として、文字コードは xlsm 内の `PROJECTCODEPAGE` から自動判定して UTF-8 に変換されます。これで `src/vba/` 配下が Git で管理できる通常のテキストファイルになりました。

詳細は [build/extractor/README.md](../build/extractor/README.md) を参照。

---

## ② VS Code でソースを確認する

`src/vba/` を VS Code で開きます。VBA Runner 拡張機能が `.bas`/`.cls`/`.frm` を認識し、以下が使えます。

| 機能 | 操作 |
|---|---|
| ホバー | シンボル上にマウスを置くとシグネチャを表示 |
| 定義へ移動 | カーソルを置いて `F12` |
| 参照箇所一覧 | カーソルを置いて `Shift+F12` |
| コード補完 | VBA キーワード・組み込み関数・自作プロシージャを `.` 入力時等に補完 |
| コードレンズ | プロシージャ宣言の上に `▶ Run` / `🐛 Debug` / `N references` / `Untested`・`✓ Tested` / `📊 Show in Call Graph` ボタンを表示 |
| コールグラフ | コマンドパレットから **VBA: Show Call Graph** |

抽出した `Module1.bas` を開くと、`ProcessData`（エントリポイント）から `CreateSampleData` → `CopyFilteredData` → `SummarizeByCategory` → `FormatResultSheet` が呼ばれる構造が、コードレンズの参照数やコールグラフで一目で確認できます。

詳細は [build/extension/README.md](../build/extension/README.md) を参照。

---

## ③ AI にリファクタリング・機能追加を依頼する準備

AI（Claude Code 等）にこのソースを触ってもらう前に、**[FOR_AI.md](../FOR_AI.md) を読ませます**。これは人間向けではなく AI 向けに書かれたガイドで、以下のサイクルにしたがって自律的に作業を進めるためのルールが書かれています。

```
Phase 1: 解析 → Phase 2: 計画 → Phase 3: 提案 → [ユーザー承認]
                                                        ↓
                              Phase 4: リファクタリング/機能追加
                              Phase 5: テスト
                              Phase 6: 効果測定
                              Phase 7: レポート記録 → コミット
```

依頼の仕方の例：

> `FOR_AI.md` の内容に従って `src/vba/Module1.bas` を解析し、リファクタリング/機能追加の計画を立ててください。

FOR_AI.md の基本ルールで重要なのは次の 3 点です。

1. **VBA ファイルの全文読み禁止** — `vba-runner analyze` が示す行番号の範囲だけ読む（巨大なマクロでもトークンを節約できる）
2. **計画単位で進める** — 提案された計画に承認を出すまで実装には進まない
3. **テストなしで変更しない** — 変更前に必ず既存動作をテストで GREEN にする

Excel オブジェクト（`ActiveSheet`、`Range`、`Cells`、`Application` など）に依存するコードをテストするにはモックが必要です。AI は必要になったタイミングで **[MOCK_GUIDE.md](MOCK_GUIDE.md) の「Step 1: 対応表」** を読み、対象オブジェクトの注入コードを確認します。

---

## ④ 既存動作をテストで記録する

`Module1.bas` の `SummarizeByCategory` はカテゴリ別の件数・売上合計・平均売上を計算して書き出す関数です。Excel に依存しているため、テストには `MockWorksheet` 相当のモックが必要です。

`SummarizeByCategory` は内部で `LastRow`（`ws.Cells(ws.Rows.Count, 1).End(xlUp).Row` で最終行を取得）を使っていますが、組み込みの `MockWorksheet` の `End()` は固定値しか返さないスタブのため、[MOCK_GUIDE.md §C](MOCK_GUIDE.md#c-未対応オブジェクトの拡張rowscount--columnscount--vba定数) の方針にしたがって、セルの実データを走査する簡易モックを自作します。

```typescript
// src/vba/Module1.test.ts
import { VBARunner, assert } from 'vba-runner';

function createMockSheet(): any {
    const cellData: Record<string, any> = {};
    const cellsFn = (r: number, c: number) => ({
        get value() { return cellData[`${r},${c}`] ?? ''; },
        set value(v: any) { cellData[`${r},${c}`] = v; },
        font: { bold: false, color: 0 },
        interior: { color: 0 },
        end: (_dir: any) => {
            let lastRow = 1;
            for (let row = 1; row <= 1000; row++) {
                if (cellData[`${row},${c}`] !== undefined && cellData[`${row},${c}`] !== '') lastRow = row;
            }
            return { row: lastRow };
        },
    });
    const rowsFn: any = (_r: number) => ({ copy: (_d: any) => {}, font: { bold: false, color: 0 }, interior: { color: 0 } });
    rowsFn.count = 1048576;
    return { cells: cellsFn, rows: rowsFn, columns: { count: 16384, autofit: () => {} }, _data: cellData };
}

const vbaRunner = new VBARunner('./Module1.bas');
vbaRunner.set('rgb', (r: number, g: number, b: number) => r + g * 256 + b * 65536);
vbaRunner.set('xlup', -4162);   // Option Explicit のため Excel 定数を明示的に注入する

const wsSource = createMockSheet();
const wsDest = createMockSheet();
[
    ['ID', '商品名', 'カテゴリ', '売上', '日付'],
    [1, 'りんご', '果物', 1200, '2026/01/05'],
    [2, 'バナナ', '果物', 800, '2026/01/07'],
    [3, '牛乳', '乳製品', 320, '2026/01/07'],
    [4, 'チーズ', '乳製品', 950, '2026/01/10'],
].forEach((row, r) => row.forEach((v, c) => { wsSource._data[`${r + 1},${c + 1}`] = v; }));

vbaRunner.run('SummarizeByCategory', [wsSource, wsDest]);

// 集計結果: ■ カテゴリ別集計(row3) / ヘッダー(row4) / 果物(row5) / 乳製品(row6)
assert.strictEqual(wsDest._data['5,1'], '果物', 'カテゴリ名: 果物');
assert.strictEqual(wsDest._data['5,2'], 2, '件数: 2件');
assert.strictEqual(wsDest._data['5,3'], 2000, '売上合計: 1200+800');
assert.strictEqual(wsDest._data['5,4'], 1000, '平均売上: 2000/2');
```

```bash
npx tsx src/vba/Module1.test.ts
```

```
[PASS] SummarizeByCategory([Object], [Object]) -> null (10ms)
```

これで「変更前の動作」がテストとして記録され、安全にリファクタリング・機能追加を始められる状態になりました（FOR_AI.md Phase 5-b）。`Excel` を一度も開かずに動作確認ができている点がポイントです。

> Excel オブジェクトのモック実装の詳細（`ActiveSheet`/`Application` の注入パターン、`Rows(i).Copy` の実装など）は [MOCK_GUIDE.md](MOCK_GUIDE.md) を参照してください。

---

## ⑤ 機能追加をテスト駆動で進める

ここではカテゴリ別集計に **「最高売上」列** を追加する機能追加を行います。AI に依頼する例：

> `SummarizeByCategory` の集計結果に、カテゴリごとの最高売上（E列）を追加してください。テストファーストで進めてください。

### 5-a. RED: 先に期待値をテストに書く

```typescript
// 新機能: E列(5列目)に「最高売上」を追加する
assert.strictEqual(wsDest._data['4,5'], '最高売上', '新ヘッダー: 最高売上');
assert.strictEqual(wsDest._data['5,5'], 1200, '果物の最高売上: 1200');
assert.strictEqual(wsDest._data['6,5'], 950, '乳製品の最高売上: 950');
```

```bash
npx tsx src/vba/Module1.test.ts
# [FAIL] 新ヘッダー: 最高売上 - Expected 最高売上 but got undefined
```

まだ実装していないので失敗します（RED）。これは期待通りです。

### 5-b. 実装する

`SummarizeByCategory` に `maxSales` という `Dictionary` を追加し、集計ループ内で最大値を更新、出力時に E 列へ書き出します。

```vb
Dim maxSales As Object     ' カテゴリ -> 最高売上
Set maxSales = CreateObject("Scripting.Dictionary")

' 集計ループ内
If totals.exists(cat) Then
    totals(cat) = totals(cat) + sales
    counts(cat) = counts(cat) + 1
    If sales > maxSales(cat) Then maxSales(cat) = sales
Else
    totals.Add cat, sales
    counts.Add cat, 1
    maxSales.Add cat, sales
End If

' 出力ヘッダー
wsDest.Cells(startRow, 5).Value = "最高売上"

' 出力ループ内
wsDest.Cells(startRow, 5).Value = maxSales(key)
```

### 5-c. GREEN: テストを再実行する

```bash
npx tsx src/vba/Module1.test.ts
```

```
[PASS] SummarizeByCategory([Object], [Object]) -> null (12ms)
```

既存テスト・新規テストの両方が GREEN になりました。これで `Module1.bas` への機能追加が、Excel を開かずに検証できたことになります（FOR_AI.md Phase 5〜7：テスト → 効果測定 → レポート記録）。

> 機能の規模が大きい場合は、`vba-runner analyze` での解析や `--diff` による効果測定など FOR_AI.md のフルサイクルを使ってください。

---

## ⑥ vba-extractor import で Excel に書き戻す

修正済みの `.bas`/`.cls` を元の `.xlsm` に書き戻します。**安全のため、まずは別名で出力することを推奨します。**

```bash
vba-extractor import sample/excel/test.xlsm src/vba sample/excel/test_updated.xlsm
```

```
⚠️  Warning: import directly modifies the Excel file.
   The file may become corrupted. It is strongly recommended to back up before proceeding.
   Input  : sample/excel/test.xlsm
   Output : sample/excel/test_updated.xlsm

Run import? [y/N]: y
Source files : 5
VBA modules  : 5
Encoding     : cp932
  ✓ ThisWorkbook (368 bytes as cp932)
  ✓ Sheet1 (372 bytes as cp932)
  ✓ Module1 (7550 bytes as cp932)
  ✓ Class1 (460 bytes as cp932)
  ✓ UserForm1 (498 bytes as cp932)
Saved: sample/excel/test_updated.xlsm
```

実行確認のための `y`/`N` 確認プロンプトが必ず出ます（`[output.xlsm]` を省略すると入力ファイルを直接上書きするため、特に注意してください）。

最後に `test_updated.xlsm` を Excel で開き、`Alt+F11` で VBA エディターを起動して `Module1` に「最高売上」列のコードが反映されていることを目視確認します。問題なければ元のファイル名にリネームするか、再度 `vba-extractor import` を出力先なしで実行して上書きしてください。

詳細は [build/extractor/README.md](../build/extractor/README.md) を参照。

---

## 次に読むべきドキュメント

| やりたいこと | 読むドキュメント |
|---|---|
| AI に渡す指示の全体仕様を理解したい | [FOR_AI.md](../FOR_AI.md) |
| Excel オブジェクトのモックをもっと使いこなしたい | [MOCK_GUIDE.md](MOCK_GUIDE.md) |
| JS テストフレームワーク（Jest 等）と組み合わせたい | [TEST_FRAMEWORK_GUIDE.md](TEST_FRAMEWORK_GUIDE.md) |
| テストが書けないコードをどう分離するか学びたい | [REFACTORING_GUIDE.md](REFACTORING_GUIDE.md) |
| 大きなリファクタリングの実例を見たい | [REFACTORING_EXAMPLE.md](REFACTORING_EXAMPLE.md) |
| VS Code 拡張機能の全機能を確認したい | [build/extension/README.md](../build/extension/README.md) |
| ドキュメント全体の地図を見たい | [docs/README.md](README.md) |
| 読み順・問題別検索が欲しい | [docs/INDEX.md](INDEX.md)（任意） |
