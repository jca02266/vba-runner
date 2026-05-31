# xlsm VBA 抽出・書き戻しツール

Excel マクロブック（`.xlsm`）から VBA ソースコードを抽出し、編集後に書き戻す Node.js スクリプト。

## ファイル構成

```
sample/excel/
├── extract-vba.mjs      # xlsm → .bas/.cls 抽出
├── import-vba.mjs       # .bas/.cls → xlsm 書き戻し
├── lib/
│   ├── ovba.mjs         # MS-OVBA 圧縮・伸長ライブラリ
│   └── dir-parser.mjs   # VBA Project dir ストリームパーサー
├── package.json
└── test.xlsm            # サンプル Excel マクロブック
```

## セットアップ

```bash
cd sample/excel
npm install
```

## 使い方

### VBA ソースの抽出

```bash
node extract-vba.mjs <input.xlsm> [出力ディレクトリ] [--encoding <cp>]
```

- 出力ディレクトリのデフォルトは `input.xlsm` と同じディレクトリ
- エンコーディングは xlsm の `PROJECTCODEPAGE` から**自動検出**する（`--encoding` で上書き可）
- 標準モジュール（`.bas`）とクラスモジュール（`.cls`）を UTF-8 で生成する

例:

```bash
node extract-vba.mjs test.xlsm ./src
# エンコーディング: cp932 (PROJECTCODEPAGE=932) のように自動検出を表示
# → ./src/Module1.bas, ThisWorkbook.cls, Sheet1.cls ...
```

### VBA ソースの書き戻し

```bash
node import-vba.mjs <input.xlsm> <ソースディレクトリ> [output.xlsm] [--encoding <cp>]
```

- `output.xlsm` のデフォルトは `input.xlsm`（上書き）
- エンコーディングは xlsm の `PROJECTCODEPAGE` から**自動検出**する（`--encoding` で上書き可）
- ソースディレクトリ内の `.bas` / `.cls` ファイルのうち、モジュール名が xlsm 内に存在するものだけを更新する

例:

```bash
node import-vba.mjs test.xlsm ./src ./test-modified.xlsm
```

---

## アーキテクチャ

### xlsm のファイル形式

`.xlsm` は ZIP アーカイブであり、VBA プロジェクトは `xl/vbaProject.bin` に格納されている。

```
test.xlsm (ZIP)
└── xl/
    └── vbaProject.bin   ← Compound File Binary (CFB) 形式
        └── VBA/
            ├── dir           ← モジュール一覧（OVBA 圧縮）
            ├── Module1       ← 各モジュールのソース（OVBA 圧縮）
            ├── ThisWorkbook
            └── ...
```

### 処理フロー

**抽出（extract-vba.mjs）:**

```
xlsm (ZIP)
  → vbaProject.bin (CFB)
    → VBA/dir ストリーム → OVBA 伸長 → parseDirStream()
        モジュール名 / ストリーム名 / ソース開始オフセット / モジュール種別
    → VBA/<モジュール名> ストリーム → offset バイトをスキップ → OVBA 伸長
        → .bas / .cls ファイルに書き出し
```

**書き戻し（import-vba.mjs）:**

```
.bas / .cls ファイル
  → OVBA 圧縮
  → CFB の VBA/<モジュール名> ストリームを更新
      （offset より前のコンパイル済みバイトコードは保持）
  → CFB をシリアライズ → ZIP に差し替え → xlsm として保存
```

### ライブラリ

#### `lib/ovba.mjs` — MS-OVBA 圧縮・伸長

[MS-OVBA] 2.4 に準拠した VBA プロジェクト専用の LZ77 圧縮。

```javascript
import { compress, decompress } from './lib/ovba.mjs';

const decompressed = decompress(compressedBuffer); // → Buffer
const compressed   = compress(sourceBuffer);       // → Buffer
```

コピートークンのビット分割は、チャンク内の出力済みバイト数に応じて動的に変化する（MS-OVBA 2.4.1.3.6）。

#### `lib/dir-parser.mjs` — dir ストリームパーサー

VBA Project の `dir` ストリームを解析してモジュール一覧を返す。

```javascript
import { parseDirStream } from './lib/dir-parser.mjs';

// buf は OVBA 伸長後の Buffer
const modules = parseDirStream(buf);
// → [{ name, streamName, offset, isClass }, ...]
```

| フィールド | 内容 |
|---|---|
| `name` | モジュール名（例: `"Module1"`） |
| `streamName` | CFB 内のストリーム名（通常 `name` と同じ） |
| `offset` | ストリーム内でソースが始まるバイトオフセット |
| `isClass` | `true` = クラス/ドキュメントモジュール（`.cls`）|

---

## 実装上の注意点

### CompressedChunkSize はヘッダー込みのサイズ（最重要）

各圧縮チャンクのヘッダー下位 12 ビット `CompressedChunkSize` は、
**2 バイトのヘッダーを含むチャンク全体のバイト数 − 3** を表す（[MS-OVBA] 2.4.1.1.5）。

```
データ長 = (header & 0x0FFF) + 3 - 2 = (header & 0x0FFF) + 1
```

これを `+ 3`（＝ データ長そのもの）と誤解すると、**チャンク境界が 2 バイトずれる**。
最初のチャンクは展開できるが、2 番目以降のチャンクのヘッダー位置がずれて
ソースが途中（4096 バイト境界付近）から壊れる。

> 注意: 抽出と書き戻しで同じ（誤った）式を使うと往復だけは一致してしまい、
> バグが隠れる。**元の Excel ファイルから抽出した結果が正しいか**を必ず確認すること。

### `PROJECTVERSION` レコードの特殊構造

`dir` ストリームのプロジェクトレベルレコードを順に読む際、
`PROJECTVERSION`（id = `0x0009`）は **size フィールドが常に 4 だが、
size 分のデータ（MajorVersion: 4 バイト）の直後に `MinorVersion`（2 バイト）が size 外として続く**。

```
id: 09 00          (2 bytes)
size: 04 00 00 00  (4 bytes) ← 常に 4
MajorVersion: ...  (4 bytes) ← size に含まれる
MinorVersion: ...  (2 bytes) ← size に含まれない！
```

この 2 バイトを読み飛ばさないと以降のレコード解析がずれる。

### 最終チャンクは raw を使わない

OVBA の圧縮チャンクには **compressed**（実際に圧縮）と **raw**（無圧縮、常に 4096 バイト）の 2 種類がある。

書き戻し時に最終（端数）チャンクを raw で書くと 4096 バイトへの null パディングが必要になり、
再抽出したソースに末尾 null が付く。

**本ツールでは最終チャンクに常に compressed を使用**することでこの問題を回避している。

### ソースのエンコーディング

VBA ソースは xlsm 内でプロジェクトのコードページでエンコードされている。
コードページは `dir` ストリームの `PROJECTCODEPAGE` レコード（id=0x0003）に格納されており、
両スクリプトはこの値を自動的に読み取る（日本語環境では 932 = Shift-JIS）。

- **抽出時**: `PROJECTCODEPAGE` のコードページ → UTF-8 に変換してファイルに書き出す
- **書き戻し時**: UTF-8 ファイルを読んで `PROJECTCODEPAGE` のコードページに変換してから OVBA 圧縮する

`--encoding` オプションで上書き可能（例: `--encoding cp1252`）。
エンコーディングの指定がなく `PROJECTCODEPAGE` も見つからない場合はエラーで終了する。

### ZIP のディレクトリエントリ

`zip.file()` はデフォルトで親ディレクトリ（`xl/` など）のエントリを自動生成する。
元の xlsm のエントリ構成を保つため、`createFolders: false` を指定している。

### PerformanceCache（バイトコード）の扱い

各モジュールストリームは `MODULEOFFSET` より前にコンパイル済みバイトコード
（PerformanceCache）、以降に圧縮ソースを持つ。本ツールはソース部分だけを差し替え、
バイトコード部分はそのまま保持する。

ソースを変更するとバイトコードと不整合になるが、Excel は通常この不整合を検出して
ソースから再コンパイルするため実用上は問題ない。

### ラウンドトリップの確認

抽出 → 書き戻し → 再抽出の 3 ステップで全モジュールのソースが完全一致することを確認済み。

```bash
node extract-vba.mjs test.xlsm /tmp/original
node import-vba.mjs  test.xlsm /tmp/original /tmp/modified.xlsm
node extract-vba.mjs /tmp/modified.xlsm /tmp/roundtrip
diff /tmp/original/Module1.bas /tmp/roundtrip/Module1.bas  # → 差分なし
```

なお、抽出結果の正しさはラウンドトリップだけでは保証されない（抽出・書き戻しが
同じ誤りを共有すると往復一致してしまう）。元の xlsm から抽出したソースが
実際の VBA コードと一致しているか目視確認することを推奨する。

---

## 依存パッケージ

| パッケージ | 用途 |
|---|---|
| [`cfb`](https://www.npmjs.com/package/cfb) | Compound File Binary (CFB) の読み書き |
| [`jszip`](https://www.npmjs.com/package/jszip) | ZIP（xlsm）の読み書き |
