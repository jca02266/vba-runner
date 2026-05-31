# sample/workspace パッケージテストワークスペース

## 目的

`sample/workspace/` を Extension Development Host の独立ワークスペースとして使い、
ソースを直接参照せずに npm パッケージとしての動作を検証する。

## ディレクトリ構成

```
vba-compiler/              ← モノレポ
├── src/                   ← エンジンのソース（TypeScript）
├── dist/
│   ├── extension.cjs      ← VS Code 拡張機能
│   ├── lib.cjs            ← npm ライブラリ（VBARunner / assert）
│   └── bin/               ← CLI ツール群
├── sample/workspace       ← このワークスペース
│   ├── node_modules/
│   │   └── vba-runner/    ← npm pack でインストールされたパッケージ
│   ├── package.json
│   ├── DataProcessor.bas
│   └── VariantDemo.bas
└── .vscode/
    ├── launch.json        ← "Run Extension (sample-workspace)" 設定済み
    └── tasks.json         ← "setup: sample-workspace" タスク設定済み
```

## F5 で起動する

1. VS Code のデバッグパネルで **"Run Extension (sample-workspace)"** を選択
2. F5 を押す

F5 を押すと `preLaunchTask: "setup: sample-workspace"` が自動で以下を実行する。

```
1. npm run build:lib      → dist/lib.cjs を生成
2. npm run build:bin      → dist/bin/vba-*.cjs を生成（shebang 付き）
3. npm pack               → vba-runner-x.x.x.tgz を生成
4. npm install --prefix sample/workspace  → node_modules/vba-runner/ に展開
5. npm run build:extension → dist/extension.cjs を生成
```

完了後、Extension Development Host が `sample/workspace/` のみを開いた状態で起動する。

## この構成で何が保証されるか

| 観点 | 状態 |
|------|------|
| EH ウィンドウが参照するファイル | `sample/workspace/` 内のファイルのみ |
| 拡張機能のコード | `dist/extension.cjs`（`src/` の TypeScript ではない） |
| VBARunner / assert の参照先 | `node_modules/vba-runner/dist/lib.cjs`（ソースではない） |
| モノレポの他のソース | EH ウィンドウからは見えない |

## VBARunner と assert の使い方

このディレクトリ内のスクリプトから次のように使える。

```js
const { VBARunner, assert } = require('vba-runner');

const runner = new VBARunner('./DataProcessor.bas');
const result = runner.run('ProcessData', []);
assert.strictEqual(result, 'expected value');
```

## なぜ `file:` や `npm link` ではなく `npm pack` を使うか

| 方法 | ソース参照の可能性 | publish 後との差異 |
|------|-------------------|--------------------|
| `file:` | TypeScript のパス解決でソースを直接参照することがある | あり |
| `npm link` | シンボリックリンク経由でソースが見える | あり |
| `npm pack` + install | `node_modules/` に展開されたファイルのみを参照 | なし |

publish 前の最終確認は必ず `npm pack` で行う。
