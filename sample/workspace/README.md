# VBA サンプル - Excel データ処理

## ファイル

| ファイル | 内容 |
|---|---|
| `DataProcessor.bas` | メインの VBA モジュール |

## 機能

1. **サンプルデータ生成** - `Data` シートに商品売上データを自動生成
2. **重複行の除去** - ID が重複している行を取り除いてコピー
3. **カテゴリ別集計** - カテゴリごとに件数・売上合計・平均売上を計算
4. **結果の出力** - 処理結果を `Result` シートに整形して出力

## 使い方

### Excel へのインポート

1. Excel を開き、`Alt + F11` で VBA エディターを起動
2. メニュー「ファイル」→「ファイルのインポート」で `DataProcessor.bas` を選択
3. エディターを閉じる

### 実行

- `Alt + F8` でマクロ一覧を開き `ProcessData` を選択して「実行」
- または VBA エディター上でカーソルを `ProcessData` 内に置いて `F5`

### 実行結果

```
Data シート   : サンプルデータ(9行・重複1件含む)
Result シート : 重複除去済みデータ + カテゴリ別集計
```

## コード構成

```
ProcessData()          ← エントリポイント
├─ CreateSampleData()  ← Data シートにサンプルデータ投入
├─ CopyFilteredData()  ← 重複除去してコピー
├─ SummarizeByCategory() ← カテゴリ別集計
└─ FormatResultSheet() ← 書式整形
```

## 学習ポイント

- `Scripting.Dictionary` を使った重複チェックと集計
- シートの取得・作成(`GetOrCreateSheet`)
- `End(xlUp).Row` による最終行取得
- セルの書式設定(`Interior.Color`, `NumberFormat`)
- `Option Explicit` による変数宣言の強制
