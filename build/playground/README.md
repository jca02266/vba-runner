# VBA Runner Playground — Web UI デモ

ブラウザ上で VBA コードを実行・確認できる Web UI デモです。
インストール不要で、VBA Runner の動作をすぐに試すことができます。

## オンラインデモ

👉 https://vba-web-runner.netlify.app/

`Debug.Print` の結果表示や構文エラーの確認など、コードを書きながら即座に動作を試すことができます。

## ローカルで起動する

```bash
# リポジトリルートから
npm run dev --prefix build/playground

# または build/playground ディレクトリから
cd build/playground
npm run dev
```

起動後は http://localhost:5173/ にアクセスします。

## ビルド

```bash
npm run build:playground   # build/playground/dist/ に出力
```

## 詳細ドキュメント

- [README.md](../../README.md) — プロジェクト概要・クイックスタート
- [REFERENCE.md](../../REFERENCE.md) — 詳細仕様
