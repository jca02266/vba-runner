# VBA Web Runner — Web UI デモ

ブラウザ上で VBA コードを実行・確認できる Web UI デモです。
インストール不要で、VBA Runner の動作をすぐに試すことができます。

> 対象: パッケージ利用者
>
> 目的別のガイド一覧: [ドキュメントガイド](../../docs/README.md)

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

- [ドキュメントガイド](../../docs/README.md) — 目的別の共通ハブ（利用者／開発者）
- [README.md](../../README.md) — プロジェクト概要
- [REFERENCE.md](../../REFERENCE.md) — 詳細仕様
