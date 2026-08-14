---
type: Design Note
title: 仮想ファイルシステムとNodeFileSystem
description: MemoryFileSystemとNodeFileSystemの実装差、保証範囲、利用境界を定義する。
tags: [vba, filesystem, sandbox, memory, node]
status: stable
generated:
  by: human:vba-runner-maintainers
  at: 2026-08-15T00:00:00+09:00
verified:
  - by: human:vba-runner-maintainers
    at: 2026-08-15T00:00:00+09:00
sources:
  - id: filesystem-interface
    resource: ../../src/engine/filesystem.ts
    title: FileSystem / MemoryFileSystem 実装
  - id: node-filesystem
    resource: ../../src/engine/node_filesystem.ts
    title: NodeFileSystem 実装
  - id: sandbox-path
    resource: ../../src/engine/sandbox.ts
    title: SandboxPath 実装
  - id: reference
    resource: ../implementation/REFERENCE.md
    title: VBA Runner 詳細仕様
  - id: testing-guide
    resource: ../testing/TEST_FRAMEWORK_GUIDE.md
    title: VBA テストフレームワーク活用ガイド
---

# 仮想ファイルシステムとNodeFileSystem

## 目的と位置づけ

VBA RunnerのファイルI/Oは、`FileSystem`インターフェースを介して
Evaluatorから分離されています。
通常の評価では、ホストOSに触れない`MemoryFileSystem`を使用します。
Node.jsの実ファイルを扱う必要があるCLIや統合テストだけが、
`NodeFileSystem`を明示的に注入します。

ここでいう「ファイルシステム実装（バックエンド）」とは、Evaluatorから
呼び出される`FileSystem`インターフェースの実装を指します。
この文書は、両ファイルシステム実装の違いと「VBA Runnerが守る契約」と
「実装上保証しないホスト依存の挙動」を分けて記録します。
VBAの各ステートメントの仕様一覧は
[VBA_SPEC_LIST.md](../implementation/VBA_SPEC_LIST.md)、利用者向けの
Sandbox説明は[REFERENCE.md](../implementation/REFERENCE.md)を参照してください。

## 実装構成

```text
VBA組み込みI/O / FileSystemObject
              │
              ▼
          Evaluator
              │  FileSystemインターフェース
       ┌──────┴────────┐
       ▼               ▼
MemoryFileSystem   NodeFileSystem
 (メモリのみ)       (Node fs / 実ファイル)
```

`SandboxPath`はVBAパスをファイルシステム実装へ渡す実パスへ変換し、相対パスと
ドライブ形式のパスをサンドボックスルート配下へ限定します。
ただし、`SandboxPath`はOSのアクセス権やシンボリックリンクを隔離する
機構ではありません。

## ファイルシステム実装（バックエンド）の差異

| 観点 | MemoryFileSystem | NodeFileSystem |
|---|---|---|
| データの所在 | JavaScriptメモリ上 | Node.jsホストのファイルシステム |
| 通常のEvaluator | デフォルト | 明示注入時のみ |
| `hostBacked` | `false` | `true` |
| 初期ルート | 仮想ディレクトリとして作成 | 呼び出し側が作成・指定 |
| ブラウザ利用 | 可能 | 不可（Node `fs`依存） |
| テスト間の分離 | インスタンス単位で容易 | 一時ディレクトリ管理が必要 |
| OS権限・属性 | エンジンが定義した近似値 | OSの権限・属性に依存 |
| 外部からの永続化 | なし | あり |
| 主な用途 | 通常評価、単体テスト、ブラウザ | CLI、実ファイル統合テスト |

## 守る契約

### 共通インターフェース

両ファイルシステム実装は、同期I/Oの次の契約を共有します。

- ファイルの存在確認、読み書き、作成、削除
- ディレクトリ作成、一覧、削除
- `stat`相当のサイズ・時刻・種別取得
- ファイルハンドルのopen/read/write/close
- ファイル・ディレクトリのcopy/move
- VBAの`GetAttr`/`SetAttr`に対応する属性境界

Evaluator側の`Open`、`Kill`、`MkDir`、`FileCopy`、
`FileSystemObject`などは、直接ファイルシステム実装を分岐せず、共通の
ファイル操作ヘルパーと`FileSystem`インターフェースを使用します。

### パス境界

- VBAの相対パスは仮想カレントディレクトリから解決する
- `C:\foo`のようなドライブ形式はサンドボックス配下の`c/foo`へ変換する
- `..`などでサンドボックスルートの外へ出るパスは拒否する
- MemoryFileSystemでは、作成したファイル・ディレクトリをメモリ内だけに保持する

### MemoryFileSystem固有の再現契約

MemoryFileSystemは単なるモックではなく、テスト可能なVBA I/Oの実装です。
次の境界を実装・回帰テストの対象としています。

- `r`、`r+`、`w`、`a`のアクセスモード
- Append時の物理的な末尾追記
- 読み取り専用ハンドルへの書込み拒否
- 無効なハンドルの二重close拒否
- ゼロ長読み書き
- Uint8Arrayのコピー時の所有権分離
- ディレクトリcopy/moveの自己包含拒否と事前検査
- 固定配列・UDT・Binary/Random I/OのVBA向け境界

## 守らない契約・意図的な制限

### Excel/Windowsの完全再現ではない

MemoryFileSystemは、WindowsのNTFSやExcel VBAの全ての副作用を再現する
ものではありません。特に次は、仕様またはホスト依存のため完全一致を
保証しません。

- Windows固有のACL、所有者、共有モード、ロック待ち時間
- 大文字小文字の扱い、Unicode正規化、コードページの全組合せ
- `mtime`やファイル属性の細かなOS差
- シンボリックリンク、ハードリンク、ジャンクション
- OSプロセス間で共有されるファイルハンドルの寿命
- Nodeの全エラー番号・エラー文言とVBAのエラー番号の完全一致

Excel実機との物理レイアウト差や境界は、必要な場合に
`tests/excel/queue`の実機プローブで別途確認します。

### NodeFileSystemは安全な仮想化層ではない

`NodeFileSystem`を使うと、変換後のパスは実際のホストOSへ到達します。
`SandboxPath`の範囲検査は通常のパス文字列を対象にしたものであり、
OSレベルの隔離や攻撃耐性を提供するものではありません。
信頼できないVBAを実行する用途では、NodeFileSystemを使用せず、
MemoryFileSystemまたはOSコンテナ等の別の隔離を使用してください。

## 利用ルール

### 通常の評価・テスト

ファイルシステムを指定しないでください。EvaluatorはMemoryFileSystemを
自動的に使用します。Evaluatorを複数生成した場合も、それぞれが別の
MemoryFileSystemインスタンスを持つため、ファイル状態はEvaluator間で
共有されません。

```ts
const ev = evalVBASingle(source);
```

複数Evaluatorで同じファイル状態を参照したいテストだけ、同じ
MemoryFileSystemインスタンスを各Evaluatorへ明示注入します。

```ts
const sharedFs = new MemoryFileSystem();
const first = evalVBASingle(writerSource, { fs: sharedFs });
const second = evalVBASingle(readerSource, { fs: sharedFs });
// firstとsecondは同じsharedFs上のファイルを参照する
```

Evaluatorごとに状態を分離したい場合は、`fs`を注入せず、それぞれの
デフォルトMemoryFileSystemを使用してください。

### 明示的なホストI/O

NodeFileSystemを使う場合は、既存の実ディレクトリを汚染しない専用の
`sandboxRoot`を必ず指定します。既定の`/sandbox`は拒否されます。

```ts
const ev = evalVBASingle(source, {
  fs: new NodeFileSystem(),
  sandboxRoot: temporaryDirectory,
});
```

CLIの`vba-run`はこのルールに従い、Nodeの一時ディレクトリを生成して
明示的に渡します。NodeFileSystemを直接注入するコードで
`sandboxRoot`を省略することはできません。

## 類似文書との役割分担

既存文書にも関連情報はありますが、両ファイルシステム実装の比較と保証境界を
一つにまとめた文書はありません。

- [REFERENCE.md](../implementation/REFERENCE.md): 利用者向けAPI、Sandbox、
  MemoryFileSystemの基本利用方法。NodeFileSystemとの安全境界は本書が補完する。
- [VBA_SPEC_LIST.md](../implementation/VBA_SPEC_LIST.md): FileSystem/VFSの
  実装項目と進捗。個別バグの履歴ではなく仕様実装一覧を管理する。
- [TEST_FRAMEWORK_GUIDE.md](../testing/TEST_FRAMEWORK_GUIDE.md):
  MemoryFileSystemを使ったテスト方法。ファイルシステム実装の設計正本ではない。
- [VBA_ENGINE_BUG_HISTORY.md](../bug-records/VBA_ENGINE_BUG_HISTORY.md):
  過去のファイルI/Oバグと修正履歴。本書は現在の設計契約を記述する。

したがって、実装契約・安全境界は本書、利用方法は`REFERENCE.md`、
仕様項目の進捗は`VBA_SPEC_LIST.md`、過去の不具合はバグ履歴を参照します。

## 実装・検証の入口

- 実装: [`src/engine/filesystem.ts`](../../src/engine/filesystem.ts)、
  [`src/engine/node_filesystem.ts`](../../src/engine/node_filesystem.ts)、
  [`src/engine/sandbox.ts`](../../src/engine/sandbox.ts)
- Memory/FileSystem回帰: [`tests/spec/filesystem.test.ts`](../../tests/spec/filesystem.test.ts)、
  [`tests/spec/filesystem-extra.test.ts`](../../tests/spec/filesystem-extra.test.ts)
- NodeFileSystemは、明示した一時ルートを使う統合テストだけで検証する
