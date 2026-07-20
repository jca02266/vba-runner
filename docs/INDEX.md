# docs/INDEX.md — 詳細索引（任意）

> 入口は [README.md](README.md) です。
>
> このファイルは、読み順・問題別検索・参照関係など **詳細索引** 専用です。
> 初めての方は README だけで十分です。

対象: 利用者（テスト・リファクタリング実践者）向け。エンジン本体の開発は [CONTRIBUTING.md](../CONTRIBUTING.md) へ。

---

## シナリオ別の読み順

### 1. VBA テストを始めたい

```
TESTING_STRATEGY.md
      ↓
TEST_FRAMEWORK_GUIDE.md
      ↓
実装開始 → REFACTORING_EXAMPLE.md（疑問時）
```

### 2. レガシーコードをテスト可能にしたい

```
TESTING_STRATEGY.md
      ↓
REFACTORING_GUIDE.md
      ↓
REFACTORING_EXAMPLE.md
      ↓
TEST_FRAMEWORK_GUIDE.md
```

### 3. Excel オブジェクト依存を避けられない

```
TESTING_STRATEGY.md
      ↓
REFACTORING_GUIDE.md（まず分離を再検討）
      ↓
MOCK_GUIDE.md（Part 1 → 2 → 3）
      ↓
INTEGRATION_TEST_EXAMPLE.md
```

### 4. 型の実装を確認したい（開発者向け）

問題発生時に [internals/TYPE_SYSTEM_SPEC.md](internals/TYPE_SYSTEM_SPEC.md) を参照。

---

## 問題別検索

| やりたいこと | 開始文書 |
|---|---|
| テストを書きたい | [TEST_FRAMEWORK_GUIDE.md](TEST_FRAMEWORK_GUIDE.md) |
| レガシーコードがテスト不可能 | [REFACTORING_GUIDE.md](REFACTORING_GUIDE.md) → [REFACTORING_EXAMPLE.md](REFACTORING_EXAMPLE.md) |
| Excel オブジェクトが避けられない | [MOCK_GUIDE.md](MOCK_GUIDE.md)（Part 1 から） |
| テスト設計の原則を理解したい | [TESTING_STRATEGY.md](TESTING_STRATEGY.md) |
| `Now()` など日時依存をテストしたい | [TEST_FRAMEWORK_GUIDE.md](TEST_FRAMEWORK_GUIDE.md)（パターン7） |
| VBARunner API を確認したい | [TEST_FRAMEWORK_GUIDE.md](TEST_FRAMEWORK_GUIDE.md)（API リファレンス） |
| 統合テストを書きたい | [INTEGRATION_TEST_EXAMPLE.md](INTEGRATION_TEST_EXAMPLE.md) |
| リファクタリング実例が見たい | [REFACTORING_EXAMPLE.md](REFACTORING_EXAMPLE.md) |
| 手法名で横断検索したい | [REFACTORING_TESTING_CATALOG.md](REFACTORING_TESTING_CATALOG.md) |
| 型の動作がおかしい / 型機能を追加したい | [internals/TYPE_SYSTEM_SPEC.md](internals/TYPE_SYSTEM_SPEC.md) |

---

## 文書間の依存関係（簡略）

```
TESTING_STRATEGY.md（原則）
    ├─→ REFACTORING_GUIDE.md
    │       ├─→ REFACTORING_EXAMPLE.md
    │       └─→ TEST_FRAMEWORK_GUIDE.md
    ├─→ TEST_FRAMEWORK_GUIDE.md
    └─→ MOCK_GUIDE.md
            └─→ INTEGRATION_TEST_EXAMPLE.md
```

| 文書 | 主な前提 |
|---|---|
| [REFACTORING_GUIDE.md](REFACTORING_GUIDE.md) | TESTING_STRATEGY |
| [REFACTORING_EXAMPLE.md](REFACTORING_EXAMPLE.md) | REFACTORING_GUIDE |
| [TEST_FRAMEWORK_GUIDE.md](TEST_FRAMEWORK_GUIDE.md) | TESTING_STRATEGY |
| [MOCK_GUIDE.md](MOCK_GUIDE.md) | TESTING_STRATEGY、可能なら REFACTORING_GUIDE |
| [INTEGRATION_TEST_EXAMPLE.md](INTEGRATION_TEST_EXAMPLE.md) | TESTING_STRATEGY、MOCK_GUIDE |
| [REFACTORING_TESTING_CATALOG.md](REFACTORING_TESTING_CATALOG.md) | 上記実践ガイド群への横断索引 |

目的別の入口・パッケージ正本への導線は [README.md](README.md) を参照してください。
リポジトリ全体のファイル台帳は [ルート INDEX.md](../INDEX.md)（メンテナ向け）です。
