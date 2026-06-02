# VBA Runner ドキュメント索引

このドキュメントは、`docs/` 配下のすべてのドキュメントを体系化し、読む順序と参照関係を示します。

---

## 📚 ドキュメント一覧

### 手法カタログ（横断参照）

| ドキュメント | 説明 |
|-----------|------|
| **[REFACTORING_TESTING_CATALOG.md](REFACTORING_TESTING_CATALOG.md)** | 各ドキュメントに登場するリファクタリング・テスト手法を名前付きで集約した参照用カタログ（R-01〜R-18、T-01〜T-15、S-01、VA-01/VA-01b/VA-02/VA-03/VA-04） |
| **[REFACTORING_CLOSURE.md](REFACTORING_CLOSURE.md)** | クロージャによるループ抽象化（キーブレイク例）— JavaScript 版・VBA クラス版の対比 |
| **[TABLE_DRIVEN_GUIDE.md](TABLE_DRIVEN_GUIDE.md)** | テーブル駆動リファクタリング（複雑な分岐を統一テーブルに） |
| **[FEATHERS_TECHNIQUES_TODO.md](FEATHERS_TECHNIQUES_TODO.md)** | Feathers "Working Effectively with Legacy Code" の手法一覧とカタログ収録状況 |
| **[LEGACY_REFACTORING_APPROACH.md](LEGACY_REFACTORING_APPROACH.md)** | レガシー VBA にどこから手を付けるか・手法選択の判断フロー（Feathers の考え方を VBA に適用した概要ガイド） |

### グループ1：テスト設計の原則（基礎）

| ドキュメント | 説明 | 対象読者 |
|-----------|------|--------|
| **[TESTING_STRATEGY.md](TESTING_STRATEGY.md)** | VBA テスト設計の 6 つの原則とパターン | すべての VBA 開発者 |
| **[REFACTORING_GUIDE.md](REFACTORING_GUIDE.md)** | Domain Logic と I/O の分離方法（4 つのパターン） | テストコードが書けない場合 |

### グループ2：実装レベルのテスト（実践）

| ドキュメント | 説明 | 対象読者 |
|-----------|------|--------|
| **[TEST_FRAMEWORK_GUIDE.md](TEST_FRAMEWORK_GUIDE.md)** | JavaScript テストフレームワークで VBA をテスト（VBARunner API・Time Mocking 含む） | テストコード実装者 |
| **[MOCK_GUIDE.md](MOCK_GUIDE.md)** | Excel オブジェクトのモック実装方法（3 段階） | Excel 依存を避けられない場合 |

### グループ3：具体例（学習用）

| ドキュメント | 説明 | 対象読者 |
|-----------|------|--------|
| **[REFACTORING_EXAMPLE.md](REFACTORING_EXAMPLE.md)** | TaskScheduler マクロのリファクタリング実例 | リファクタリング手法を学びたい |
| **[INTEGRATION_TEST_EXAMPLE.md](INTEGRATION_TEST_EXAMPLE.md)** | TaskScheduler の統合テスト実装例 | 統合テスト手法を学びたい |

### グループ4：内部実装の詳細（必要に応じて参照）

| ドキュメント | 説明 | 対象読者 |
|-----------|------|--------|
| **[internals/TYPE_SYSTEM_SPEC.md](internals/TYPE_SYSTEM_SPEC.md)** | VBA 型の仕様とVBA Runnerでの実装方法 | 型関連のバグ修正・機能追加時に参照 |
| **[internals/NAME_RESOLUTION.md](internals/NAME_RESOLUTION.md)** | 名前解決の仕様と実装（MS-VBAL §5.6.10） | 名前解決・スコープ関連のバグ修正時に参照 |
| **[internals/DATAFLOW_ANALYSIS_TODO.md](internals/DATAFLOW_ANALYSIS_TODO.md)** | データフロー解析の実装計画 | データフロー解析を実装する際に参照 |
| **[internals/IDENTIFIER_CLASSIFICATION_TODO.md](internals/IDENTIFIER_CLASSIFICATION_TODO.md)** | 識別子分類の正規化 TODO | 識別子解決の改善を行う際に参照 |

---

## 🎯 用途別ガイド

### シナリオ 1: VBA テストを始めたい

**推奨読み順序**:

```
1️⃣  TESTING_STRATEGY.md（20分）
   └─ テスト設計の基本原則を理解
   
2️⃣  TEST_FRAMEWORK_GUIDE.md（30分）
   └─ JavaScript でテストを書く方法を習得
   
3️⃣  REFACTORING_EXAMPLE.md（10分）
   └─ 実例を見て理解を深める
```

**参照フロー図**:
```
TESTING_STRATEGY.md
      ↓ 基礎を学ぶ
TEST_FRAMEWORK_GUIDE.md
      ↓ 実装方法を習得
実装開始 → REFACTORING_EXAMPLE.md（疑問に直面時に参照）
```

---

### シナリオ 2: レガシーコードをテスト可能にしたい

**推奨読み順序**:

```
1️⃣  TESTING_STRATEGY.md（20分）
   └─ なぜ分離が必要か理解
   
2️⃣  REFACTORING_GUIDE.md（30分）
   └─ リファクタリングの 4 パターンを習得
   
3️⃣  REFACTORING_EXAMPLE.md（15分）
   └─ TaskScheduler の実例でコンクリート化
   
4️⃣  TEST_FRAMEWORK_GUIDE.md（30分）
   └─ リファクタリング後のテスト実装
```

**参照フロー図**:
```
TESTING_STRATEGY.md（原則）
      ↓
REFACTORING_GUIDE.md（方法）
      ↓
REFACTORING_EXAMPLE.md（実例）
      ↓
TEST_FRAMEWORK_GUIDE.md（テスト実装）
      ↓
リファクタリング + テスト作成
```

---

### シナリオ 3: Excel オブジェクト依存は避けられない

**推奨読み順序**:

```
1️⃣  TESTING_STRATEGY.md（20分）
   └─ テスト設計の原則を理解
   
2️⃣  REFACTORING_GUIDE.md（30分）
   └─ ここまでで分離できないか再検討
   
3️⃣  MOCK_GUIDE.md（45分）
   └─ Part 1: MockWorksheet を使う
      Part 2: 自作モックの実装方法
      Part 3: 設計の一般論
   
4️⃣  INTEGRATION_TEST_EXAMPLE.md（15分）
   └─ 統合テストの実装例
```

**参照フロー図**:
```
TESTING_STRATEGY.md
      ↓ 分離を試みる
REFACTORING_GUIDE.md
      ↓ 分離が困難な場合
MOCK_GUIDE.md
      ├─ Part 1（実用）
      ├─ Part 2（応用）
      └─ Part 3（背景知識）
      ↓
INTEGRATION_TEST_EXAMPLE.md
      ↓
モック使用テスト作成
```

---

### シナリオ 4: 型の実装方法を確認したい（内部構造）

**使用場面**:
- TypeName() / VarType() の動作がおかしい
- 新しい型を追加したい
- 型変換ロジックを理解したい
- オーバーフロー検査を実装したい

**参照方法**:

問題に直面したときに **[TYPE_SYSTEM_SPEC.md](internals/TYPE_SYSTEM_SPEC.md)** を参照：

```
internals/TYPE_SYSTEM_SPEC.md
（VBA 型仕様 + VBA Runnerでの実装方法）
```

**参照フロー図**:
```
問題発生
      ↓
internals/TYPE_SYSTEM_SPEC.md を参照
（VBA 型の仕様と実装方法を確認）
      ↓
evaluator.ts や Environment クラスのコード実装に反映
```

---

## 📊 ドキュメント間の参照関係

### 依存グラフ

```
TESTING_STRATEGY.md（最上位の原則）
    ↓
    ├─→ REFACTORING_GUIDE.md
    │       ↓
    │       ├─→ REFACTORING_EXAMPLE.md
    │       └─→ TEST_FRAMEWORK_GUIDE.md
    │
    ├─→ TEST_FRAMEWORK_GUIDE.md（テスト実装）
    │
    ├─→ MOCK_GUIDE.md
    │       ↓
    │       └─→ INTEGRATION_TEST_EXAMPLE.md
    │
    └─→ internals/TYPE_SYSTEM_SPEC.md（リファレンス）
```

### 詳細な参照関係

| ドキュメント | 参照先 | 理由 |
|-----------|--------|------|
| [REFACTORING_TESTING_CATALOG.md](REFACTORING_TESTING_CATALOG.md) | [TESTING_STRATEGY.md](TESTING_STRATEGY.md), [REFACTORING_GUIDE.md](REFACTORING_GUIDE.md), [TEST_FRAMEWORK_GUIDE.md](TEST_FRAMEWORK_GUIDE.md), [MOCK_GUIDE.md](MOCK_GUIDE.md) | R/T/S/VA 手法の一覧。各手法の詳細説明ドキュメントへのリンクを提供 |
| [REFACTORING_GUIDE.md](REFACTORING_GUIDE.md) | [TESTING_STRATEGY.md](TESTING_STRATEGY.md) | テスト設計原則を前提に実装パターンを提示 |
| [REFACTORING_EXAMPLE.md](REFACTORING_EXAMPLE.md) | [REFACTORING_GUIDE.md](REFACTORING_GUIDE.md) | 4 つのリファクタリングパターンの実例 |
| [TEST_FRAMEWORK_GUIDE.md](TEST_FRAMEWORK_GUIDE.md) | [TESTING_STRATEGY.md](TESTING_STRATEGY.md) | テスト設計原則を実装に落とし込む |
| [MOCK_GUIDE.md](MOCK_GUIDE.md) | [TESTING_STRATEGY.md](TESTING_STRATEGY.md), [REFACTORING_GUIDE.md](REFACTORING_GUIDE.md) | Domain Logic 分離を最優先、それでもダメな場合 |
| [INTEGRATION_TEST_EXAMPLE.md](INTEGRATION_TEST_EXAMPLE.md) | [TESTING_STRATEGY.md](TESTING_STRATEGY.md), [MOCK_GUIDE.md](MOCK_GUIDE.md) | 統合テスト設計とモック活用 |

---

## 📖 各ドキュメントの概要と主要セクション

### 1. [TESTING_STRATEGY.md](TESTING_STRATEGY.md)
**目的**: VBA テスト設計の原則とパターン  
**対象**: すべての VBA 開発者

**主要セクション**:
- 原則 1: Domain Logic と Excel I/O の徹底的分離
- 原則 2-6: テスト可能な設計のパターン
- テストピラミッド（Unit → Integration → E2E）

**主なメッセージ**:
> Domain Logic と I/O を分離することが、テスト可能で保守性の高いコードへの第一歩

---

### 2. [REFACTORING_GUIDE.md](REFACTORING_GUIDE.md)
**目的**: 既存コードをテスト可能にするリファクタリング方法  
**対象**: 「テストが書けない」という状況に直面している人

**主要セクション**:
- 4 つのリファクタリングパターン（抽出、適応、段階的、テスト駆動）
- 実装チェックリスト
- よくある質問

**主なメッセージ**:
> リファクタリングは「難しい変更」ではなく「段階的な改善」

---

### 3. [REFACTORING_EXAMPLE.md](REFACTORING_EXAMPLE.md)
**目的**: TaskScheduler マクロのリファクタリング実例を通じた学習  
**対象**: リファクタリング手法を具体例で学びたい人

**主要セクション**:
- TaskScheduler_v1.bas の処理概要（393 行の複雑さを分解）
- Excel との密結合によるテスト困難性
- リファクタリング前後の比較
- 改善された設計（UDT、関数分割）

**主なメッセージ**:
> 実業務で必要な複雑なロジックも、分離と分割で理解しやすく、テスト可能にできる

---

### 4. [TEST_FRAMEWORK_GUIDE.md](TEST_FRAMEWORK_GUIDE.md)
**目的**: JavaScript テストフレームワークで VBA をテストする方法  
**対象**: テストコード実装者

**主要セクション**:
- Jest との統合方法
- 7 つのテストパターン（パターン7: Time Mocking 含む）
- **VBARunner API リファレンス**（`run` / `eval` / `set` / `mockDate` / `registerExternalObject`）、**assert ユーティリティ**（`strictEqual` / `deepStrictEqual` / `ok` / `isTrue` / `isFalse` / `fail`）
- エラーデバッグ情報（`err.vbaLine` / `err.number`）
- よくある質問

**主なメッセージ**:
> JavaScript テストフレームワークを活用することで、VBA テストが身近で効率的になる

---

### 5. [MOCK_GUIDE.md](MOCK_GUIDE.md)
**目的**: Excel オブジェクトのモック実装方法（3 段階学習）  
**対象**: Excel 依存を避けられない場合

**主要セクション**:
- **Part 1**: VBA Runnerの MockWorksheet を使う（実用）
- **Part 2**: MockWorksheet を参考に自分でモックを実装（応用）
- **Part 3**: モック設計の一般論（背景知識）

**主なメッセージ**:
> モックは「最後の手段」。まず Domain Logic 分離を試みよ

---

### 6. [INTEGRATION_TEST_EXAMPLE.md](INTEGRATION_TEST_EXAMPLE.md)
**目的**: 統合テストの実装例（TaskScheduler を使用）  
**対象**: 統合テスト手法を学びたい人

**主要セクション**:
- 単体テスト vs 統合テスト
- TaskScheduler の 7 つの統合テストシナリオ
- テスト実行方法と結果検証

**主なメッセージ**:
> 統合テストで全体の動作を確認。単体テストで詳細を検証。

---

### 7. [TYPE_SYSTEM_SPEC.md](internals/TYPE_SYSTEM_SPEC.md)
**目的**: VBA 型の仕様とVBA Runnerでの実装方法（必要に応じて参照）  
**対象**: 型関連のバグ修正・機能追加を行う人

**主要セクション**:
- VBA 型の基本仕様
- VBA Runnerでの型情報の保持方式
- TypeName() / VarType() の実装方法
- オーバーフロー検査
- 型変換の規則

**主なメッセージ**:
> 型の動作がおかしい / 新しい型を追加したい、という具体的な問題発生時に参照

---

## 🗂️ ドキュメント構成の論理

```
基礎層（原則）
  ↓
TESTING_STRATEGY.md
  ├─ テスト設計の基本原則 6 個
  └─ テストピラミッド構造
  
  ↓
  
実装層（方法）
  ├─ REFACTORING_GUIDE.md
  │   └─ Domain Logic 分離の 4 パターン
  │
  ├─ TEST_FRAMEWORK_GUIDE.md
  │   └─ JavaScript テストフレームワーク活用
  │
  └─ MOCK_GUIDE.md
      └─ Excel モック実装（3 段階）
  
  ↓
  
具体例層（学習用）
  ├─ REFACTORING_EXAMPLE.md
  │   └─ TaskScheduler リファクタリング実例
  │
  └─ INTEGRATION_TEST_EXAMPLE.md
      └─ TaskScheduler 統合テスト実例
  
  ↓
  
内部実装層（必要に応じて参照）
  └─ internals/TYPE_SYSTEM_SPEC.md
      └─ VBA 型仕様 + 実装方法
         （バグ修正・機能追加時に参照）
```

---

## ⏱️ 推奨学習時間

### 最小限（急ぐ場合）: 1 時間
- [TESTING_STRATEGY.md](TESTING_STRATEGY.md) (20分)
- [TEST_FRAMEWORK_GUIDE.md](TEST_FRAMEWORK_GUIDE.md) (40分)

### 標準（推奨）: 2-3 時間
- [TESTING_STRATEGY.md](TESTING_STRATEGY.md) (20分)
- [REFACTORING_GUIDE.md](REFACTORING_GUIDE.md) (30分)
- [TEST_FRAMEWORK_GUIDE.md](TEST_FRAMEWORK_GUIDE.md) (40分)
- [REFACTORING_EXAMPLE.md](REFACTORING_EXAMPLE.md) (15分)

### 完全（全体を理解したい）: 4-5 時間
- すべてのドキュメント
- 特に [MOCK_GUIDE.md](MOCK_GUIDE.md) の Part 3（設計の一般論）に時間をかける

### 深掘り（実装者向け）: 8-10 時間
- すべてのドキュメント（INDEX.md を除く）
- [REFACTORING_EXAMPLE.md](REFACTORING_EXAMPLE.md) で実装を読む
- [INTEGRATION_TEST_EXAMPLE.md](INTEGRATION_TEST_EXAMPLE.md) でテストコード全体を読む
- 必要に応じて [TYPE_SYSTEM_SPEC.md](internals/TYPE_SYSTEM_SPEC.md) を参照

---

## 🔍 問題別ドキュメント検索

### 「テストを書きたい」
→ **[TEST_FRAMEWORK_GUIDE.md](TEST_FRAMEWORK_GUIDE.md)** から開始

### 「レガシーコードがテスト不可能」
→ **[REFACTORING_GUIDE.md](REFACTORING_GUIDE.md)** → **[REFACTORING_EXAMPLE.md](REFACTORING_EXAMPLE.md)**

### 「Excel オブジェクトが避けられない」
→ **[MOCK_GUIDE.md](MOCK_GUIDE.md)**（Part 1 から始める）

### 「テスト設計の原則を理解したい」
→ **[TESTING_STRATEGY.md](TESTING_STRATEGY.md)** 全体

### 「`Now()` など日時依存のコードをテストしたい」
→ **[TEST_FRAMEWORK_GUIDE.md — パターン7](TEST_FRAMEWORK_GUIDE.md)** (`vbaRunner.mockDate()`)

### 「VBARunner の API を確認したい」
→ **[TEST_FRAMEWORK_GUIDE.md — VBARunner API リファレンス](TEST_FRAMEWORK_GUIDE.md)**

### 「型の動作がおかしい / 型機能を追加したい」
→ **[TYPE_SYSTEM_SPEC.md](internals/TYPE_SYSTEM_SPEC.md)**（内部実装の詳細を参照）

### 「統合テストを書きたい」
→ **[INTEGRATION_TEST_EXAMPLE.md](INTEGRATION_TEST_EXAMPLE.md)**

### 「リファクタリングの実例が見たい」
→ **[REFACTORING_EXAMPLE.md](REFACTORING_EXAMPLE.md)**

---

## 📝 ドキュメント更新ガイド

### 新しいドキュメントを追加する場合

1. 論理的な層を決定（基礎 / 実装 / 具体例 / リファレンス）
2. 参照先ドキュメントを明示（「前提：〜を読む」）
3. このファイルの参照関係セクションを更新

### 既存ドキュメントを更新する場合

1. 内容の大幅変更の場合、参照関係を確認
2. 参照先・被参照がある場合、相手側も確認
3. セクション追加時は、このファイルの概要セクションも更新

---

## 🎓 まとめ

VBA Runnerのドキュメントは、以下の層で構成されています：

1. **基礎層**: テスト設計原則（[TESTING_STRATEGY.md](TESTING_STRATEGY.md)）
2. **実装層**: リファクタリング、テストフレームワーク、モック（3 ドキュメント）
3. **具体例層**: TaskScheduler の実例（2 ドキュメント）
4. **内部実装層**: VBA 型の仕様と実装方法（[TYPE_SYSTEM_SPEC.md](internals/TYPE_SYSTEM_SPEC.md)）
   - 通常は「必要に応じて参照」する

**学習の流れ**:
```
基礎層（原則）
  ↓
実装層（方法）
  ↓
具体例層（実例）
  
※ 内部実装層は「問題発生時に参照」
```

**最も重要なドキュメント**:
- **入門者**: [TESTING_STRATEGY.md](TESTING_STRATEGY.md) → [TEST_FRAMEWORK_GUIDE.md](TEST_FRAMEWORK_GUIDE.md)
- **リファクタリング者**: [REFACTORING_GUIDE.md](REFACTORING_GUIDE.md) → [REFACTORING_EXAMPLE.md](REFACTORING_EXAMPLE.md)
- **モック活用者**: [MOCK_GUIDE.md](MOCK_GUIDE.md)（3 段階で学習）
- **型バグ修正・追加**: [TYPE_SYSTEM_SPEC.md](internals/TYPE_SYSTEM_SPEC.md)（参照）

用途に応じて、このインデックスを参考に適切なドキュメントから読み始めてください。
