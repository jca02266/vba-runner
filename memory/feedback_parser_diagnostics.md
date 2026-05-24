---
name: feedback-parser-diagnostics
description: パーサーのエラーは throw されず diagnostics に記録される。try-catch では検出できない
metadata:
  type: feedback
---

パーサーはエラーを `throw` せず `ast.diagnostics[]` に記録し、`ast.body` は空または不完全な状態で返す。

**Why:** `tryParse` で try-catch のみ確認し `Width: OK` と誤判断した実例あり（実際は `body: []` でパース失敗していた）。

**How to apply:** パーサー結果を確認するときは必ず以下を両方チェックする:
1. `ast.diagnostics.length === 0`（エラーなし）
2. `ast.body` に期待するノードが含まれているか

```typescript
// 誤り（例外のみ見る）
try { new Parser(...).parse(); console.log('OK'); } catch(e) { ... }

// 正しい
const ast = new Parser(...).parse();
if (ast.diagnostics.length > 0) { /* エラーあり */ }
```
