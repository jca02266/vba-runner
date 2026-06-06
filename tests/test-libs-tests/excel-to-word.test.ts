/**
 * Excel VBA から CreateObject("Word.Application") で Word を操作するテスト
 *
 * injectWordStub により CreateObject("Word.Application") が MockWordApplication を
 * 返すため、Excel 依存コードを追加の設定なしで実行できる。
 */
import { VBARunner, assert } from '../../test-libs/test-runner';
import { injectWordStub } from '../../test-libs/word-stub';

const runner = new VBARunner('tests/fixtures/excel-to-word');
injectWordStub(runner.evaluator);

// --- CreateObject("Word.Application") で文書を作成してテキストを設定 ---
assert.strictEqual(
    runner.run('CreateWordDoc', []),
    'Hello from Excel VBA',
    'CreateObject("Word.Application").Documents.Add().Content.Text'
);
console.log('[PASS] CreateObject("Word.Application") → Documents.Add() → Content.Text');

// --- ブックマークにテキストを差し込む ---
assert.strictEqual(
    runner.run('FillTemplate', []),
    'Acme Corp / 2024-03-15',
    'Bookmarks("bmk").Range.Text への差し込み'
);
console.log('[PASS] Bookmarks への差し込み');

// --- 複数ドキュメントの Documents.Count ---
assert.strictEqual(
    runner.run('ProcessMultipleDocs', []),
    // MockDocumentsCollection は Add() のたびに増える。
    // Word Application インスタンスがそれぞれ独立している（新規 new）ので
    // 各呼び出しで新しい app が作られ Add() 2回 → Count = 3（初期 1 + Add 2）
    3,
    'Documents.Count'
);
console.log('[PASS] Documents.Add() × 2 → Count = 3');

// --- setContent で段落数を確認 ---
assert.strictEqual(
    runner.run('AddParagraphs', []),
    3,
    '3行の内容 → Paragraphs.Count = 3'
);
console.log('[PASS] setContent("Line1\\nLine2\\nLine3") → Paragraphs.Count = 3');

// --- Selection.TypeText でテキスト入力 ---
assert.strictEqual(
    runner.run('TypeIntoDoc', []),
    'Hello, World',
    'Selection.TypeText の連結'
);
console.log('[PASS] Selection.TypeText "Hello" + ", World"');

console.log('\n✅ excel-to-word: 全テスト通過');
