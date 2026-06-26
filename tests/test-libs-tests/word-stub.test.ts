/**
 * Word API スタブモード（injectWordStub）のテスト
 */
import { VBARunner, assert } from '../../test-libs/test-runner';
import { injectWordStub } from '../../test-libs/word-stub';

const runner = new VBARunner('tests/fixtures/word-stub');
const app = injectWordStub(runner.evaluator);

// --- ActiveDocument.Content.Text の読み書き ---
assert.strictEqual(runner.run('TestContent', []), 'Hello, World', 'Content.Text');
console.log('[PASS] ActiveDocument.Content.Text 読み書き');

// --- Selection.TypeText ---
assert.strictEqual(runner.run('TestTypeText', []), 'Hello World', 'Selection.TypeText');
console.log('[PASS] Selection.TypeText');

// --- Bookmarks ---
assert.strictEqual(runner.run('TestBookmark', []), 'Report 2024', 'Bookmark Range.Text');
console.log('[PASS] Bookmarks("bmkTitle").Range.Text');

// --- Paragraphs の For Each ---
app.doc.setContent('Line 1\nLine 2\nLine 3');
assert.strictEqual(runner.run('TestParagraphLoop', []), 3, 'Paragraphs ループ 3件');
console.log('[PASS] For Each p In ActiveDocument.Paragraphs');

// --- Tables(1).Cell(1,1) ---
app.doc.addTable(3, 3);
assert.strictEqual(runner.run('TestTableCell', []), 'Header', 'Table Cell Range.Text');
console.log('[PASS] Tables(1).Cell(1,1).Range.Text');

// --- Application.Name ---
assert.strictEqual(runner.run('TestAppName', []), 'Microsoft Word', 'Application.Name');
console.log('[PASS] Application.ScreenUpdating / Name');

// --- Find.Execute（スタブ: Found = False）---
assert.strictEqual(runner.run('TestFind', []), false, 'Find.Found = False');
console.log('[PASS] Selection.Find.Execute (スタブ: Found = False)');

// --- Documents.Open ---
assert.strictEqual(runner.run('TestDocumentsOpen', []), 'sales.docx', 'Documents.Open → Name');
console.log('[PASS] Documents.Open("C:\\reports\\sales.docx").Name');

console.log('\n✅ word-stub: 全テスト通過');
