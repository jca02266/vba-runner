import { Lexer } from '../../../src/engine/lexer';
import { Parser } from '../../../src/engine/parser';
import { Evaluator } from '../../../src/engine/evaluator';
import { assert } from '../../../test-libs/test-runner';
import * as fs from 'fs';
import * as path from 'path';

const libCode = fs.readFileSync(
    path.join(__dirname, '../../src/vba/LibPath.bas'),
    'utf-8'
);

const ev = (() => {
    const tokens = new Lexer(libCode).tokenize();
    const ast = new Parser(tokens).parse();
    const ev = new Evaluator(console.log);
    ev.evaluate(ast);
    return ev;
})();

function r(name: string, ...args: any[]): any {
    return ev.callProcedure(name, args);
}

// --- NormalizeSeparators ---
{
    assert.strictEqual(r('NormalizeSeparators', 'C:/foo/bar'), 'C:\\foo\\bar', '/ → \\');
    assert.strictEqual(r('NormalizeSeparators', 'C:\\foo\\bar'), 'C:\\foo\\bar', '\\ はそのまま');
    assert.strictEqual(r('NormalizeSeparators', 'foo/bar/baz'), 'foo\\bar\\baz', '相対パス');
    console.log('[PASS] NormalizeSeparators');
}

// --- GetFileName ---
{
    assert.strictEqual(r('GetFileName', 'C:\\foo\\bar\\baz.txt'), 'baz.txt', 'フルパスからファイル名');
    assert.strictEqual(r('GetFileName', 'baz.txt'),               'baz.txt', '区切りなし');
    assert.strictEqual(r('GetFileName', 'C:\\foo\\'),             '',        '末尾が区切り文字');
    assert.strictEqual(r('GetFileName', 'C:/foo/bar.xlsx'),       'bar.xlsx', '/ 区切りも対応');
    console.log('[PASS] GetFileName');
}

// --- GetExtension ---
{
    assert.strictEqual(r('GetExtension', 'C:\\foo\\bar.txt'),  '.txt',   '拡張子あり');
    assert.strictEqual(r('GetExtension', 'C:\\foo\\bar.xlsx'), '.xlsx',  '複数文字拡張子');
    assert.strictEqual(r('GetExtension', 'C:\\foo\\bar'),      '',       '拡張子なし');
    assert.strictEqual(r('GetExtension', '.gitignore'),        '.gitignore', 'ドット始まりは全体が拡張子');
    console.log('[PASS] GetExtension');
}

// --- GetBaseName ---
{
    assert.strictEqual(r('GetBaseName', 'C:\\foo\\bar.txt'),  'bar',  '拡張子を除いたファイル名');
    assert.strictEqual(r('GetBaseName', 'C:\\foo\\bar'),      'bar',  '拡張子なし');
    assert.strictEqual(r('GetBaseName', 'report.xlsx'),       'report', '相対パス');
    console.log('[PASS] GetBaseName');
}

// --- GetDirectory ---
{
    assert.strictEqual(r('GetDirectory', 'C:\\foo\\bar\\baz.txt'), 'C:\\foo\\bar', '2階層');
    assert.strictEqual(r('GetDirectory', 'C:\\foo\\bar'),          'C:\\foo',      '1階層');
    assert.strictEqual(r('GetDirectory', 'C:\\foo'),               'C:\\',         'ドライブルート直下');
    assert.strictEqual(r('GetDirectory', 'foo\\bar'),              'foo',          '相対パス');
    assert.strictEqual(r('GetDirectory', 'file.txt'),              '',             '区切りなし');
    console.log('[PASS] GetDirectory');
}

// --- GetDriveLetter ---
{
    assert.strictEqual(r('GetDriveLetter', 'C:\\foo\\bar'), 'C', 'C ドライブ');
    assert.strictEqual(r('GetDriveLetter', 'D:\\'),         'D', 'D ドライブ');
    assert.strictEqual(r('GetDriveLetter', '\\foo\\bar'),   '',  'ドライブなし');
    assert.strictEqual(r('GetDriveLetter', 'foo\\bar'),     '',  '相対パス');
    console.log('[PASS] GetDriveLetter');
}

// --- IsAbsolutePath ---
{
    assert.strictEqual(r('IsAbsolutePath', 'C:\\foo\\bar'),   -1, 'ドライブ付き絶対パス');
    assert.strictEqual(r('IsAbsolutePath', '\\\\server\\s'),  -1, 'UNC パス');
    assert.strictEqual(r('IsAbsolutePath', 'foo\\bar'),        0, '相対パス');
    assert.strictEqual(r('IsAbsolutePath', 'bar.txt'),         0, 'ファイル名のみ');
    console.log('[PASS] IsAbsolutePath');
}

// --- IsUNCPath ---
{
    assert.strictEqual(r('IsUNCPath', '\\\\server\\share'), -1, 'UNC パス');
    assert.strictEqual(r('IsUNCPath', '//server/share'),    -1, '/ 区切り UNC');
    assert.strictEqual(r('IsUNCPath', 'C:\\foo'),            0, '通常の絶対パス');
    assert.strictEqual(r('IsUNCPath', 'foo\\bar'),           0, '相対パス');
    console.log('[PASS] IsUNCPath');
}

// --- PathJoin ---
{
    assert.strictEqual(r('PathJoin', 'C:\\foo',  'bar.txt'),   'C:\\foo\\bar.txt',  '基本結合');
    assert.strictEqual(r('PathJoin', 'C:\\foo\\', 'bar.txt'),  'C:\\foo\\bar.txt',  '末尾 \\ 付き dir');
    assert.strictEqual(r('PathJoin', 'C:\\foo',  '\\bar.txt'), 'C:\\foo\\bar.txt',  '先頭 \\ 付き file');
    assert.strictEqual(r('PathJoin', '',          'bar.txt'),  'bar.txt',            '空 dir');
    assert.strictEqual(r('PathJoin', 'C:\\foo',  ''),          'C:\\foo',            '空 file');
    assert.strictEqual(r('PathJoin', 'C:/foo',   'bar/baz'),   'C:\\foo\\bar\\baz', '/ 区切り正規化');
    console.log('[PASS] PathJoin');
}

// --- EnsureTrailingSlash / RemoveTrailingSlash ---
{
    assert.strictEqual(r('EnsureTrailingSlash', 'C:\\foo'),    'C:\\foo\\', '追加');
    assert.strictEqual(r('EnsureTrailingSlash', 'C:\\foo\\'),  'C:\\foo\\', 'すでにあり');
    assert.strictEqual(r('EnsureTrailingSlash', ''),           '',          '空文字はそのまま');
    assert.strictEqual(r('RemoveTrailingSlash', 'C:\\foo\\'),  'C:\\foo',   '除去');
    assert.strictEqual(r('RemoveTrailingSlash', 'C:\\foo'),    'C:\\foo',   'なければそのまま');
    assert.strictEqual(r('RemoveTrailingSlash', 'C:\\'),       'C:\\',      'ドライブルートは除去しない');
    console.log('[PASS] EnsureTrailingSlash / RemoveTrailingSlash');
}

// --- ChangeExtension ---
{
    assert.strictEqual(r('ChangeExtension', 'C:\\foo\\bar.txt',  '.csv'),  'C:\\foo\\bar.csv',  '拡張子変更');
    assert.strictEqual(r('ChangeExtension', 'C:\\foo\\bar.xlsx', '.txt'),  'C:\\foo\\bar.txt',  'xlsx→txt');
    assert.strictEqual(r('ChangeExtension', 'report',             '.xlsx'), 'report.xlsx',       'ディレクトリなし');
    assert.strictEqual(r('ChangeExtension', 'C:\\foo\\bar',       '.txt'), 'C:\\foo\\bar.txt',  '拡張子なし→追加');
    console.log('[PASS] ChangeExtension');
}

// --- PathDepth ---
{
    assert.strictEqual(r('PathDepth', 'C:\\foo\\bar\\baz'), 3, '3階層');
    assert.strictEqual(r('PathDepth', 'C:\\foo\\bar'),      2, '2階層');
    assert.strictEqual(r('PathDepth', 'C:\\foo'),           1, '1階層');
    assert.strictEqual(r('PathDepth', 'C:\\'),              0, 'ドライブルート=0');
    assert.strictEqual(r('PathDepth', 'foo\\bar'),          1, '相対パス');
    assert.strictEqual(r('PathDepth', '\\\\server\\share'), 1, 'UNC share');
    console.log('[PASS] PathDepth');
}

console.log('\n✅ LibPath: 全テスト通過');
