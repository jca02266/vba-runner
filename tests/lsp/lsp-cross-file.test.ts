import { LSPServer } from '../../src/lsp/server';
import { assert } from '../../test-libs/test-runner';

const URI_A = 'file:///moduleA.bas';
const URI_B = 'file:///moduleB.bas';

function makeServer(docs: { uri: string; content: string }[]): LSPServer {
    const server = new LSPServer();
    for (const { uri, content } of docs) {
        server.didOpen(uri, content);
    }
    return server;
}

/** ワークスペーススキャン相当: loadWorkspaceFile でディスクファイルを登録し、1ファイルだけ didOpen する */
function makeServerWithWorkspace(
    workspaceFiles: { uri: string; content: string }[],
    openUri: string,
): LSPServer {
    const server = new LSPServer();
    for (const { uri, content } of workspaceFiles) {
        server.loadWorkspaceFile(uri, content);
    }
    const openFile = workspaceFiles.find(f => f.uri === openUri)!;
    server.didOpen(openFile.uri, openFile.content);
    return server;
}

function lineOf(src: string, needle: string): number {
    const lines = src.split('\n');
    for (let i = 0; i < lines.length; i++) {
        if (lines[i].includes(needle)) return i;
    }
    throw new Error(`"${needle}" not found`);
}

function colOf(src: string, needle: string, lineIdx: number): number {
    return src.split('\n')[lineIdx].indexOf(needle);
}

// ─── getDefinition: 別ファイルへのジャンプ ────────────────────────────────────

// 1. ファイル B に定義された Public Sub へ A からジャンプ
{
    const srcA = [
        'Sub Main()',
        '    CalcTotal',
        'End Sub',
    ].join('\n');
    const srcB = [
        'Sub CalcTotal()',
        'End Sub',
    ].join('\n');

    const server = makeServer([
        { uri: URI_A, content: srcA },
        { uri: URI_B, content: srcB },
    ]);

    const line = lineOf(srcA, 'CalcTotal');
    const col = colOf(srcA, 'CalcTotal', line);
    const def = server.getDefinition(URI_A, line, col);

    assert.ok(def, '別ファイルの Sub 定義が見つかること');
    assert.strictEqual(def.uri, URI_B, 'URI がファイル B を指すこと');
    assert.strictEqual(def.range.start.line, 0, 'ファイル B の行 0 に定義');
    console.log('[PASS] 別ファイルの Sub 定義ジャンプ');
}

// 2. ファイル B に定義された Function へ A からジャンプ
{
    const srcA = [
        'Sub Main()',
        '    Dim v As Long',
        '    v = CalcValue()',
        'End Sub',
    ].join('\n');
    const srcB = [
        'Function CalcValue() As Long',
        '    CalcValue = 42',
        'End Function',
    ].join('\n');

    const server = makeServer([
        { uri: URI_A, content: srcA },
        { uri: URI_B, content: srcB },
    ]);

    const line = lineOf(srcA, 'CalcValue');
    const col = colOf(srcA, 'CalcValue', line);
    const def = server.getDefinition(URI_A, line, col);

    assert.ok(def, '別ファイルの Function 定義が見つかること');
    assert.strictEqual(def.uri, URI_B, 'URI がファイル B を指すこと');
    console.log('[PASS] 別ファイルの Function 定義ジャンプ');
}

// 3. 現在ファイル内に定義がある場合は別ファイルより優先
{
    const srcA = [
        'Sub CalcTotal()',
        'End Sub',
        'Sub Main()',
        '    CalcTotal',
        'End Sub',
    ].join('\n');
    const srcB = [
        'Sub CalcTotal()',
        'End Sub',
    ].join('\n');

    const server = makeServer([
        { uri: URI_A, content: srcA },
        { uri: URI_B, content: srcB },
    ]);

    const line = lineOf(srcA, 'CalcTotal');
    const col = colOf(srcA, 'CalcTotal', line);
    const def = server.getDefinition(URI_A, line, col);

    assert.ok(def, '定義が見つかること');
    assert.strictEqual(def.uri, URI_A, '現在ファイルの定義が優先されること');
    console.log('[PASS] 現在ファイル内定義が別ファイルより優先される');
}

// ─── getReferences: 全ファイル横断検索 ───────────────────────────────────────

// 4. ファイル A で宣言した Sub の参照がファイル B でも見つかる
{
    const srcA = [
        'Sub CalcTotal()',
        'End Sub',
    ].join('\n');
    const srcB = [
        'Sub Main()',
        '    CalcTotal',
        '    CalcTotal',
        'End Sub',
    ].join('\n');

    const server = makeServer([
        { uri: URI_A, content: srcA },
        { uri: URI_B, content: srcB },
    ]);

    const line = lineOf(srcA, 'CalcTotal');
    const col = colOf(srcA, 'CalcTotal', line);
    const refs = server.getReferences(URI_A, line, col, true);

    const uris = refs.map((r: any) => r.uri);
    assert.ok(uris.includes(URI_A), 'ファイル A の宣言が含まれること');
    assert.ok(uris.includes(URI_B), 'ファイル B の呼び出しが含まれること');
    assert.strictEqual(refs.length, 3, '宣言 1件 + 呼び出し 2件 = 3件');
    console.log('[PASS] 別ファイルをまたぐ Sub 参照検索 (3件)');
}

// 5. ローカル変数はクロスファイル検索されない
{
    const srcA = [
        'Sub Foo()',
        '    Dim total As Long',
        '    total = 10',
        'End Sub',
    ].join('\n');
    const srcB = [
        'Sub Bar()',
        '    Dim total As Long',
        '    total = 99',
        'End Sub',
    ].join('\n');

    const server = makeServer([
        { uri: URI_A, content: srcA },
        { uri: URI_B, content: srcB },
    ]);

    // ファイル A の Foo 内の total を検索
    const line = lineOf(srcA, 'total = 10');
    const col = colOf(srcA, 'total', line);
    const refs = server.getReferences(URI_A, line, col, true);

    const urisInResult = refs.map((r: any) => r.uri);
    assert.ok(!urisInResult.includes(URI_B), 'ローカル変数はファイル B を検索しないこと');
    console.log('[PASS] ローカル変数はクロスファイル検索しない');
}

// 6. includeDeclaration=false で宣言除外のクロスファイル参照
{
    const srcA = [
        'Sub CalcTotal()',
        'End Sub',
    ].join('\n');
    const srcB = [
        'Sub Main()',
        '    CalcTotal',
        'End Sub',
    ].join('\n');

    const server = makeServer([
        { uri: URI_A, content: srcA },
        { uri: URI_B, content: srcB },
    ]);

    const line = lineOf(srcA, 'CalcTotal');
    const col = colOf(srcA, 'CalcTotal', line);
    const refs = server.getReferences(URI_A, line, col, false);

    assert.strictEqual(refs.length, 1, '宣言除外: ファイル B の呼び出しのみ 1件');
    assert.strictEqual(refs[0].uri, URI_B, 'ファイル B の参照のみ');
    console.log('[PASS] includeDeclaration=false のクロスファイル参照');
}

// ─── loadWorkspaceFile: 起動時スキャン相当 ───────────────────────────────────

// 7. エディターで開いていない（loadWorkspaceFile だけの）ファイルへ定義ジャンプ
{
    const srcA = [
        'Sub Main()',
        '    CalcTotal',
        'End Sub',
    ].join('\n');
    const srcB = [
        'Sub CalcTotal()',
        'End Sub',
    ].join('\n');

    // B はディスクファイルとして登録のみ（didOpen なし）
    const server = makeServerWithWorkspace(
        [{ uri: URI_A, content: srcA }, { uri: URI_B, content: srcB }],
        URI_A,
    );

    const line = lineOf(srcA, 'CalcTotal');
    const col = colOf(srcA, 'CalcTotal', line);
    const def = server.getDefinition(URI_A, line, col);

    assert.ok(def, 'ディスクのみのファイルへも定義ジャンプできること');
    assert.strictEqual(def.uri, URI_B, 'URI がファイル B を指すこと');
    console.log('[PASS] エディター未オープンのファイルへ定義ジャンプ（ワークスペーススキャン経由）');
}

// 8. エディターで開いていない（loadWorkspaceFile だけの）ファイルの参照も検索される
{
    const srcA = [
        'Sub CalcTotal()',
        'End Sub',
    ].join('\n');
    const srcB = [
        'Sub Main()',
        '    CalcTotal',
        '    CalcTotal',
        'End Sub',
    ].join('\n');

    // B はディスクファイルとして登録のみ
    const server = makeServerWithWorkspace(
        [{ uri: URI_A, content: srcA }, { uri: URI_B, content: srcB }],
        URI_A,
    );

    const line = lineOf(srcA, 'CalcTotal');
    const col = colOf(srcA, 'CalcTotal', line);
    const refs = server.getReferences(URI_A, line, col, true);

    assert.strictEqual(refs.length, 3, 'ディスクファイルも含め 3件');
    const uris = refs.map((r: any) => r.uri);
    assert.ok(uris.includes(URI_B), 'ディスクのみのファイル B の参照が含まれること');
    console.log('[PASS] エディター未オープンのファイルも参照検索（ワークスペーススキャン経由）');
}

// 9. didChange でワークスペースキャッシュが更新される
{
    const srcA_old = ['Sub Main()', '    OldSub', 'End Sub'].join('\n');
    const srcA_new = ['Sub Main()', '    NewSub', 'End Sub'].join('\n');
    const srcB = ['Sub NewSub()', 'End Sub'].join('\n');

    const server = makeServerWithWorkspace(
        [{ uri: URI_A, content: srcA_old }, { uri: URI_B, content: srcB }],
        URI_A,
    );
    // エディターで編集（保存前）
    server.didChange(URI_A, srcA_new);

    const line = lineOf(srcA_new, 'NewSub');
    const col = colOf(srcA_new, 'NewSub', line);
    const refs = server.getReferences(URI_A, line, col, true);

    assert.strictEqual(refs.length, 2, '編集後の内容が反映されて 2件');
    console.log('[PASS] didChange でワークスペースキャッシュが更新される');
}

// 10. unloadWorkspaceFile で削除後は検索されない
{
    const srcA = ['Sub Main()', '    CalcTotal', 'End Sub'].join('\n');
    const srcB = ['Sub CalcTotal()', 'End Sub'].join('\n');

    const server = makeServerWithWorkspace(
        [{ uri: URI_A, content: srcA }, { uri: URI_B, content: srcB }],
        URI_A,
    );
    server.unloadWorkspaceFile(URI_B);

    const line = lineOf(srcA, 'CalcTotal');
    const col = colOf(srcA, 'CalcTotal', line);
    const def = server.getDefinition(URI_A, line, col);

    assert.strictEqual(def, null, '削除後はファイル B の定義が見つからないこと');
    console.log('[PASS] unloadWorkspaceFile 後は検索されない');
}

// ─── 遅延ディレクトリスキャン（実ディスク）────────────────────────────────

import { writeFileSync, unlinkSync, mkdirSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

{
    // 11. 同ディレクトリの .bas ファイルへ遅延ジャンプできる
    const dir = join(tmpdir(), `vba-test-${Date.now()}`);
    mkdirSync(dir);

    const mainPath = join(dir, 'Main.bas');
    const helperPath = join(dir, 'Helper.bas');

    const mainSrc = ['Sub Main()', '    CalcHelper', 'End Sub'].join('\n');
    const helperSrc = ['Sub CalcHelper()', 'End Sub'].join('\n');

    writeFileSync(mainPath, mainSrc, 'utf-8');
    writeFileSync(helperPath, helperSrc, 'utf-8');

    const mainUri = `file://${mainPath.replace(/\\/g, '/')}`;

    const server = new LSPServer();
    server.didOpen(mainUri, mainSrc); // Helper.bas は didOpen しない

    const line = lineOf(mainSrc, 'CalcHelper');
    const col = colOf(mainSrc, 'CalcHelper', line);
    const def = server.getDefinition(mainUri, line, col);

    rmSync(dir, { recursive: true });

    assert.ok(def, '遅延スキャンで同ディレクトリの Sub 定義が見つかること');
    assert.ok(def.uri.includes('Helper.bas'), '定義の URI が Helper.bas を指すこと');
    console.log('[PASS] 遅延ディレクトリスキャンで別ファイルの Sub 定義ジャンプ');
}

{
    // 12. 同ディレクトリの .bas ファイルの参照も遅延スキャンで拾える
    const dir = join(tmpdir(), `vba-test-${Date.now()}`);
    mkdirSync(dir);

    const mainPath = join(dir, 'Main.bas');
    const helperPath = join(dir, 'Helper.bas');

    const helperSrc = ['Sub CalcHelper()', 'End Sub'].join('\n');
    const mainSrc = ['Sub Main()', '    CalcHelper', '    CalcHelper', 'End Sub'].join('\n');

    writeFileSync(helperPath, helperSrc, 'utf-8');
    writeFileSync(mainPath, mainSrc, 'utf-8');

    const helperUri = `file://${helperPath.replace(/\\/g, '/')}`;

    const server = new LSPServer();
    server.didOpen(helperUri, helperSrc); // Main.bas は didOpen しない

    const line = lineOf(helperSrc, 'CalcHelper');
    const col = colOf(helperSrc, 'CalcHelper', line);
    const refs = server.getReferences(helperUri, line, col, true);

    rmSync(dir, { recursive: true });

    assert.strictEqual(refs.length, 3, '遅延スキャンで 宣言+呼び出し2件 = 3件');
    console.log('[PASS] 遅延ディレクトリスキャンで別ファイルの参照検索');
}

// 13. 別ファイルが同名シンボルを独自宣言している場合、そのファイルは参照候補に含めない
//     (Private/Public 問わず、別ファイルの同名 Dim は別定義)
{
    const srcA = [
        'Private m_cells As Object',
        'Sub UseIt()',
        '    m_cells.Add 1',
        'End Sub',
    ].join('\n');
    const srcB = [
        'Dim m_cells As Object',   // 全く別の変数
        'Sub Bar()',
        '    m_cells.Clear',
        'End Sub',
    ].join('\n');

    const server = makeServer([
        { uri: URI_A, content: srcA },
        { uri: URI_B, content: srcB },
    ]);

    const line = lineOf(srcA, 'm_cells As Object');
    const col = colOf(srcA, 'm_cells', line);
    const refs = server.getReferences(URI_A, line, col, true);

    const uris = refs.map((r: any) => r.uri);
    assert.ok(!uris.includes(URI_B), '別ファイルが同名を独自宣言している場合、そのファイルの参照を含めない');
    assert.ok(uris.includes(URI_A), '現在ファイルの参照は含まれる');
    console.log('[PASS] 同名独自宣言ファイルはクロスファイル参照から除外される');
}

// 14. Public シンボルでも別ファイルが同名を独自宣言していれば除外
{
    const srcA = [
        'Public SharedName As Long',
    ].join('\n');
    const srcB = [
        'Public SharedName As String',  // 別の定義
        'Sub Use()',
        '    SharedName = "x"',
        'End Sub',
    ].join('\n');
    const srcC = [
        'Sub CallIt()',
        '    SharedName = 42',          // A の SharedName を参照
        'End Sub',
    ].join('\n');

    const URI_C = 'file:///C.bas';
    const server = makeServer([
        { uri: URI_A, content: srcA },
        { uri: URI_B, content: srcB },
        { uri: URI_C, content: srcC },
    ]);

    const line = lineOf(srcA, 'SharedName');
    const col = colOf(srcA, 'SharedName', line);
    const refs = server.getReferences(URI_A, line, col, true);

    const uris = refs.map((r: any) => r.uri);
    assert.ok(!uris.includes(URI_B), 'B は独自に SharedName を宣言 → 除外');
    assert.ok(uris.includes(URI_C), 'C は独自宣言なし → A の SharedName を参照 → 含まれる');
    console.log('[PASS] Public シンボルでも同名独自宣言ファイルを除外し、非宣言ファイルは含める');
}

// 15. 別ファイルのプロシージャ内ローカル変数も除外される
{
    const srcA = [
        'Public Total As Long',
        'Sub SetTotal()',
        '    Total = 100',
        'End Sub',
    ].join('\n');
    const srcB = [
        'Sub Calc()',
        '    Dim Total As Long',   // ローカル変数（別物）
        '    Total = 50',
        'End Sub',
    ].join('\n');
    const srcC = [
        'Sub Print()',
        '    Debug.Print Total',   // A の Total を参照
        'End Sub',
    ].join('\n');

    const URI_C = 'file:///C.bas';
    const server = makeServer([
        { uri: URI_A, content: srcA },
        { uri: URI_B, content: srcB },
        { uri: URI_C, content: srcC },
    ]);

    const line = lineOf(srcA, 'Total As Long');
    const col = colOf(srcA, 'Total', line);
    const refs = server.getReferences(URI_A, line, col, true);

    const uris = refs.map((r: any) => r.uri);
    assert.ok(!uris.includes(URI_B), 'B はローカル変数 Total を持つ → 除外');
    assert.ok(uris.includes(URI_C), 'C は独自宣言なし → A の Total を参照 → 含まれる');
    console.log('[PASS] プロシージャ内ローカル変数を持つファイルもクロスファイル参照から除外');
}

// 16. 予約語（Set/And/Not 等）はシンボルテーブルに存在しないため参照検索で空を返す
{
    const src = [
        'Sub Foo()',
        '    Dim obj As Object',
        '    Set obj = New Object',
        'End Sub',
    ].join('\n');

    const server = makeServer([{ uri: URI_A, content: src }]);

    // "Set" キーワードの位置
    const line = lineOf(src, 'Set obj');
    const col = colOf(src, 'Set', line);
    const refs = server.getReferences(URI_A, line, col, true);

    assert.strictEqual(refs.length, 0, '予約語 Set は参照検索で空を返す');
    console.log('[PASS] 予約語（Set）は参照検索で空を返す');
}

// 17. 同名の識別子（ユーザー定義 Sub）がある場合は正常に参照検索できる
{
    const srcA = [
        'Sub SetValue()',
        'End Sub',
    ].join('\n');
    const srcB = [
        'Sub Main()',
        '    SetValue',
        'End Sub',
    ].join('\n');

    const server = makeServer([
        { uri: URI_A, content: srcA },
        { uri: URI_B, content: srcB },
    ]);

    const line = lineOf(srcA, 'SetValue');
    const col = colOf(srcA, 'SetValue', line);
    const refs = server.getReferences(URI_A, line, col, true);

    assert.strictEqual(refs.length, 2, 'ユーザー定義 SetValue は宣言+呼び出しで2件');
    console.log('[PASS] ユーザー定義識別子は正常に参照検索できる');
}

// 18. Private Sub はクロスファイル検索されない
{
    const srcA = [
        'Private Sub PrivFoo()',
        'End Sub',
        'Sub CallIt()',
        '    PrivFoo',
        'End Sub',
    ].join('\n');
    const srcB = [
        'Sub Bar()',
        '    PrivFoo',   // 別の文脈（モジュール外から呼べないはずだが名前一致）
        'End Sub',
    ].join('\n');

    const server = makeServer([
        { uri: URI_A, content: srcA },
        { uri: URI_B, content: srcB },
    ]);

    const line = lineOf(srcA, 'PrivFoo()');
    const col = colOf(srcA, 'PrivFoo', line);
    const refs = server.getReferences(URI_A, line, col, true);

    const uris = refs.map((r: any) => r.uri);
    assert.ok(!uris.includes(URI_B), 'Private Sub はクロスファイル検索しない');
    assert.ok(uris.includes(URI_A), '現在ファイル内の参照は含まれる');
    console.log('[PASS] Private Sub はクロスファイル参照検索されない');
}

// 19. モジュールレベルの Dim（スコープなし）でも別ファイルの同名 Dim は除外される
{
    const srcA = [
        'Dim counter As Long',
        'Sub Inc()',
        '    counter = counter + 1',
        'End Sub',
    ].join('\n');
    const srcB = [
        'Dim counter As Long',   // 全く別の変数
        'Sub Reset()',
        '    counter = 0',
        'End Sub',
    ].join('\n');

    const server = makeServer([
        { uri: URI_A, content: srcA },
        { uri: URI_B, content: srcB },
    ]);

    const line = lineOf(srcA, 'counter As Long');
    const col = colOf(srcA, 'counter', line);
    const refs = server.getReferences(URI_A, line, col, true);

    const uris = refs.map((r: any) => r.uri);
    assert.ok(!uris.includes(URI_B), 'スコープなし Dim でも別ファイルの同名宣言は除外');
    assert.ok(uris.includes(URI_A), '現在ファイルの参照は含まれる');
    console.log('[PASS] モジュールレベル Dim（スコープなし）でも別ファイルの同名宣言を除外');
}

console.log('\n✅ LSP Cross-file Definition / References: 全テスト通過');
