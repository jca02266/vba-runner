/**
 * vba-extractor `import` の出力構造の回帰テスト。
 *
 * sample/excel/test.xlsm をエクスポート → 再インポートし、生成された
 * vbaProject.bin が「ソースのみ（デコンパイル状態）」の正しい形になっていることを検証する。
 * 詳細は docs/internals/VBA_EXPORT_IMPORT.md を参照。
 *
 * 特に重要なのは「CFB ディレクトリの赤黒木に到達可能な type=0（無効）ノードが無い」こと。
 * 過去に Sh33tJ5 署名ストリームを type=0 へ書き換えてダングリングポインターを残し、
 * Excel の「シートをコピーできません」破損を起こしたことがある（その回帰を防ぐ）。
 */
import { execFileSync } from 'child_process';
import { mkdtempSync, readFileSync, renameSync, rmSync, writeFileSync } from 'fs';
import { tmpdir } from 'os';
import { join } from 'path';
import CFB from 'cfb';
import JSZip from 'jszip';
import iconv from 'iconv-lite';
import { assert } from '../../test-libs/test-runner';
import { decompress } from '../../tools/extractor/ovba';
import { parseDirStreamFull } from '../../tools/extractor/dir-parser';

const XLSM = 'sample/excel/test.xlsm';
const EMPTY_XLSM = 'sample/excel/empty_with_macro.xlsm';
const NO_VBA_XLSM = 'sample/excel/empty.xlsm';

const tmp = mkdtempSync(join(tmpdir(), 'vba-extractor-test-'));
const CLI = join(tmp, 'vba-extractor.cjs');
const srcDir = join(tmp, 'src');
const outXlsm = join(tmp, 'out.xlsm');

try {
    // Bundle once, then run with node. This keeps the multi-invocation CLI test
    // within the all-tests runner's per-test timeout.
    execFileSync('./node_modules/.bin/esbuild', [
        'tools/extractor/vba-extractor.ts', '--bundle', `--outfile=${CLI}`, '--platform=node',
        '--define:import.meta.dirname=__dirname',
    ], { stdio: 'pipe' });

    // 1. export → 2. import（確認プロンプトに 'y' を渡す）
    execFileSync('node', [CLI, 'export', XLSM, srcDir], { stdio: 'pipe' });
    execFileSync('node', [CLI, 'import', XLSM, srcDir, outXlsm], { input: 'y\n', stdio: 'pipe' });

    // 生成された vbaProject.bin を解析
    const bin = await JSZip.loadAsync(readFileSync(outXlsm))
        .then(z => z.file('xl/vbaProject.bin')!.async('nodebuffer'));
    const cfb = CFB.read(bin, { type: 'buffer' });
    // L/R/C are red-black tree pointers present at runtime but absent from cfb.js types.
    const fi = cfb.FileIndex as unknown as { type: number; name: string; L: number; R: number; C: number }[];

    // --- 全モジュールの MODULEOFFSET が 0（プリアンブルなし）かつソースが展開できる ---
    const dir = parseDirStreamFull(
        decompress(Buffer.from(CFB.find(cfb, '/VBA/dir')!.content as unknown as ArrayBuffer)),
    );
    assert.ok(dir.modules.length > 0, 'モジュールが 1 つ以上ある');
    for (const m of dir.modules) {
        assert.strictEqual(m.offset, 0, `${m.name} の MODULEOFFSET=0`);
        const entry = CFB.find(cfb, `/VBA/${m.streamName}`)!;
        const stream = Buffer.from(entry.content as unknown as ArrayBuffer);
        // 例外を投げずに展開でき、モジュール名属性を含むこと
        const src = iconv.decode(decompress(stream.subarray(m.offset)), 'cp932');
        assert.ok(/Attribute\s+VB_Name/i.test(src), `${m.name} のソースが展開できる`);
    }
    console.log('[PASS] 全モジュール MODULEOFFSET=0 / ソース展開');

    // --- _VBA_PROJECT が 7 バイトの正規ヘッダー（Version=0xFFFF） ---
    const vp = Buffer.from(CFB.find(cfb, '/VBA/_VBA_PROJECT')!.content as unknown as ArrayBuffer);
    assert.strictEqual(vp.toString('hex'), 'cc61ffff000000', '_VBA_PROJECT = CC61FFFF000000');
    console.log('[PASS] _VBA_PROJECT 7バイトヘッダー');

    // --- __SRP_* が有効ストリームとして残っていない ---
    const activeSrp = fi.filter(e => e.type === 2 && String(e.name).includes('SRP_')).length;
    assert.strictEqual(activeSrp, 0, '__SRP_* は全削除');
    console.log('[PASS] __SRP_* 全削除');

    // --- Sh33tJ5 署名は有効ストリーム（type=2）として残す（消そうとしない） ---
    const sig = fi.find(e => String(e.name).includes('Sh33tJ5'));
    assert.ok(sig, 'Sh33tJ5 署名ストリームが存在する');
    assert.strictEqual(sig.type, 2, 'Sh33tJ5 は有効ストリーム(type=2)');
    console.log('[PASS] Sh33tJ5 は有効ストリームのまま');

    // --- CFB の木に到達可能な type=0(無効)ノードが無い（破損の回帰ガード） ---
    const reachable = new Set<number>();
    (function walk(i: number) {
        if (i < 0 || i == null || reachable.has(i)) return;
        reachable.add(i);
        const e = fi[i];
        if (!e) return;
        walk(e.L); walk(e.R); walk(e.C);
    })(0);
    const reachableInvalid = [...reachable].filter(i => fi[i] && fi[i].type === 0);
    assert.strictEqual(reachableInvalid.length, 0, '木に到達可能な type=0 ノードが無い');
    console.log('[PASS] 赤黒木に無効ノードなし（Sh33tJ5 破損の回帰ガード）');

    // --- ホストの Document モジュールをソース省略で削除しない ---
    // ThisWorkbook/Sheet1 は PROJECT の Document= 宣言と OOXML ホストに
    // 結び付くため、ソースディレクトリに無くても空のコードストリームを保持する。
    const hostSrcDir = join(tmp, 'host-src');
    const hostOut = join(tmp, 'host-omitted.xlsm');
    execFileSync('node', [CLI, 'export', EMPTY_XLSM, hostSrcDir], { stdio: 'pipe' });
    renameSync(join(hostSrcDir, 'ThisWorkbook.cls'), join(tmp, 'ThisWorkbook.omitted.cls'));
    renameSync(join(hostSrcDir, 'Sheet1.cls'), join(tmp, 'Sheet1.omitted.cls'));
    execFileSync('node', [CLI, 'import', EMPTY_XLSM, hostSrcDir, hostOut, '--yes'], { stdio: 'pipe' });
    const hostBin = await JSZip.loadAsync(readFileSync(hostOut))
        .then(z => z.file('xl/vbaProject.bin')!.async('nodebuffer'));
    const hostCfb = CFB.read(hostBin, { type: 'buffer' });
    assert.ok(CFB.find(hostCfb, '/VBA/ThisWorkbook'), 'ThisWorkbook ストリームを保持する');
    assert.ok(CFB.find(hostCfb, '/VBA/Sheet1'), 'Sheet1 ストリームを保持する');
    const hostProject = iconv.decode(
        Buffer.from(CFB.find(hostCfb, '/PROJECT')!.content as unknown as ArrayBuffer), 'cp932',
    );
    assert.ok(/^Document=ThisWorkbook\//mi.test(hostProject), 'ThisWorkbook の Document 宣言を保持する');
    assert.ok(/^Document=Sheet1\//mi.test(hostProject), 'Sheet1 の Document 宣言を保持する');
    const hostDir = parseDirStreamFull(
        decompress(Buffer.from(CFB.find(hostCfb, '/VBA/dir')!.content as unknown as ArrayBuffer)),
    );
    for (const name of ['ThisWorkbook', 'Sheet1']) {
        const mod = hostDir.modules.find(m => m.name === name)!;
        const stream = Buffer.from(CFB.find(hostCfb, `/VBA/${mod.streamName}`)!.content as unknown as ArrayBuffer);
        const source = iconv.decode(decompress(stream.subarray(mod.offset)), 'cp932');
        assert.ok(/Attribute\s+VB_Name/i.test(source), `${name} の空ソースが有効`);
    }
    console.log('[PASS] ホスト Document モジュールはソース省略時も保持');

    // --- VBAプロジェクトを持たないXLSMにも新規プロジェクトを生成する ---
    const generatedSrcDir = join(tmp, 'generated-src');
    const generatedOut = join(tmp, 'generated.xlsm');
    // CLIはソースディレクトリを走査するため、エクスポートでディレクトリを作成する。
    execFileSync('node', [CLI, 'export', EMPTY_XLSM, join(tmp, 'unused-src')], { stdio: 'pipe' });
    writeFileSync(join(tmp, 'unused-src', 'Module1.bas'), 'Attribute VB_Name = "Module1"\r\n');
    writeFileSync(join(tmp, 'unused-src', 'DataProcessor.bas'), 'Option Explicit\r\n');
    writeFileSync(join(tmp, 'unused-src', 'Inventory.cls'), 'VERSION 1.0 CLASS\r\nBEGIN\r\n  MultiUse = -1\r\nEND\r\nAttribute VB_Name = "Inventory"\r\n');
    writeFileSync(join(tmp, 'unused-src', 'Product.cls'), 'VERSION 1.0 CLASS\r\nBEGIN\r\n  MultiUse = -1\r\nEND\r\nAttribute VB_Name = "Product"\r\n');
    execFileSync('node', [CLI, 'import', NO_VBA_XLSM, join(tmp, 'unused-src'), generatedOut, '--yes'], { stdio: 'pipe' });
    const generatedZip = await JSZip.loadAsync(readFileSync(generatedOut));
    assert.ok(generatedZip.file('xl/vbaProject.bin'), 'VBAプロジェクトを新規生成する');
    assert.ok(/macroEnabled\.main\+xml/i.test(await generatedZip.file('[Content_Types].xml')!.async('string')),
        'Workbookをマクロ有効Content Typeにする');
    assert.ok(/relationships\/vbaProject/i.test(await generatedZip.file('xl/_rels/workbook.xml.rels')!.async('string')),
        'VBAプロジェクトのリレーションを追加する');
    const generatedCfb = CFB.read(await generatedZip.file('xl/vbaProject.bin')!.async('nodebuffer'), { type: 'buffer' });
    assert.ok(CFB.find(generatedCfb, '/VBA/ThisWorkbook'), '生成プロジェクトにThisWorkbookを含める');
    assert.ok(CFB.find(generatedCfb, '/VBA/Sheet1'), '生成プロジェクトにSheet1を含める');
    assert.ok(CFB.find(generatedCfb, '/VBA/Module1'), '生成プロジェクトにソースModule1を含める');
    const generatedProject = Buffer.from(CFB.find(generatedCfb, '/PROJECT')!.content as unknown as ArrayBuffer)
        .toString('latin1');
    assert.ok(/^Class=Inventory$/m.test(generatedProject), 'InventoryをClassとして登録する');
    assert.ok(/^Class=Product$/m.test(generatedProject), 'ProductをClassとして登録する');
    assert.ok(/^Module=DataProcessor$/m.test(generatedProject), 'DataProcessorの大文字小文字を保持する');
    assert.ok(!/^Module=dataprocessor$/m.test(generatedProject), '小文字化したモジュール名を生成しない');
    const generatedWorkbook = await generatedZip.file('xl/workbook.xml')!.async('string');
    const generatedSheet = await generatedZip.file('xl/worksheets/sheet1.xml')!.async('string');
    assert.ok(/<workbookPr\b[^>]*\bcodeName="ThisWorkbook"/i.test(generatedWorkbook),
        'Workbook DocumentのcodeNameを設定する');
    assert.ok(/<sheetPr\b[^>]*\bcodeName="Sheet1"/i.test(generatedSheet),
        'Sheet DocumentのcodeNameを設定する');
    console.log('[PASS] VBAプロジェクトなしXLSMへ最小プロジェクトを生成');

    // --- 新規クラスの MODULEPRIVATE (0x0028) ---
    // この空レコードを欠く dir stream は Excel で開けても VBE 保存時の検証に失敗する。
    writeFileSync(join(srcDir, 'NewClass.cls'), [
        'Attribute VB_Name = "NewClass"',
        'Attribute VB_PredeclaredId = False',
        'Option Explicit',
    ].join('\r\n'));
    const classOut = join(tmp, 'new-class.xlsm');
    execFileSync('node', [CLI, 'import', XLSM, srcDir, classOut], { input: 'y\n', stdio: 'pipe' });
    const classBin = await JSZip.loadAsync(readFileSync(classOut))
        .then(z => z.file('xl/vbaProject.bin')!.async('nodebuffer'));
    const classCfb = CFB.read(classBin, { type: 'buffer' });
    const classDir = parseDirStreamFull(
        decompress(Buffer.from(CFB.find(classCfb, '/VBA/dir')!.content as unknown as ArrayBuffer)),
    );
    const newClass = classDir.modules.find(m => m.name === 'NewClass')!;
    assert.ok(newClass, '新規クラスが dir stream に存在する');
    assert.ok(newClass.rawBlock.includes(Buffer.from([0x28, 0x00, 0x00, 0x00, 0x00, 0x00])),
        '新規クラスに MODULEPRIVATE (0x0028) がある');
    const newClassSource = iconv.decode(
        decompress(Buffer.from(CFB.find(classCfb, '/VBA/NewClass')!.content as unknown as ArrayBuffer)), 'cp932',
    );
    assert.ok(/Attribute VB_Base = "0\{FCFB3D2A-A0FA-1068-A738-08002B3371B5\}"/i.test(newClassSource),
        '新規クラスに Excel の VB_Base 属性を補う');
    assert.ok(/Attribute VB_TemplateDerived = False/i.test(newClassSource),
        '新規クラスに VB_TemplateDerived 属性を補う');
    assert.ok(/Attribute VB_Customizable = False/i.test(newClassSource),
        '新規クラスに VB_Customizable 属性を補う');
    console.log('[PASS] 新規クラス MODULEPRIVATE');


    console.log('\n✅ import-structure: 全テスト通過');
} finally {
    rmSync(tmp, { recursive: true, force: true });
}
