/**
 * VBA Test Runner - .cls / .vba ファイルを自動実行
 * tests/spec/vba/ 直下のファイルと各サブディレクトリを別スイートとして実行する。
 * テストプロシージャは test で始まる名前（大文字小文字を区別しない）で検出される。
 */

import * as fs from 'fs';
import * as path from 'path';
import { VBARunner } from '../../../test-libs/test-runner';
import { Lexer } from '../../../src/compiler/lexer';
import { Parser } from '../../../src/compiler/parser';

const vbaDir = path.join(__dirname);
const VBA_EXTS = new Set(['.cls', '.vba']);

console.log('=== VBA Test Runner ===\n');

let totalFiles = 0;
let totalTests = 0;
let totalPass = 0;
const allFailed: string[] = [];

function collectProcedures(environment: any): Map<string, any> {
    const procs = new Map();
    if (environment && environment.procedures) {
        for (const [name, proc] of environment.procedures.entries()) {
            if (typeof name === 'string') procs.set(name, proc);
        }
    }
    if (environment && environment.enclosing) {
        const parent = collectProcedures(environment.enclosing);
        for (const [name, proc] of parent) {
            if (!procs.has(name)) procs.set(name, proc);
        }
    }
    return procs;
}

function injectFile(vbaTest: VBARunner, filePath: string): void {
    let source = fs.readFileSync(filePath, 'utf-8');
    const ext = path.extname(filePath).toLowerCase();
    const moduleName = path.basename(filePath, path.extname(filePath));
    const evaluator = (vbaTest as any).evaluator;
    evaluator.setSourceModule(moduleName);
    if (ext === '.cls' && !source.trim().toLowerCase().startsWith('class ') && !source.toLowerCase().includes('end class')) {
        source = `Class ${moduleName}\n${source}\nEnd Class`;
    }
    const ast = new Parser(new Lexer(source).tokenize()).parse();
    evaluator.evaluate(ast);
}

function runSuite(label: string, vbaTest: VBARunner): void {
    const evaluator = (vbaTest as any).evaluator;
    const allProcs = collectProcedures(evaluator.env);

    for (const [procName] of allProcs) {
        if (typeof procName !== 'string') continue;
        const lower = procName.toLowerCase();
        const baseProcName = lower.includes(':') ? lower.split(':')[1] : lower;
        const moduleName = lower.includes(':') ? lower.split(':')[0] : '';

        if (!baseProcName.startsWith('test')) continue;

        totalTests++;
        console.log(`[Test] ${label} / ${procName}`);

        try {
            if (moduleName && allProcs.has(`${moduleName}:setup`)) {
                vbaTest.run(`${moduleName}:setup`, []);
            }

            const wrapperCode = [
                `Sub __test_wrapper__()`,
                `    Dim assert As New AssertHelper`,
                `    On Error Resume Next`,
                `    ${baseProcName} assert`,
                `    On Error GoTo 0`,
                `    If assert.Failed Then`,
                `        Err.Raise vbObjectError + 1, "__test_wrapper__", assert.FailMessage`,
                `    End If`,
                `End Sub`
            ].join('\n');
            const ast = new Parser(new Lexer(wrapperCode).tokenize()).parse();
            evaluator.evaluate(ast);
            evaluator.callProcedure('__test_wrapper__', []);

            totalPass++;
            console.log(`  ✅ Pass\n`);

            if (moduleName && allProcs.has(`${moduleName}:teardown`)) {
                vbaTest.run(`${moduleName}:teardown`, []);
            }
        } catch (e: any) {
            console.error(`  ❌ FAILED: ${e.message}\n`);
            allFailed.push(`${label} / ${procName}`);
        }
    }
}

// --- 1. Top-level suite ---
const topLevelFiles = fs.readdirSync(vbaDir).filter(f => {
    const fullPath = path.join(vbaDir, f);
    return VBA_EXTS.has(path.extname(f).toLowerCase()) &&
           !f.toLowerCase().includes('run-all-vba-tests') &&
           fs.statSync(fullPath).isFile();
});

if (topLevelFiles.length > 0) {
    console.log(`Top-level suite (${topLevelFiles.length} file(s)):`);
    topLevelFiles.forEach(f => console.log(`  - ${f}`));
    console.log('');
    totalFiles += topLevelFiles.length;
    runSuite('top-level', new VBARunner(vbaDir));
}

// --- 2. Subdirectory suites ---
const assertHelperPath = path.join(vbaDir, 'AssertHelper.cls');

const subdirs = fs.readdirSync(vbaDir).filter(f => {
    return fs.statSync(path.join(vbaDir, f)).isDirectory();
});

for (const subdir of subdirs) {
    const subdirPath = path.join(vbaDir, subdir);
    const subdirFiles = fs.readdirSync(subdirPath).filter(f =>
        VBA_EXTS.has(path.extname(f).toLowerCase())
    );

    if (subdirFiles.length === 0) continue;

    console.log(`Suite: ${subdir} (${subdirFiles.length} file(s)):`);
    subdirFiles.forEach(f => console.log(`  - ${f}`));
    console.log('');
    totalFiles += subdirFiles.length;

    const subdirTest = new VBARunner(subdirPath);
    if (fs.existsSync(assertHelperPath)) {
        injectFile(subdirTest, assertHelperPath);
    }
    runSuite(subdir, subdirTest);
}

// --- Summary ---
console.log('=== Summary ===');
console.log(`Files loaded: ${totalFiles}`);
console.log(`Tests run: ${totalTests}`);
console.log(`Tests passed: ${totalPass}`);

if (allFailed.length > 0) {
    console.log(`Tests failed: ${allFailed.length}`);
    console.log('\nFailed tests:');
    allFailed.forEach(name => console.log(`  - ${name}`));
    process.exit(1);
} else if (totalTests > 0) {
    console.log(`\n✅ All ${totalTests} test(s) passed!`);
    process.exit(0);
} else {
    console.log('\n⚠️  No tests found to run');
    process.exit(0);
}
