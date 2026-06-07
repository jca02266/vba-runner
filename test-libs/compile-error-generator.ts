/**
 * CompileError.bas を解析して vba_compile_error.test.ts を生成するツール。
 *
 * CompileError.bas の構造:
 *   - Private Sub/Function ... End Sub/Function  → PREAMBLE（ヘルパー定義）として扱う
 *   - ' CASE: name
 *     ' TYPE: parse | prerun
 *     ' VBA: <VBE のエラーメッセージ>
 *     ' RUNNER: <正規表現>
 *     Sub Case_<name>()
 *       ...body...
 *     End Sub
 *
 * Usage:
 *   npx tsx test-libs/compile-error-generator.ts tests/vba/CompileError.bas
 *   npx tsx test-libs/compile-error-generator.ts tests/vba/CompileError.bas --output tests/spec/vba_compile_error.test.ts
 */

import * as fs from 'fs';
import * as path from 'path';

interface CompileErrorCase {
    name: string;
    type: 'parse' | 'prerun';
    vbaError: string;
    runnerPattern: string;
    code: string[];
}

function parseCompileErrorBas(source: string): { preamble: string[], cases: CompileErrorCase[] } {
    const lines = source.split('\n');
    const preamble: string[] = [];
    const cases: CompileErrorCase[] = [];

    let meta: Partial<CompileErrorCase> | null = null;
    let inCaseSub = false;
    let caseBody: string[] = [];
    let inPreambleSub = false;
    let preambleBuffer: string[] = [];

    for (const raw of lines) {
        const trimmed = raw.trim();

        // CASE メタコメント
        if (trimmed.startsWith("' CASE:")) {
            meta = { name: trimmed.slice(7).trim() };
            continue;
        }
        if (meta && trimmed.startsWith("' TYPE:")) {
            meta.type = trimmed.slice(7).trim() as 'parse' | 'prerun';
            continue;
        }
        if (meta && trimmed.startsWith("' VBA:")) {
            meta.vbaError = trimmed.slice(6).trim();
            continue;
        }
        if (meta && trimmed.startsWith("' RUNNER:")) {
            meta.runnerPattern = trimmed.slice(9).trim();
            continue;
        }

        // Case_* Sub の開始（メタがある場合）
        if (meta && /^\s*Sub\s+Case_\w+\s*\(\s*\)/i.test(raw)) {
            inCaseSub = true;
            caseBody = [];
            continue;
        }

        // Case_* Sub の終了
        if (inCaseSub && /^\s*End\s+Sub\s*$/i.test(trimmed)) {
            // 末尾の空行を除去
            while (caseBody.length > 0 && caseBody[caseBody.length - 1].trim() === '') {
                caseBody.pop();
            }
            cases.push({
                name: meta!.name!,
                type: meta!.type ?? 'parse',
                vbaError: meta!.vbaError ?? '',
                runnerPattern: meta!.runnerPattern ?? '/.+/',
                code: caseBody,
            });
            meta = null;
            inCaseSub = false;
            caseBody = [];
            continue;
        }

        if (inCaseSub) {
            // Sub ボディの行（先頭インデントを 1 レベル除去）
            caseBody.push(raw.replace(/^    /, ''));
            continue;
        }

        // PREAMBLE: Private Sub/Function ... End Sub/Function
        if (!meta && /^\s*Private\s+(Sub|Function)\s+/i.test(raw)) {
            inPreambleSub = true;
            preambleBuffer = [raw];
            continue;
        }
        if (inPreambleSub) {
            preambleBuffer.push(raw);
            if (/^\s*End\s+(Sub|Function)\s*$/i.test(trimmed)) {
                preamble.push(...preambleBuffer);
                preamble.push('');
                inPreambleSub = false;
                preambleBuffer = [];
            }
        }
    }

    return { preamble, cases };
}

function generateTestFile(preamble: string[], cases: CompileErrorCase[], sourceFile: string, outputFile?: string): string {
    // import パスを出力ファイルの位置から計算する
    const outDir = outputFile ? path.dirname(path.resolve(outputFile)) : path.resolve('tests/spec');
    const root = path.resolve('.');
    const rel = (target: string) => {
        const r = path.relative(outDir, path.join(root, target));
        return r.startsWith('.') ? r : './' + r;
    };

    const testBlocks = cases.map(c => {
        const pattern = c.runnerPattern;

        const tryWrap = (inner: string) => `
// [${c.type}] ${c.name}
// VBA: ${c.vbaError}
{
    try {
${inner}
        console.log('[PASS] ${c.name}');
        __pass__++;
    } catch (e: any) {
        console.error('[FAIL] ${c.name}:', e.message);
        __fail__++;
    }
}`;

        if (c.type === 'parse') {
            const codeLines = c.code.map(l => `    ${l}`).join('\n');
            return tryWrap(`        const src = \`
${codeLines}
\`;
        assert.throwsMatch(() => {
            const tokens = new Lexer(src).tokenize();
            new Parser(tokens).parse();
        }, ${pattern}, '${c.name}');`);
        } else {
            const bodyLines = c.code.map(l => `        ${l}`).join('\n');
            return tryWrap(`        assert.throwsMatch(() => evalVBASingle(\`
${preamble.map(l => `      ${l}`).join('\n')}
      Sub __test__()
${bodyLines}
      End Sub
      __test__
    \`), ${pattern}, '${c.name}');`);
        }
    });

    return `/**
 * VBA コンパイルエラー自動テスト
 * このファイルは ${path.basename(sourceFile)} から自動生成されました。
 * 再生成: npx tsx test-libs/compile-error-generator.ts ${sourceFile} --output <このファイルのパス>
 *
 * [parse]  Parser.parse() 時に例外が発生するケース
 * [prerun] プロシージャ呼び出し直前の静的チェックで例外が発生するケース
 */

import { Lexer } from '${rel('src/engine/lexer')}';
import { Parser } from '${rel('src/engine/parser')}';
import { evalVBASingle, assert } from '${rel('test-libs/test-runner')}';

let __pass__ = 0, __fail__ = 0;
${testBlocks.join('\n')}
console.log(\`\\n=== Summary: \${__pass__} passed, \${__fail__} failed ===\`);
if (__fail__ > 0) process.exit(1);
`;
}

// CLI 実行
if (typeof process !== 'undefined' && process.argv[1]?.includes('compile-error-generator')) {
    const args = process.argv.slice(2).filter(a => !a.startsWith('--'));
    const outputFlag = process.argv.indexOf('--output');
    const outputPath = outputFlag !== -1 ? process.argv[outputFlag + 1] : undefined;

    if (args.length === 0) {
        console.error('Usage: compile-error-generator.ts <CompileError.bas> [--output <test.ts>]');
        process.exit(1);
    }

    const inputPath = args[0];
    if (!fs.existsSync(inputPath)) {
        console.error(`File not found: ${inputPath}`);
        process.exit(1);
    }

    const source = fs.readFileSync(inputPath, 'utf-8');
    const { preamble, cases } = parseCompileErrorBas(source);

    console.error(`Parsed ${cases.length} case(s) from ${inputPath}`);
    cases.forEach(c => console.error(`  [${c.type}] ${c.name}`));

    const output = generateTestFile(preamble, cases, inputPath, outputPath);

    if (outputPath) {
        fs.writeFileSync(outputPath, output, 'utf-8');
        console.error(`✓ Generated: ${outputPath}`);
    } else {
        console.log(output);
    }
}

export { parseCompileErrorBas, generateTestFile };
