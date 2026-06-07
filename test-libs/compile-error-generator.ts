/**
 * CompileError.bas を解析して vba_compile_error.test.ts を生成するツール。
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

    let inPreamble = false;
    let currentCase: Partial<CompileErrorCase> & { code?: string[] } | null = null;
    let inCode = false;

    for (let i = 0; i < lines.length; i++) {
        const raw = lines[i];
        const trimmed = raw.trim();

        // PREAMBLE block
        if (trimmed === "' PREAMBLE:") {
            inPreamble = true;
            continue;
        }
        if (inPreamble) {
            if (trimmed.startsWith("' ") || trimmed === "'") {
                const content = trimmed === "'" ? '' : trimmed.slice(2);
                preamble.push(content);
                continue;
            } else {
                inPreamble = false;
            }
        }

        // CASE block start
        if (trimmed.startsWith("' CASE:")) {
            if (currentCase && currentCase.name) {
                cases.push(currentCase as CompileErrorCase);
            }
            currentCase = { name: trimmed.slice(7).trim(), code: [] };
            inCode = false;
            continue;
        }

        if (!currentCase) continue;

        if (trimmed.startsWith("' TYPE:")) {
            const t = trimmed.slice(7).trim();
            currentCase.type = t as 'parse' | 'prerun';
        } else if (trimmed.startsWith("' VBA:")) {
            currentCase.vbaError = trimmed.slice(6).trim();
        } else if (trimmed.startsWith("' RUNNER:")) {
            currentCase.runnerPattern = trimmed.slice(9).trim();
        } else if (trimmed === "' CODE:") {
            inCode = true;
        } else if (inCode) {
            if (trimmed.startsWith("' ") || trimmed === "'") {
                const content = trimmed === "'" ? '' : trimmed.slice(2);
                currentCase.code!.push(content);
            } else {
                inCode = false;
                cases.push(currentCase as CompileErrorCase);
                currentCase = null;
            }
        }
    }

    if (currentCase && currentCase.name) {
        cases.push(currentCase as CompileErrorCase);
    }

    return { preamble, cases };
}

function generateTestFile(preamble: string[], cases: CompileErrorCase[], sourceFile: string): string {
    const preambleCode = preamble.join('\n');

    const testBlocks = cases.map(c => {
        const codeLines = c.code.map(l => `    ${l}`).join('\n');
        const pattern = c.runnerPattern; // e.g. /syntax error/i

        if (c.type === 'parse') {
            return `
  test('${c.name} - parse error', () => {
    // VBA: ${c.vbaError}
    const src = \`
${codeLines}
\`;
    expect(() => {
      const tokens = new Lexer(src).tokenize();
      new Parser(tokens).parse();
    }).toThrow(${pattern});
  });`;
        } else {
            return `
  test('${c.name} - prerun compile error', () => {
    // VBA: ${c.vbaError}
    expect(() => evalVBASingle(\`
${preamble.map(l => `      ${l}`).join('\n')}
      Sub __test__()
${c.code.map(l => `        ${l}`).join('\n')}
      End Sub
      __test__
    \`)).toThrow(${pattern});
  });`;
        }
    });

    return `/**
 * VBA コンパイルエラー自動テスト
 * このファイルは ${path.basename(sourceFile)} から自動生成されました。
 * 再生成: npx tsx test-libs/compile-error-generator.ts ${sourceFile}
 *
 * [parse]  Parser.parse() 時に例外が発生するケース
 * [prerun] プロシージャ呼び出し直前の静的チェックで例外が発生するケース
 */

import { Lexer } from '../src/engine/lexer';
import { Parser } from '../src/engine/parser';
import { evalVBASingle } from '../test-libs/test-runner';

const _PREAMBLE = \`
${preamble.map(l => `  ${l}`).join('\n')}
\`;

describe('VBA compile errors', () => {
${testBlocks.join('\n')}
});
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

    const output = generateTestFile(preamble, cases, inputPath);

    if (outputPath) {
        fs.writeFileSync(outputPath, output, 'utf-8');
        console.error(`✓ Generated: ${outputPath}`);
    } else {
        console.log(output);
    }
}

export { parseCompileErrorBas, generateTestFile };
