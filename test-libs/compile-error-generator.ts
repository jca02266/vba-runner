/**
 * CompileError.bas を解析して vba_compile_error.test.ts を生成するツール。
 *
 * CompileError.bas の構造:
 *   - Private Sub/Function ... End Sub/Function  → PREAMBLE（ヘルパー定義）として扱う
 *
 *   通常ケース（Sub Case_<name>() でラップ）:
 *   - ' CASE: name
 *     ' TYPE: parse | resolve
 *     ' VBA: <VBE のエラーメッセージ>
 *     ' RUNNER: <正規表現（VBA の意図を包含した VBARunner の期待メッセージ）>
 *     Sub Case_<name>()
 *       Dim v
 *       v = MySub ' @error   ← エラー行末尾に ' @error マーカーを付ける
 *                              ジェネレーターが行番号を自動計算し、マーカーを除去する
 *       ...body...
 *     End Sub
 *
 *   モジュールレベルケース（'@case-begin / '@case-end でラップ）:
 *   - ' CASE: name
 *     ...メタ行...
 *     '@case-begin
 *     Sub Foo()         ← モジュールレベルのコードをそのまま記述
 *     End Sub
 *     ''Sub Foo()  ' @error   ← VBE でのコンパイルエラーを避けるため '' でコメントアウト
 *     ''End Sub          ← ジェネレーターが '' を除去して実コードとして展開
 *     '@case-end
 *
 * Usage:
 *   npx tsx test-libs/compile-error-generator.ts tests/vba/CompileError.bas
 *   npx tsx test-libs/compile-error-generator.ts tests/vba/CompileError.bas --output tests/spec/vba_compile_error.test.ts
 */

import * as fs from 'fs';
import * as path from 'path';

interface CompileErrorCase {
    name: string;
    type: 'parse' | 'resolve' | 'preproc' | 'exec';
    vbaError: string;
    errorLine: number | null;  // Sub ボディ内のエラー行（1 始まり）。' @error マーカーから自動計算。
    runnerPattern: string;
    code: string[];  // ' @error マーカーを除去済みのコード行
    isModuleLevel?: boolean;  // true のとき code はモジュールレベルに展開（Sub __test__() でラップしない）
    procName?: string;        // preproc/exec: callProcedure に渡すプロシージャ名（PROC: メタから取得）
    evalOptions?: string;     // resolve: Evaluator 設定の上書き（EVAL_OPTIONS: メタから取得。例: { allowTopLevelStatements: false }）
}

// ' @error マーカーを除去する正規表現
const ERROR_MARKER = /\s*'\s*@error\s*$/i;

function parseCompileErrorBas(source: string): { preamble: string[], cases: CompileErrorCase[] } {
    const lines = source.split('\n');
    const preamble: string[] = [];
    const cases: CompileErrorCase[] = [];

    let meta: Partial<CompileErrorCase> | null = null;
    let inCaseSub = false;
    let currentSubName: string | null = null;
    let caseBody: string[] = [];
    let caseErrorLine: number | null = null;
    let inPreambleSub = false;
    let preambleBuffer: string[] = [];
    // '@case-begin / '@case-end で囲まれたモジュールレベルケースの収集状態
    let inModuleLevelCase = false;
    let moduleCaseLines: string[] = [];
    let moduleCaseErrorLine: number | null = null;

    for (const raw of lines) {
        const trimmed = raw.trim();

        // CASE メタコメント — 新しい CASE が来たら meta を上書き（前の meta は消去）
        if (trimmed.startsWith("' CASE:")) {
            meta = { name: trimmed.slice(7).trim() };
            continue;
        }
        if (meta && trimmed.startsWith("' TYPE:")) {
            meta.type = trimmed.slice(7).trim() as 'parse' | 'resolve' | 'preproc' | 'exec';
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
        if (meta && trimmed.startsWith("' PROC:")) {
            meta.procName = trimmed.slice(7).trim();
            continue;
        }
        if (meta && trimmed.startsWith("' EVAL_OPTIONS:")) {
            meta.evalOptions = trimmed.slice(15).trim();
            continue;
        }
        // NOTE: は読み飛ばす（テストには不要）
        if (meta && trimmed.startsWith("' NOTE:")) continue;

        // '@case-begin: モジュールレベルケースの開始
        if (trimmed === "'@case-begin") {
            inModuleLevelCase = true;
            moduleCaseLines = [];
            moduleCaseErrorLine = null;
            continue;
        }

        // '@case-begin ～ '@case-end の間: '' を除去してコードを収集
        if (inModuleLevelCase) {
            if (trimmed === "'@case-end") {
                while (moduleCaseLines.length > 0 && moduleCaseLines[moduleCaseLines.length - 1].trim() === '') {
                    moduleCaseLines.pop();
                }
                // PROC: 未指定の場合はブロック内の最初の Sub/Function 名を自動検出
                let resolvedProcName = meta!.procName;
                if (!resolvedProcName) {
                    for (const l of moduleCaseLines) {
                        const m = l.match(/^\s*(?:Public\s+|Private\s+)?(?:Sub|Function)\s+(\w+)\s*\(/i);
                        if (m) { resolvedProcName = m[1]; break; }
                    }
                }
                cases.push({
                    name: meta!.name!,
                    type: meta!.type ?? 'resolve',
                    vbaError: meta!.vbaError ?? '',
                    errorLine: moduleCaseErrorLine,
                    runnerPattern: meta!.runnerPattern ?? '/.+/',
                    code: moduleCaseLines,
                    isModuleLevel: true,
                    procName: resolvedProcName,
                    evalOptions: meta!.evalOptions,
                });
                inModuleLevelCase = false;
                meta = null;
            } else {
                let codeLine = raw.startsWith("''") ? raw.slice(2) : raw;
                if (ERROR_MARKER.test(codeLine)) {
                    moduleCaseErrorLine = moduleCaseLines.length + 1;
                    codeLine = codeLine.replace(ERROR_MARKER, '');
                }
                moduleCaseLines.push(codeLine);
            }
            continue;
        }

        // Case_* Sub の開始（メタがある場合）
        // Sub 名全体（Case_ 以降）をテストケース名として使う。
        // 例: Sub Case_foo1() → name = 'foo1'
        //     Sub Case_foo()  → name = 'foo'
        const caseSubMatch = raw.match(/^\s*Sub\s+Case_(\w+)\s*\(\s*\)/i);
        if (meta && caseSubMatch) {
            inCaseSub = true;
            currentSubName = caseSubMatch[1];
            caseBody = [];
            caseErrorLine = null;
            continue;
        }

        // Case_* Sub の終了
        if (inCaseSub && /^\s*End\s+Sub\s*$/i.test(trimmed)) {
            // 末尾の空行を除去
            while (caseBody.length > 0 && caseBody[caseBody.length - 1].trim() === '') {
                caseBody.pop();
            }
            cases.push({
                name: currentSubName!,
                type: meta!.type ?? 'parse',
                vbaError: meta!.vbaError ?? '',
                errorLine: caseErrorLine,
                runnerPattern: meta!.runnerPattern ?? '/.+/',
                code: caseBody,
                procName: meta!.procName,
                evalOptions: meta!.evalOptions,
            });
            // meta はクリアしない — 同一 CASE コメントに複数 Sub を許容
            inCaseSub = false;
            currentSubName = null;
            caseBody = [];
            caseErrorLine = null;
            continue;
        }

        if (inCaseSub) {
            // Sub ボディの行（先頭インデントを 1 レベル除去）
            let codeLine = raw.replace(/^    /, '');
            // ' @error マーカーが付いていたら行番号を記録してマーカーを除去
            if (ERROR_MARKER.test(codeLine)) {
                caseErrorLine = caseBody.length + 1;  // 1 始まり
                codeLine = codeLine.replace(ERROR_MARKER, '');
            }
            caseBody.push(codeLine);
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

    // resolve テストでの行番号オフセット:
    //   テンプレートリテラルの構造:
    //     line 1: (空行 — テンプレート開始の改行)
    //     lines 2..N+1: preamble (N 行)
    //     line N+2: Sub __test__()
    //     lines N+3..: body
    //   よって body 行 L の絶対行番号 = 1 + preamble.length + 1 + L = preamble.length + L + 2
    const resolveLineOffset = preamble.length + 2;

    // ソース文字列を構築するヘルパー
    const buildParseSrc = (c: CompileErrorCase) => c.code.join('\n');
    // resolve: Sub __test__() でラップするが呼び出しなし（resolveIdentifiers が静的検出する）
    // preproc/exec: 同じテンプレートだが呼び出しは assertCompileErrorPreproc/Exec が行う
    const buildSubWrappedSrc = (c: CompileErrorCase) =>
        `\n${preamble.map(l => `      ${l}`).join('\n')}\n      Sub __test__()\n${c.code.map(l => `        ${l}`).join('\n')}\n      End Sub\n    `;
    const buildModuleLevelSrc = (c: CompileErrorCase) =>
        `\n${preamble.map(l => `      ${l}`).join('\n')}\n${c.code.map(l => `      ${l}`).join('\n')}\n    `;

    const testBlocks = cases.flatMap(c => {
        const tryWrap = (inner: string) => `
// [${c.type}] ${c.name}
// VBA: ${c.vbaError}${c.errorLine != null ? `\n// VBA error line (within Sub body): ${c.errorLine}` : ''}
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

        // RUNNER: TBD — CompileError.bas の RUNNER が未確定。
        // テストを必ず失敗させ、実際のエラーメッセージを表示して RUNNER 値の決定を促す。
        if (c.runnerPattern === 'TBD') {
            const src = c.isModuleLevel ? buildModuleLevelSrc(c) : (c.type === 'parse' ? buildParseSrc(c) : buildSubWrappedSrc(c));
            const procArg = c.procName ?? (c.isModuleLevel ? '' : '__test__');
            const optsArg = c.evalOptions ? `, ${c.evalOptions}` : '';
            return [tryWrap(
`        // RUNNER: TBD — 実際のエラーを表示して CompileError.bas を更新してください
        const _tbd_msg = captureCompileErrorMessage('${c.type}', \`${src}\`, '${procArg}'${optsArg});
        throw new Error('[${c.name}] RUNNER: TBD — fill in CompileError.bas. Actual error seen: ' + _tbd_msg);`
            )];
        }

        const pattern = c.runnerPattern;

        if (c.type === 'parse') {
            const src = buildParseSrc(c);
            const expectedLine = c.errorLine ?? 1;
            return tryWrap(
`        assertCompileErrorPass1(\`${src}\`, ${expectedLine}, ${pattern}, '${c.name}');`
            );
        }

        if (c.type === 'resolve') {
            const src = c.isModuleLevel ? buildModuleLevelSrc(c) : buildSubWrappedSrc(c);
            const absLine = c.errorLine != null
                ? (c.isModuleLevel ? 1 + preamble.length + c.errorLine : resolveLineOffset + c.errorLine)
                : null;
            const lineArg = absLine != null ? String(absLine) : 'undefined';
            const optsArg = c.evalOptions ? `, ${c.evalOptions}` : '';
            return tryWrap(
`        assertCompileErrorResolve(\`${src}\`, ${lineArg}, ${pattern}, '${c.name}'${optsArg});`
            );
        }

        if (c.type === 'preproc') {
            const src = c.isModuleLevel ? buildModuleLevelSrc(c) : buildSubWrappedSrc(c);
            const procArg = c.procName ?? (c.isModuleLevel ? '' : '__test__');
            const absLine = c.errorLine != null
                ? (c.isModuleLevel ? 1 + preamble.length + c.errorLine : resolveLineOffset + c.errorLine)
                : null;
            const lineArg = absLine != null ? String(absLine) : 'undefined';
            return tryWrap(
`        assertCompileErrorPreproc(\`${src}\`, '${procArg}', ${lineArg}, ${pattern}, '${c.name}');`
            );
        }

        // exec
        {
            const src = c.isModuleLevel ? buildModuleLevelSrc(c) : buildSubWrappedSrc(c);
            const procArg = c.procName ?? (c.isModuleLevel ? '' : '__test__');
            const absLine = c.errorLine != null
                ? (c.isModuleLevel ? 1 + preamble.length + c.errorLine : resolveLineOffset + c.errorLine)
                : null;
            const lineArg = absLine != null ? String(absLine) : 'undefined';
            return tryWrap(
`        assertCompileErrorExec(\`${src}\`, '${procArg}', ${lineArg}, ${pattern}, '${c.name}');`
            );
        }
    });

    return `/**
 * VBA コンパイルエラー自動テスト
 * このファイルは ${path.basename(sourceFile)} から自動生成されました。
 * 再生成: npx tsx test-libs/compile-error-generator.ts ${sourceFile} --output <このファイルのパス>
 *
 * [parse]  Parser.parse() 時に例外が発生するケース（行番号も検証）
 * [resolve] resolveIdentifiers で例外が発生するケース（実行なし）
 * [preproc] precheckProc（OE チェック）で例外が発生するケース
 * [exec]  precheckProc 後の実行中に例外が発生するケース
 */

import { assertCompileErrorPass1, assertCompileErrorResolve, assertCompileErrorPreproc, assertCompileErrorExec, captureCompileErrorMessage } from '${rel('test-libs/test-runner')}';

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
    cases.forEach(c => {
        const lineStr = c.errorLine != null ? ` @error=line${c.errorLine}` : ' (no @error marker)';
        console.error(`  [${c.type}]${lineStr} ${c.name}`);
    });

    const output = generateTestFile(preamble, cases, inputPath, outputPath);

    if (outputPath) {
        fs.writeFileSync(outputPath, output, 'utf-8');
        console.error(`✓ Generated: ${outputPath}`);
    } else {
        console.log(output);
    }
}

export { parseCompileErrorBas, generateTestFile };
