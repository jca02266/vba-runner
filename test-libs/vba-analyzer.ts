// VBA Code Analyzer
//
// レガシーVBAのリファクタリング支援を目的とした静的解析ツール。
// 単一ファイルまたはディレクトリを引数に取り、リファクタリング候補となる
// 問題箇所（巨大プロシージャ・深いネスト・連続代入・Excel依存・繰り返し数値リテラル等）を列挙する。
//
// Usage:
//   ./node_modules/.bin/esbuild test-libs/vba-analyzer.ts --bundle --outfile=test-libs/vba-analyzer.cjs --platform=node
//   node test-libs/vba-analyzer.cjs <path>            # テキスト形式
//   node test-libs/vba-analyzer.cjs <path> --json     # JSON 形式
//
// 既知の制約（TODO_NEXT.md「実証実験で判明したギャップ」を参照）:
//   - Expression レベルの loc が欠落しているため、Excel I/O アクセス等の行番号が -1 になる
//   - 連続代入の「形状クラスタリング」は未実装（件数のみ）
//   - データフロー解析（Def-Use）は未実装
//   - 接頭辞クラスタによる UDT 抽出候補検出は未実装

import { Lexer } from '../src/compiler/lexer';
import { Parser } from '../src/compiler/parser';
import * as fs from 'fs';
import * as path from 'path';

interface ProcedureMetrics {
    name: string;
    kind: 'Sub' | 'Function' | 'Property';
    startLine: number;
    endLine: number;
    lineCount: number;
    maxNestDepth: number;
    localDeclCount: number;
    assignmentBlocks: Array<{ startLine: number; endLine: number; count: number; context: string }>;
    excelAccessCount: number;
    excelAccessSamples: Array<{ line: number; expr: string }>;
    repeatedNumericLiterals: Array<{ value: string; occurrences: number; lines: number[] }>;
}

interface FileReport {
    filePath: string;
    totalLines: number;
    procedureCount: number;
    procedures: ProcedureMetrics[];
    warnings: string[];
}

// ---------------------------------------------------------
// AST walker helpers
// ---------------------------------------------------------

function isStatementArray(v: any): v is any[] {
    return Array.isArray(v);
}

function measureNestDepth(node: any, depth: number, current: { max: number }): void {
    if (!node) return;
    if (depth > current.max) current.max = depth;
    const nestingTypes = new Set(['IfStatement', 'ForStatement', 'WhileStatement', 'DoStatement', 'SelectStatement', 'WithStatement']);

    const walk = (children: any, childDepth: number) => {
        if (!children) return;
        if (Array.isArray(children)) {
            for (const c of children) {
                const nextDepth = nestingTypes.has(c?.type) ? childDepth + 1 : childDepth;
                measureNestDepth(c, nextDepth, current);
            }
        } else {
            const nextDepth = nestingTypes.has(children?.type) ? childDepth + 1 : childDepth;
            measureNestDepth(children, nextDepth, current);
        }
    };

    walk(node.body, depth);
    walk(node.consequent, depth);
    walk(node.alternate, depth);
    if (node.elseIfClauses) for (const c of node.elseIfClauses) walk(c.consequent, depth);
    if (node.cases) for (const c of node.cases) walk(c.body, depth);
}

function countLocalDeclarations(stmts: any): number {
    if (!isStatementArray(stmts)) return 0;
    let n = 0;
    for (const s of stmts) {
        if (s.type === 'VariableDeclaration') n += s.declarations?.length || 1;
        if (s.type === 'ConstDeclaration') n += 1;
        if (s.body) n += countLocalDeclarations(s.body);
        if (s.consequent) n += countLocalDeclarations(s.consequent);
        if (s.alternate) n += countLocalDeclarations(Array.isArray(s.alternate) ? s.alternate : [s.alternate]);
        if (s.elseIfClauses) for (const c of s.elseIfClauses) n += countLocalDeclarations(c.consequent || []);
        if (s.cases) for (const c of s.cases) n += countLocalDeclarations(c.body || []);
    }
    return n;
}

function findConsecutiveAssignmentBlocks(
    stmts: any,
    minRun = 5,
    context = 'root',
): Array<{ startLine: number; endLine: number; count: number; context: string }> {
    const out: Array<{ startLine: number; endLine: number; count: number; context: string }> = [];
    if (!isStatementArray(stmts)) return out;
    let runStart = -1, runEnd = -1, runCount = 0;
    const flush = () => {
        if (runCount >= minRun) out.push({ startLine: runStart, endLine: runEnd, count: runCount, context });
        runStart = -1; runEnd = -1; runCount = 0;
    };
    const isAssignmentLike = (s: any) =>
        s.type === 'AssignmentStatement' || s.type === 'SetStatement' || s.type === 'VariableDeclaration';

    for (const s of stmts) {
        if (isAssignmentLike(s) && s.loc) {
            if (runStart === -1) runStart = s.loc.start.line;
            runEnd = s.loc.end.line;
            runCount++;
        } else {
            flush();
        }
        const children = [
            ['If', s.consequent],
            ['Else', Array.isArray(s.alternate) ? s.alternate : null],
            ['Body', s.body],
        ] as const;
        for (const [label, child] of children) {
            if (child) out.push(...findConsecutiveAssignmentBlocks(child, minRun, `${context}>${s.type}/${label}`));
        }
    }
    flush();
    return out;
}

function findExcelAccess(stmts: any, samplesLimit = 10): { count: number; samples: Array<{ line: number; expr: string }> } {
    const excelRoots = new Set(['sheets', 'range', 'cells', 'application', 'activesheet', 'activeworkbook', 'thisworkbook', 'workbook', 'worksheet', 'ws']);
    let count = 0;
    const samples: Array<{ line: number; expr: string }> = [];

    function visit(node: any) {
        if (!node || typeof node !== 'object') return;
        if (node.type === 'MemberAccess' || node.type === 'MemberExpression') {
            const objName = (node.object?.name || node.object?.callee?.name || '').toLowerCase();
            if (excelRoots.has(objName)) {
                count++;
                if (samples.length < samplesLimit) {
                    samples.push({
                        line: node.loc?.start.line ?? -1,
                        expr: `${objName}.${node.property?.name || '?'}`,
                    });
                }
            }
        }
        for (const key of Object.keys(node)) {
            const v = node[key];
            if (Array.isArray(v)) v.forEach(visit);
            else if (v && typeof v === 'object' && v.type) visit(v);
        }
    }

    if (isStatementArray(stmts)) for (const s of stmts) visit(s);
    return { count, samples };
}

function findRepeatedNumericLiterals(
    stmts: any,
    minOccurrences = 2,
): Array<{ value: string; occurrences: number; lines: number[] }> {
    const buckets = new Map<string, number[]>();
    function visit(node: any) {
        if (!node || typeof node !== 'object') return;
        if (node.type === 'NumberLiteral' && typeof node.value === 'number') {
            const v = node.value;
            // 0, 1, -1 と単純な小整数は除外（マジックナンバーになりにくい）
            const isTrivial = (v === 0 || v === 1 || v === -1) || (Number.isInteger(v) && v >= 2 && v <= 2);
            if (!isTrivial) {
                const key = String(v);
                if (!buckets.has(key)) buckets.set(key, []);
                buckets.get(key)!.push(node.loc?.start.line ?? -1);
            }
        }
        for (const key of Object.keys(node)) {
            const v = node[key];
            if (Array.isArray(v)) v.forEach(visit);
            else if (v && typeof v === 'object' && v.type) visit(v);
        }
    }
    if (isStatementArray(stmts)) for (const s of stmts) visit(s);

    return [...buckets.entries()]
        .filter(([, lines]) => lines.length >= minOccurrences)
        .map(([value, lines]) => ({ value, occurrences: lines.length, lines }))
        .sort((a, b) => b.occurrences - a.occurrences);
}

// ---------------------------------------------------------
// 単一ファイル解析
// ---------------------------------------------------------

function analyzeFile(filePath: string): FileReport {
    const src = fs.readFileSync(filePath, 'utf-8');
    const lines = src.split('\n');
    const warnings: string[] = [];

    let ast: any;
    try {
        const tokens = new Lexer(src).tokenize();
        ast = new Parser(tokens).parse();
    } catch (e: any) {
        warnings.push(`Parse error: ${e.message}`);
        return { filePath, totalLines: lines.length, procedureCount: 0, procedures: [], warnings };
    }

    if (ast.diagnostics && ast.diagnostics.length > 0) {
        for (const d of ast.diagnostics) {
            warnings.push(`L${d.loc?.start.line ?? '?'}: ${d.message}`);
        }
    }

    const procs = ast.body.filter((s: any) => s.type === 'ProcedureDeclaration');
    const procedures: ProcedureMetrics[] = procs.map((proc: any) => {
        const startLine = proc.loc?.start.line ?? -1;
        const endLine = proc.loc?.end.line ?? -1;
        const nestState = { max: 0 };
        measureNestDepth(proc, 0, nestState);

        const kind: 'Sub' | 'Function' | 'Property' =
            proc.isProperty ? 'Property' : proc.isFunction ? 'Function' : 'Sub';

        const excel = findExcelAccess(proc.body);
        const numericLits = findRepeatedNumericLiterals(proc.body);

        return {
            name: proc.name?.name ?? '<anonymous>',
            kind,
            startLine,
            endLine,
            lineCount: endLine - startLine + 1,
            maxNestDepth: nestState.max,
            localDeclCount: countLocalDeclarations(proc.body),
            assignmentBlocks: findConsecutiveAssignmentBlocks(proc.body),
            excelAccessCount: excel.count,
            excelAccessSamples: excel.samples,
            repeatedNumericLiterals: numericLits.slice(0, 10),
        };
    });

    return {
        filePath,
        totalLines: lines.length,
        procedureCount: procedures.length,
        procedures,
        warnings,
    };
}

// ---------------------------------------------------------
// テキストフォーマッタ
// ---------------------------------------------------------

function formatReport(r: FileReport): string {
    const out: string[] = [];
    out.push(`=== ${r.filePath} ===`);
    out.push(`  総行数: ${r.totalLines}`);
    out.push(`  プロシージャ数: ${r.procedureCount}`);
    if (r.warnings.length) {
        out.push(`  ⚠ 警告:`);
        for (const w of r.warnings) out.push(`    ${w}`);
    }

    for (const p of r.procedures) {
        out.push('');
        out.push(`  [${p.kind}] ${p.name}  (${p.lineCount}行, L${p.startLine}-L${p.endLine})`);
        const flags: string[] = [];
        if (p.lineCount >= 100) flags.push(`📏 LARGE (${p.lineCount}行)`);
        if (p.maxNestDepth >= 5) flags.push(`🌀 DEEP_NEST (${p.maxNestDepth}段)`);
        if (p.localDeclCount >= 30) flags.push(`🧮 MANY_LOCALS (${p.localDeclCount}個)`);
        if (p.excelAccessCount >= 10) flags.push(`📊 EXCEL_HEAVY (${p.excelAccessCount}件)`);
        if (flags.length) out.push(`    フラグ: ${flags.join(' / ')}`);
        else out.push(`    最大ネスト ${p.maxNestDepth} / Dim ${p.localDeclCount} / Excel ${p.excelAccessCount}`);

        if (p.assignmentBlocks.length) {
            out.push(`    連続代入ブロック（抽出候補）:`);
            for (const b of p.assignmentBlocks) {
                out.push(`      L${b.startLine}-L${b.endLine}: ${b.count}件 [${b.context}]`);
            }
        }
        if (p.excelAccessSamples.length) {
            out.push(`    Excel I/O サンプル:`);
            for (const e of p.excelAccessSamples.slice(0, 5)) {
                out.push(`      L${e.line === -1 ? '?' : e.line}: ${e.expr}`);
            }
        }
        if (p.repeatedNumericLiterals.length) {
            const top = p.repeatedNumericLiterals.slice(0, 5);
            out.push(`    繰り返し数値リテラル: ${top.map(n => `${n.value}(×${n.occurrences})`).join(', ')}`);
        }
    }
    return out.join('\n');
}

// ---------------------------------------------------------
// CLI
// ---------------------------------------------------------

function collectVbaFiles(target: string): string[] {
    const stat = fs.statSync(target);
    if (stat.isFile()) return [target];
    if (!stat.isDirectory()) return [];
    const out: string[] = [];
    function walk(dir: string) {
        for (const name of fs.readdirSync(dir)) {
            const full = path.join(dir, name);
            const st = fs.statSync(full);
            if (st.isDirectory()) walk(full);
            else if (/\.(vba|bas|cls|frm)$/i.test(name)) out.push(full);
        }
    }
    walk(target);
    return out.sort();
}

function main() {
    const args = process.argv.slice(2);
    if (args.length === 0 || args.includes('--help')) {
        console.log('Usage: node vba-analyzer.cjs <file-or-dir> [--json]');
        console.log('');
        console.log('  --json   JSON 形式で出力（プログラム連携用）');
        process.exit(args.length === 0 ? 1 : 0);
    }
    const wantJson = args.includes('--json');
    const target = args.find(a => !a.startsWith('--'));
    if (!target) {
        console.error('対象ファイルまたはディレクトリを指定してください');
        process.exit(1);
    }

    const files = collectVbaFiles(target);
    if (files.length === 0) {
        console.error(`VBA ファイルが見つかりません: ${target}`);
        process.exit(1);
    }

    const reports = files.map(analyzeFile);

    if (wantJson) {
        console.log(JSON.stringify(reports, null, 2));
    } else {
        for (const r of reports) console.log(formatReport(r));
    }
}

main();
