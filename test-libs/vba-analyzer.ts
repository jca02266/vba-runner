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
    scope: 'public' | 'private' | 'friend';
    startLine: number;
    endLine: number;
    lineCount: number;
    maxNestDepth: number;
    localDeclCount: number;
    assignmentBlocks: Array<{ startLine: number; endLine: number; count: number; context: string; shape: string }>;
    excelAccessCount: number;
    excelAccessSamples: Array<{ line: number; expr: string }>;
    excelObjectsUsed: string[];           // 例: ['Sheets', 'Range', 'Application'] - モック必要候補
    repeatedNumericLiterals: Array<{ value: string; occurrences: number; lines: number[] }>;
    referenceCount: number;               // 他のプロシージャから呼ばれている回数（クロスファイル含む）
}

interface PrefixCluster {
    prefix: string;
    members: string[];
    suggestion: string;
}

interface FileReport {
    filePath: string;
    totalLines: number;
    procedureCount: number;
    procedures: ProcedureMetrics[];
    prefixClusters: PrefixCluster[];
    warnings: string[];
}

interface CallGraphEdge {
    from: string;
    fromFile: string;
    to: string;
    toFile: string;
}

interface WorkspaceReport {
    files: FileReport[];
    entryPointCandidates: Array<{ file: string; name: string; kind: string; line: number; reason: string }>;
    deadCodeCandidates: Array<{ file: string; name: string; kind: string; line: number }>;
    excelMockTargets: Array<{ file: string; procName: string; objects: string[]; sampleAccesses: string[] }>;
    callGraph: CallGraphEdge[];
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

// Classify a single assignment-like statement into a shape category.
function classifyStmtShape(s: any): string {
    if (s.type === 'ConstDeclaration') return 'const-decl';
    if (s.type === 'VariableDeclaration') return 'dim-decl';
    if (s.type === 'SetStatement') {
        const rhs = s.right || s.value;
        if (rhs?.type === 'MemberExpression' || rhs?.type === 'CallExpression') {
            const root = (rhs.object?.name || rhs.callee?.name || '').toLowerCase();
            if (EXCEL_ROOT_OBJECTS.has(root)) return 'set-excel';
        }
        return 'set-obj';
    }
    if (s.type === 'AssignmentStatement') {
        const lhs = s.left;
        const rhs = s.right;
        // LHS is Excel object?
        const lhsRoot = (lhs?.object?.name || '').toLowerCase();
        if (EXCEL_ROOT_OBJECTS.has(lhsRoot)) return 'range-write';
        // RHS is Excel object?
        const rhsRoot = (rhs?.object?.name || rhs?.callee?.object?.name || '').toLowerCase();
        if (EXCEL_ROOT_OBJECTS.has(rhsRoot)) return 'range-read';
        // RHS is a simple literal → variable initialization
        const rhsType = rhs?.type;
        if (rhsType === 'NumberLiteral' || rhsType === 'StringLiteral' || rhsType === 'DateLiteral') return 'var-init';
        return 'assign';
    }
    return 'other';
}

// Compute dominant shape from a list of individual statement shapes.
function dominantShape(shapes: string[]): string {
    if (shapes.length === 0) return 'mixed';
    const counts = new Map<string, number>();
    for (const s of shapes) counts.set(s, (counts.get(s) ?? 0) + 1);
    const sorted = [...counts.entries()].sort((a, b) => b[1] - a[1]);
    const [topShape, topCount] = sorted[0];
    // All same → use that shape; majority (>50%) → "mostly-X"; otherwise mixed
    if (topCount === shapes.length) return topShape;
    if (topCount / shapes.length > 0.5) return `mostly-${topShape}`;
    return 'mixed';
}

function findConsecutiveAssignmentBlocks(
    stmts: any,
    minRun = 5,
    context = 'root',
): Array<{ startLine: number; endLine: number; count: number; context: string; shape: string }> {
    const out: Array<{ startLine: number; endLine: number; count: number; context: string; shape: string }> = [];
    if (!isStatementArray(stmts)) return out;
    let runStart = -1, runEnd = -1, runCount = 0;
    const runShapes: string[] = [];
    const flush = () => {
        if (runCount >= minRun) out.push({ startLine: runStart, endLine: runEnd, count: runCount, context, shape: dominantShape(runShapes) });
        runStart = -1; runEnd = -1; runCount = 0; runShapes.length = 0;
    };
    const isAssignmentLike = (s: any) =>
        s.type === 'AssignmentStatement' || s.type === 'SetStatement' ||
        s.type === 'VariableDeclaration' || s.type === 'ConstDeclaration';

    for (const s of stmts) {
        if (isAssignmentLike(s) && s.loc) {
            if (runStart === -1) runStart = s.loc.start.line;
            runEnd = s.loc.end.line;
            runCount++;
            runShapes.push(classifyStmtShape(s));
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

// Excel系の "ルート" オブジェクト名（モック化の必要性判定に使う）
const EXCEL_ROOT_OBJECTS = new Set([
    'sheets', 'range', 'cells', 'application',
    'activesheet', 'activeworkbook', 'activecell',
    'thisworkbook', 'workbook', 'workbooks',
    'worksheet', 'worksheets', 'columns', 'rows',
    'selection', 'workbookbeforesave',
]);

// 表示用の正規化（小文字 → よく使われる大文字表記）
const EXCEL_DISPLAY_NAMES: Record<string, string> = {
    sheets: 'Sheets',
    range: 'Range',
    cells: 'Cells',
    application: 'Application',
    activesheet: 'ActiveSheet',
    activeworkbook: 'ActiveWorkbook',
    activecell: 'ActiveCell',
    thisworkbook: 'ThisWorkbook',
    workbook: 'Workbook',
    workbooks: 'Workbooks',
    worksheet: 'Worksheet',
    worksheets: 'Worksheets',
    columns: 'Columns',
    rows: 'Rows',
    selection: 'Selection',
};

function findExcelAccess(stmts: any, samplesLimit = 10): {
    count: number;
    samples: Array<{ line: number; expr: string }>;
    objectsUsed: string[];
} {
    let count = 0;
    const samples: Array<{ line: number; expr: string }> = [];
    const objectsUsed = new Set<string>();

    function recordObject(name: string) {
        const lower = name.toLowerCase();
        objectsUsed.add(EXCEL_DISPLAY_NAMES[lower] ?? name);
    }

    function visit(node: any) {
        if (!node || typeof node !== 'object') return;
        if (node.type === 'MemberAccess' || node.type === 'MemberExpression') {
            const objName = (node.object?.name || node.object?.callee?.name || '').toLowerCase();
            if (EXCEL_ROOT_OBJECTS.has(objName)) {
                count++;
                recordObject(objName);
                if (samples.length < samplesLimit) {
                    samples.push({
                        line: node.loc?.start.line ?? -1,
                        expr: `${EXCEL_DISPLAY_NAMES[objName] ?? objName}.${node.property?.name || '?'}`,
                    });
                }
            }
        }
        // 識別子単体での Excel 参照（例: `Set ws = ActiveSheet` の右辺）
        if (node.type === 'Identifier' && typeof node.name === 'string') {
            const lower = node.name.toLowerCase();
            if (EXCEL_ROOT_OBJECTS.has(lower)) {
                recordObject(lower);
            }
        }
        for (const key of Object.keys(node)) {
            if (key === 'loc') continue;
            const v = node[key];
            if (Array.isArray(v)) v.forEach(visit);
            else if (v && typeof v === 'object' && v.type) visit(v);
        }
    }

    if (isStatementArray(stmts)) for (const s of stmts) visit(s);
    return { count, samples, objectsUsed: [...objectsUsed].sort() };
}

// プロシージャ本体から「呼び出している識別子名」を収集する
// 簡易ヒューリスティック: CallExpression の callee が Identifier の場合 / Call/サブ名直接呼びの場合
function collectProcedureCalls(stmts: any, definedTypes: Set<string>): Set<string> {
    const calls = new Set<string>();

    function visit(node: any) {
        if (!node || typeof node !== 'object') return;
        if (node.type === 'CallExpression') {
            if (node.callee?.type === 'Identifier' && node.callee.name) {
                calls.add(node.callee.name.toLowerCase());
            } else if (node.callee?.name) {
                calls.add(String(node.callee.name).toLowerCase());
            }
        }
        if (node.type === 'CallStatement' && node.name) {
            calls.add(String(node.name).toLowerCase());
        }
        // 型名（UDT）として登場するものは呼び出しとみなさない
        for (const key of Object.keys(node)) {
            if (key === 'loc') continue;
            const v = node[key];
            if (Array.isArray(v)) v.forEach(visit);
            else if (v && typeof v === 'object' && v.type) visit(v);
        }
    }

    if (isStatementArray(stmts)) for (const s of stmts) visit(s);
    // UDT 型名は呼び出しではないので除外
    for (const t of definedTypes) calls.delete(t.toLowerCase());
    return calls;
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
// 接頭辞クラスタ検出（UDT抽出候補）
// ---------------------------------------------------------

// Extract the string name from a ConstDeclaration or VariableDeclaration node.
function extractDeclName(s: any): string | null {
    if (s.type === 'ConstDeclaration') {
        // name is an Identifier object: { type: 'Identifier', name: '...' }
        return s.name?.name ?? (typeof s.name === 'string' ? s.name : null);
    }
    return null;
}

// Collect all Const identifier names from module-level and all procedure bodies.
function collectTopLevelIdentifiers(stmts: any[]): string[] {
    const names: string[] = [];
    for (const s of stmts) {
        const n = extractDeclName(s);
        if (n) names.push(n);
        if (s.type === 'VariableDeclaration' && Array.isArray(s.declarations)) {
            for (const d of s.declarations) {
                const dName = d.name?.name ?? (typeof d.name === 'string' ? d.name : null);
                if (dName) names.push(dName);
            }
        }
        // Recurse into ProcedureDeclaration body
        if (s.type === 'ProcedureDeclaration' && Array.isArray(s.body)) {
            for (const inner of s.body) {
                const inner_n = extractDeclName(inner);
                if (inner_n) names.push(inner_n);
            }
        }
    }
    return names;
}

// Detect groups of identifiers sharing a common UPPER_CASE prefix (e.g. COL_, ROW_).
function detectPrefixClusters(stmts: any[], minClusterSize = 3): PrefixCluster[] {
    const names = collectTopLevelIdentifiers(stmts);
    // Only consider identifiers with at least one '_' separator and an uppercase-looking prefix
    const prefixBuckets = new Map<string, string[]>();
    for (const name of names) {
        const underPos = name.indexOf('_');
        if (underPos < 2) continue;  // too short prefix
        const prefix = name.slice(0, underPos + 1).toUpperCase();  // e.g. "COL_"
        if (!/^[A-Z][A-Z0-9]*_$/.test(prefix)) continue;
        if (!prefixBuckets.has(prefix)) prefixBuckets.set(prefix, []);
        prefixBuckets.get(prefix)!.push(name);
    }
    const clusters: PrefixCluster[] = [];
    for (const [prefix, members] of prefixBuckets) {
        if (members.length < minClusterSize) continue;
        const uniqueMembers = [...new Set(members)].sort();
        clusters.push({
            prefix,
            members: uniqueMembers,
            suggestion: `Extract ${uniqueMembers.length} constants with prefix "${prefix}" to a VBA Type or Enum`,
        });
    }
    return clusters.sort((a, b) => b.members.length - a.members.length);
}

// ---------------------------------------------------------
// 単一ファイル解析
// ---------------------------------------------------------

interface FileAnalysis {
    report: FileReport;
    // クロスファイル分析に使う中間データ
    definedProcs: Map<string, { proc: any; kind: string; scope: string }>;
    definedTypes: Set<string>;
    callsByProc: Map<string, Set<string>>;  // procName -> called identifiers
}

function analyzeFile(filePath: string): FileAnalysis {
    const src = fs.readFileSync(filePath, 'utf-8');
    const lines = src.split('\n');
    const warnings: string[] = [];

    let ast: any;
    try {
        const tokens = new Lexer(src).tokenize();
        ast = new Parser(tokens).parse();
    } catch (e: any) {
        warnings.push(`Parse error: ${e.message}`);
        return {
            report: { filePath, totalLines: lines.length, procedureCount: 0, procedures: [], warnings },
            definedProcs: new Map(),
            definedTypes: new Set(),
            callsByProc: new Map(),
        };
    }

    if (ast.diagnostics && ast.diagnostics.length > 0) {
        for (const d of ast.diagnostics) {
            warnings.push(`L${d.loc?.start.line ?? '?'}: ${d.message}`);
        }
    }

    // 定義された型名（UDT）を集めておく - call と区別するため
    const definedTypes = new Set<string>();
    for (const s of ast.body) {
        if (s.type === 'TypeDeclaration' && s.name?.name) {
            definedTypes.add(s.name.name);
        }
        if (s.type === 'ClassDeclaration' && s.name) {
            definedTypes.add(s.name);
        }
    }

    const procs = ast.body.filter((s: any) => s.type === 'ProcedureDeclaration');
    const definedProcs = new Map<string, { proc: any; kind: string; scope: string }>();
    const callsByProc = new Map<string, Set<string>>();

    const procedures: ProcedureMetrics[] = procs.map((proc: any) => {
        const startLine = proc.loc?.start.line ?? -1;
        const endLine = proc.loc?.end.line ?? -1;
        const nestState = { max: 0 };
        measureNestDepth(proc, 0, nestState);

        const kind: 'Sub' | 'Function' | 'Property' =
            proc.isProperty ? 'Property' : proc.isFunction ? 'Function' : 'Sub';

        // VBA のスコープデフォルトは Public
        const scope = (proc.scope ?? 'public') as 'public' | 'private' | 'friend';

        const excel = findExcelAccess(proc.body);
        const numericLits = findRepeatedNumericLiterals(proc.body);
        const procName = proc.name?.name ?? '<anonymous>';

        definedProcs.set(procName.toLowerCase(), { proc, kind, scope });
        callsByProc.set(procName.toLowerCase(), collectProcedureCalls(proc.body, definedTypes));

        return {
            name: procName,
            kind,
            scope,
            startLine,
            endLine,
            lineCount: endLine - startLine + 1,
            maxNestDepth: nestState.max,
            localDeclCount: countLocalDeclarations(proc.body),
            assignmentBlocks: findConsecutiveAssignmentBlocks(proc.body),
            excelAccessCount: excel.count,
            excelAccessSamples: excel.samples,
            excelObjectsUsed: excel.objectsUsed,
            repeatedNumericLiterals: numericLits.slice(0, 10),
            referenceCount: 0,  // 後段のワークスペース解析で埋める
        };
    });

    const prefixClusters = detectPrefixClusters(ast.body);

    return {
        report: { filePath, totalLines: lines.length, procedureCount: procedures.length, procedures, prefixClusters, warnings },
        definedProcs,
        definedTypes,
        callsByProc,
    };
}

// ワークスペース全体での参照カウント / エントリーポイント候補 / モック必要箇所 を集計
function buildWorkspaceReport(analyses: FileAnalysis[]): WorkspaceReport {
    // 全プロシージャ名（lowercase）→ 定義ファイル
    const procDefs = new Map<string, { file: string; name: string; kind: string; scope: string }>();
    for (const a of analyses) {
        for (const [lname, info] of a.definedProcs) {
            // 同名定義がある場合は最初を採用（後段で警告するかは将来課題）
            if (!procDefs.has(lname)) {
                procDefs.set(lname, { file: a.report.filePath, name: info.proc.name?.name ?? lname, kind: info.kind, scope: info.scope });
            }
        }
    }

    // 全呼び出しを集計
    const refCounts = new Map<string, number>();
    for (const a of analyses) {
        for (const calls of a.callsByProc.values()) {
            for (const callee of calls) {
                if (procDefs.has(callee)) {
                    refCounts.set(callee, (refCounts.get(callee) ?? 0) + 1);
                }
            }
        }
    }

    // 各 FileReport の referenceCount を埋める
    for (const a of analyses) {
        for (const p of a.report.procedures) {
            p.referenceCount = refCounts.get(p.name.toLowerCase()) ?? 0;
        }
    }

    // エントリーポイント候補 / Dead code 候補 を抽出
    const entryPointCandidates: WorkspaceReport['entryPointCandidates'] = [];
    const deadCodeCandidates: WorkspaceReport['deadCodeCandidates'] = [];
    for (const a of analyses) {
        for (const p of a.report.procedures) {
            if (p.referenceCount === 0) {
                if (p.scope === 'private') {
                    deadCodeCandidates.push({
                        file: a.report.filePath,
                        name: p.name,
                        kind: p.kind,
                        line: p.startLine,
                    });
                } else {
                    // public / friend / 暗黙のpublic
                    let reason = 'Public Sub/Function with 0 internal references';
                    // ヒューリスティック: Workbook_Open など特定の名前はイベントハンドラ
                    if (/^(Workbook|Worksheet|Auto)_/i.test(p.name)) {
                        reason = 'Excel event handler / auto macro';
                    } else if (/^Test_/i.test(p.name)) {
                        reason = 'Test procedure';
                    } else if (p.kind === 'Sub' && p.lineCount > 5) {
                        reason = 'Public Sub - likely button/macro entry point';
                    }
                    entryPointCandidates.push({
                        file: a.report.filePath,
                        name: p.name,
                        kind: p.kind,
                        line: p.startLine,
                        reason,
                    });
                }
            }
        }
    }

    // Excel モック必要箇所
    const excelMockTargets: WorkspaceReport['excelMockTargets'] = [];
    for (const a of analyses) {
        for (const p of a.report.procedures) {
            if (p.excelObjectsUsed.length > 0) {
                excelMockTargets.push({
                    file: a.report.filePath,
                    procName: p.name,
                    objects: p.excelObjectsUsed,
                    sampleAccesses: p.excelAccessSamples.slice(0, 3).map(s => s.expr),
                });
            }
        }
    }

    // コールグラフ構築
    const callGraph: CallGraphEdge[] = [];
    for (const a of analyses) {
        for (const [callerLower, callees] of a.callsByProc) {
            const callerDef = a.definedProcs.get(callerLower);
            if (!callerDef) continue;
            const fromName = callerDef.proc.name?.name ?? callerLower;
            for (const callee of callees) {
                const calleeDef = procDefs.get(callee);
                if (!calleeDef) continue;  // 外部/組み込み関数は除外
                callGraph.push({
                    from: fromName,
                    fromFile: a.report.filePath,
                    to: calleeDef.name,
                    toFile: calleeDef.file,
                });
            }
        }
    }

    return {
        files: analyses.map(a => a.report),
        entryPointCandidates,
        deadCodeCandidates,
        excelMockTargets,
        callGraph,
    };
}

// ---------------------------------------------------------
// テキストフォーマッタ
// ---------------------------------------------------------

function formatFileReport(r: FileReport): string {
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
        const scopeLabel = p.scope === 'public' ? '' : `[${p.scope}] `;
        out.push(`  ${scopeLabel}[${p.kind}] ${p.name}  (${p.lineCount}行, L${p.startLine}-L${p.endLine}, refs=${p.referenceCount})`);
        const flags: string[] = [];
        if (p.lineCount >= 100) flags.push(`📏 LARGE (${p.lineCount}行)`);
        if (p.maxNestDepth >= 5) flags.push(`🌀 DEEP_NEST (${p.maxNestDepth}段)`);
        if (p.localDeclCount >= 30) flags.push(`🧮 MANY_LOCALS (${p.localDeclCount}個)`);
        if (p.excelAccessCount >= 10) flags.push(`📊 EXCEL_HEAVY (${p.excelAccessCount}件)`);
        if (flags.length) out.push(`    フラグ: ${flags.join(' / ')}`);
        else out.push(`    最大ネスト ${p.maxNestDepth} / Dim ${p.localDeclCount} / Excel ${p.excelAccessCount}`);

        if (p.excelObjectsUsed.length > 0) {
            out.push(`    🧪 モック必要候補: ${p.excelObjectsUsed.join(', ')}`);
        }
        if (p.assignmentBlocks.length) {
            out.push(`    連続代入ブロック（抽出候補）:`);
            for (const b of p.assignmentBlocks) {
                out.push(`      L${b.startLine}-L${b.endLine}: ${b.count}件 [shape:${b.shape}] [${b.context}]`);
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

    if (r.prefixClusters.length > 0) {
        out.push('');
        out.push('  📦 接頭辞クラスタ（UDT/Enum 抽出候補）:');
        for (const c of r.prefixClusters) {
            out.push(`    ${c.prefix} × ${c.members.length}件  → ${c.suggestion}`);
            out.push(`      例: ${c.members.slice(0, 4).join(', ')}${c.members.length > 4 ? ' ...' : ''}`);
        }
    }

    return out.join('\n');
}

function formatWorkspaceSummary(w: WorkspaceReport): string {
    const out: string[] = [];

    if (w.entryPointCandidates.length > 0) {
        out.push('');
        out.push('========================================');
        out.push('🚪 エントリーポイント候補（Public・参照0）');
        out.push('========================================');
        out.push('  以下は他のVBAコードから呼ばれていない Public プロシージャです。');
        out.push('  Excel のボタン・イベントハンドラ・Application.Run 等から呼ばれている可能性が高く、');
        out.push('  そのまま「マクロの入口」として扱うべき候補です（安易に削除しないこと）。');
        out.push('');
        for (const e of w.entryPointCandidates) {
            const fname = path.basename(e.file);
            out.push(`  • [${e.kind}] ${e.name}  (${fname}:L${e.line})`);
            out.push(`      → ${e.reason}`);
        }
    }

    if (w.deadCodeCandidates.length > 0) {
        out.push('');
        out.push('========================================');
        out.push('💀 Dead code 候補（Private・参照0）');
        out.push('========================================');
        out.push('  Private 宣言で他から参照されていない関数。削除候補ですが、');
        out.push('  Application.Run による動的呼び出しの可能性があるため最終判断は人間が行うこと。');
        out.push('');
        for (const d of w.deadCodeCandidates) {
            const fname = path.basename(d.file);
            out.push(`  • [${d.kind}] ${d.name}  (${fname}:L${d.line})`);
        }
    }

    if (w.excelMockTargets.length > 0) {
        out.push('');
        out.push('========================================');
        out.push('🧪 Excel モック必要箇所一覧');
        out.push('========================================');
        out.push('  リファクタリングで純粋ロジックを抽出する際、');
        out.push('  以下のプロシージャは Excel オブジェクトのモック化（または依存分離）が必要です。');
        out.push('');

        // モック対象オブジェクト別に集計
        const byObject = new Map<string, Array<{ file: string; procName: string }>>();
        for (const t of w.excelMockTargets) {
            for (const obj of t.objects) {
                if (!byObject.has(obj)) byObject.set(obj, []);
                byObject.get(obj)!.push({ file: t.file, procName: t.procName });
            }
        }
        out.push('  ◆ オブジェクト別の利用箇所:');
        const sortedObjs = [...byObject.entries()].sort((a, b) => b[1].length - a[1].length);
        for (const [obj, users] of sortedObjs) {
            out.push(`    ${obj}: ${users.length}プロシージャ`);
            for (const u of users.slice(0, 5)) {
                out.push(`      - ${u.procName} (${path.basename(u.file)})`);
            }
            if (users.length > 5) out.push(`      ... 他 ${users.length - 5} 件`);
        }

        out.push('');
        out.push('  ◆ プロシージャ別のモック必要オブジェクト:');
        for (const t of w.excelMockTargets) {
            out.push(`    ${t.procName} (${path.basename(t.file)})`);
            out.push(`      要モック: ${t.objects.join(', ')}`);
            if (t.sampleAccesses.length) {
                out.push(`      使用例: ${t.sampleAccesses.join(' / ')}`);
            }
        }
    }

    if (w.callGraph.length > 0) {
        out.push('');
        out.push('========================================');
        out.push('📞 コールグラフ（呼び出し関係）');
        out.push('========================================');
        // Group by caller
        const byCaller = new Map<string, string[]>();
        for (const e of w.callGraph) {
            const key = `${e.from} (${path.basename(e.fromFile)})`;
            if (!byCaller.has(key)) byCaller.set(key, []);
            byCaller.get(key)!.push(e.to);
        }
        for (const [caller, callees] of byCaller) {
            out.push(`  ${caller}`);
            for (const callee of callees) out.push(`    → ${callee}`);
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

// Workspace Outline: compact summary suitable for pasting into AI context
function formatWorkspaceOutline(workspace: WorkspaceReport): string {
    const out: string[] = [];
    const refCounts = new Map<string, number>();
    for (const f of workspace.files) {
        for (const p of f.procedures) {
            refCounts.set(p.name.toLowerCase(), p.referenceCount);
        }
    }
    for (const f of workspace.files) {
        if (f.procedures.length === 0) continue;
        out.push(`[${path.basename(f.filePath, path.extname(f.filePath))}]  (${f.totalLines}L)`);
        for (const p of f.procedures) {
            const scopePrefix = p.scope === 'private' ? 'Private ' : '';
            const kindLabel = p.kind === 'Function' ? 'Function' : p.kind === 'Property' ? 'Property' : 'Sub';
            const refs = p.referenceCount === 0 ? '  ← 0 refs' : '';
            const flags: string[] = [];
            if (p.lineCount >= 100) flags.push(`${p.lineCount}L`);
            if (p.maxNestDepth >= 5) flags.push(`nest=${p.maxNestDepth}`);
            if (p.excelAccessCount > 0) flags.push(`Excel×${p.excelAccessCount}`);
            const flagStr = flags.length ? `  [${flags.join(', ')}]` : '';
            out.push(`  ${scopePrefix}${kindLabel} ${p.name}${flagStr}${refs}`);
        }
    }
    return out.join('\n');
}

function main() {
    const args = process.argv.slice(2);
    if (args.length === 0 || args.includes('--help')) {
        console.log('Usage: node vba-analyzer.cjs <file-or-dir> [options]');
        console.log('');
        console.log('  --json          JSON 形式で出力（プログラム連携用）');
        console.log('  --summary-only  ワークスペース要約（エントリーポイント候補・モック必要箇所）のみ表示');
        console.log('  --outline       Workspace Outline（AI向けコンテキスト圧縮）を出力');
        process.exit(args.length === 0 ? 1 : 0);
    }
    const wantJson = args.includes('--json');
    const summaryOnly = args.includes('--summary-only');
    const wantOutline = args.includes('--outline');
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

    const analyses = files.map(analyzeFile);
    const workspace = buildWorkspaceReport(analyses);

    if (wantJson) {
        console.log(JSON.stringify(workspace, null, 2));
        return;
    }

    if (wantOutline) {
        console.log(formatWorkspaceOutline(workspace));
        return;
    }

    if (!summaryOnly) {
        for (const r of workspace.files) console.log(formatFileReport(r));
    }
    console.log(formatWorkspaceSummary(workspace));
}

main();
