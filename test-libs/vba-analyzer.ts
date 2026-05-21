// VBA Code Analyzer
//
// レガシーVBAのリファクタリング支援を目的とした静的解析ツール。
// 単一ファイルまたはディレクトリを引数に取り、リファクタリング候補となる
// 問題箇所（巨大プロシージャ・深いネスト・連続代入・Excel依存・繰り返し数値リテラル等）を列挙する。
//
// Usage:
//   ./node_modules/.bin/esbuild test-libs/vba-analyzer.ts --bundle --outfile=test-libs/vba-analyzer.cjs --platform=node
//   node test-libs/vba-analyzer.cjs <path>                        # テキスト形式（全項目）
//   node test-libs/vba-analyzer.cjs <path> --json                 # JSON 形式
//   node test-libs/vba-analyzer.cjs <path> --summary-only         # エントリーポイント・モック・重複のみ
//   node test-libs/vba-analyzer.cjs <path> --outline              # AI向けコンパクト要約
//   node test-libs/vba-analyzer.cjs <path> --commented-code       # コメントアウトコード候補のみ
//
// オプション別の用途:
//   --outline         ファイル全体をAIに渡す前の「どのプロシージャが問題か」の絞り込みに使う。
//                     出力が小さいため、AI への読み込みトークンを節約できる。
//   --commented-code  コメントアウトされた死コードの確認に特化。出力をそのまま AI に貼り付けて
//                     「これはコードか説明文か」を判断させることを想定している。
//                     ファイル全体を AI に渡す前にこのフラグで先に絞り込むと読み込みトークンを削減できる。
//   --summary-only    エントリーポイント候補・Dead code・Excel モック・重複ブロックに集中したいとき。
//
// 既知の制約（TODO_NEXT.md「実証実験で判明したギャップ」を参照）:
//   - 連続代入の「形状クラスタリング」は未実装（件数のみ）
//   - データフロー解析（Def-Use）は未実装
//   - 接頭辞クラスタによる UDT 抽出候補検出は未実装
//   - --commented-code のスコアリングはヒューリスティック。自然言語の仕様説明も低スコアで検出される場合がある

import { Lexer } from '../src/engine/lexer';
import { Parser } from '../src/engine/parser';
import { preprocess } from '../src/engine/preprocessor';
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
    magicLiteralsInCalls: Array<{ callee: string; argIndex: number; value: string | number; lines: number[] }>;
    byRefAssignments: Array<{ paramName: string; paramType: string | null; lines: number[] }>;
    parameters: Array<{ name: string; paramType: string | null; isByVal: boolean }>;
    returnType: string | null;            // Function の戻り値型（Sub/Property は null）
    ioSideEffectCount: number;            // MsgBox / InputBox / Debug.Print の呼び出し回数
    referenceCount: number;               // 他のプロシージャから呼ばれている回数（クロスファイル含む）
    // 凝集度指標
    hardcodedSheetCount: number;          // Sheets("名前")/Worksheets("名前") の固定シート名の種類数
    hardcodedAddressCount: number;        // Range("A1")/Cells(1,1) の固定アドレス引数の種類数
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

interface DuplicateOccurrence {
    file: string;
    procName: string;
    startLine: number;
    endLine: number;
}

interface DuplicateBlock {
    length: number;             // 重複ブロックの行数（statement数ではなくソース行数）
    stmtCount: number;          // statement数
    occurrences: DuplicateOccurrence[];
    shape: string;              // 正規化後のパターン（先頭1行のみ表示）
}

interface CommentedCodeBlock {
    file: string;
    startLine: number;   // 1-based
    endLine: number;     // 1-based
    lineCount: number;
    score: number;
    confidence: 'high' | 'medium' | 'low';
    detectedKeywords: Record<string, number>;
    strippedContent: string;  // ' を除去したテキスト（AI評価用）
    rawContent: string;       // ' を保持したテキスト（表示用）
    contextBefore: string[];  // ブロック直前の数行（元のソース行）
    contextAfter: string[];   // ブロック直後の数行（元のソース行）
}

interface WorkspaceReport {
    files: FileReport[];
    entryPointCandidates: Array<{ file: string; name: string; kind: string; line: number; reason: string }>;
    deadCodeCandidates: Array<{ file: string; name: string; kind: string; line: number }>;
    excelMockTargets: Array<{ file: string; procName: string; objects: string[]; sampleAccesses: string[] }>;
    callGraph: CallGraphEdge[];
    duplicateBlocks: DuplicateBlock[];
    commentedCodeBlocks: CommentedCodeBlock[];
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
    // alternate が IfStatement (ElseIf の再帰表現) の場合は同じ depth で再帰する
    if (node.alternate && !Array.isArray(node.alternate) && node.alternate.type === 'IfStatement') {
        measureNestDepth(node.alternate, depth, current);
    } else {
        walk(node.alternate, depth);
    }
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

// ---------------------------------------------------------
// Range/Cells/Sheets 引数の即値検出
// ---------------------------------------------------------

const RANGE_CELLS_SHEETS = new Set(['range', 'cells', 'sheets', 'worksheets']);

// callee名（小文字）→ 位置引数の名前テーブル。不明な位置は undefined（argN にフォールバック）
const CALLEE_PARAM_NAMES: Record<string, string[]> = {
    cells:          ['RowIndex', 'ColumnIndex'],
    'cells.item':   ['RowIndex', 'ColumnIndex'],
    'cells()':      ['RowIndex', 'ColumnIndex'],
    range:          ['Cell1', 'Cell2'],
    'range.item':   ['RowIndex', 'ColumnIndex'],
    'range()':      ['RowIndex', 'ColumnIndex'],
    sheets:         ['Index'],
    'sheets.item':  ['Index'],
    worksheets:     ['Index'],
    'worksheets.item': ['Index'],
};

export function paramName(callee: string, argIndex: number): string {
    const names = CALLEE_PARAM_NAMES[callee.toLowerCase()];
    return names?.[argIndex] ?? `arg${argIndex + 1}`;
}

export function findByRefAssignments(
    proc: any,
): Array<{ paramName: string; paramType: string | null; lines: number[] }> {
    const byRefParams = new Map<string, string | null>();
    for (const p of proc.parameters ?? []) {
        if (!p.isByVal) {
            byRefParams.set((p.name as string).toLowerCase(), p.paramType ?? null);
        }
    }
    if (byRefParams.size === 0) return [];

    const lines = new Map<string, number[]>();

    function visit(node: any): void {
        if (!node || typeof node !== 'object') return;
        if (node.type === 'AssignmentStatement') {
            const lhs = node.left;
            const name = lhs?.type === 'Identifier' ? (lhs.name as string).toLowerCase() : null;
            if (name && byRefParams.has(name)) {
                if (!lines.has(name)) lines.set(name, []);
                lines.get(name)!.push(lhs.loc?.start?.line ?? node.loc?.start?.line ?? -1);
            }
        }
        for (const key of Object.keys(node)) {
            if (key === 'loc') continue;
            const v = node[key];
            if (Array.isArray(v)) v.forEach(visit);
            else if (v && typeof v === 'object' && v.type) visit(v);
        }
    }

    for (const s of proc.body ?? []) visit(s);

    return [...lines.entries()].map(([name, ls]) => ({
        paramName: name,
        paramType: byRefParams.get(name) ?? null,
        lines: ls,
    }));
}

export function findMagicLiteralsInCalls(
    stmts: any,
): Array<{ callee: string; argIndex: number; value: string | number; lines: number[] }> {
    const raw: Array<{ line: number; callee: string; argIndex: number; value: string | number }> = [];

    function getCalleeName(callee: any): string | null {
        if (!callee) return null;
        if (callee.type === 'Identifier') return callee.name;
        if (callee.type === 'MemberExpression') {
            const prop = callee.property;
            return prop?.name ?? (typeof prop === 'string' ? prop : null);
        }
        return null;
    }

    // X / X.prop / X(...) の「末端オブジェクト名」を返す（.Item 検出用）
    function getObjectName(node: any): string | null {
        if (!node) return null;
        if (node.type === 'Identifier') return node.name;
        if (node.type === 'MemberExpression') {
            const prop = node.property;
            return prop?.name ?? (typeof prop === 'string' ? prop : null);
        }
        if (node.type === 'CallExpression') return getCalleeName(node.callee);
        return null;
    }

    function pushLiterals(callNode: any, label: string): void {
        const args: any[] = callNode.args ?? [];
        args.forEach((arg, i) => {
            if (arg.type === 'NumberLiteral' || arg.type === 'StringLiteral') {
                raw.push({
                    line: callNode.callee?.loc?.start?.line ?? arg.loc?.start?.line ?? -1,
                    callee: label,
                    argIndex: i,
                    value: arg.value,
                });
            }
        });
    }

    function visit(node: any): void {
        if (!node || typeof node !== 'object') return;
        if (node.type === 'CallExpression') {
            const callee = node.callee;
            const calleeName = getCalleeName(callee);

            if (calleeName && RANGE_CELLS_SHEETS.has(calleeName.toLowerCase())) {
                // Cells(3,5) / ws.Cells(3,5) / Worksheets("X") / Range("A1")
                pushLiterals(node, calleeName);
            } else if (callee?.type === 'MemberExpression' && calleeName?.toLowerCase() === 'item') {
                // Cells.Item(3,5) / ws.Cells.Item(3,5) / Range("A1:B3").Item(3,5)
                const objName = getObjectName(callee.object);
                if (objName && RANGE_CELLS_SHEETS.has(objName.toLowerCase())) {
                    pushLiterals(node, objName + '.Item');
                }
            } else if (callee?.type === 'CallExpression') {
                // Range("A1:B3")(3,5) — Range 結果への直接インデックス
                const innerName = getCalleeName(callee.callee);
                if (innerName && RANGE_CELLS_SHEETS.has(innerName.toLowerCase())) {
                    pushLiterals(node, innerName + '()');
                }
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

    // (callee, argIndex, value) でグループ化して出現行リストを返す
    const buckets = new Map<string, { callee: string; argIndex: number; value: string | number; lines: number[] }>();
    for (const r of raw) {
        const key = `${r.callee}:${r.argIndex}:${r.value}`;
        if (!buckets.has(key)) buckets.set(key, { callee: r.callee, argIndex: r.argIndex, value: r.value, lines: [] });
        buckets.get(key)!.lines.push(r.line);
    }
    return [...buckets.values()].sort((a, b) => a.lines[0] - b.lines[0]);
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
// 重複ブロック検出（クローン検出）
// ---------------------------------------------------------

// Expression を正規化した文字列にシリアライズ。
// - Identifier → $ID（変数名の違いを無視）
// - NumberLiteral → 実値を保持（3 と 5 は区別）
// - StringLiteral → $STR
// - MemberExpression のプロパティ名は保持（.Value と .Name は区別）
function serExpr(node: any): string {
    if (!node) return '?';
    switch (node.type) {
        case 'Identifier':   return '$ID';
        case 'NumberLiteral': return String(node.value);
        case 'StringLiteral': return '$STR';
        case 'DateLiteral':  return '$DATE';
        case 'MemberExpression':
            return `${serExpr(node.object)}.${node.property?.name ?? '?'}`;
        case 'DictionaryAccessExpression':
            return `${serExpr(node.object)}!${node.property?.name ?? '?'}`;
        case 'CallExpression':
            return `${serExpr(node.callee)}(${(node.args ?? []).map(serExpr).join(',')})`;
        case 'BinaryExpression':
            return `(${serExpr(node.left)}${node.operator}${serExpr(node.right)})`;
        case 'UnaryExpression':
            return `${node.operator}${serExpr(node.argument)}`;
        case 'NewExpression':
            return `New $ID`;
        default:
            return node.type ?? '?';
    }
}

// Statement を正規化した1行文字列にシリアライズ。
// 制御構造はボディを含まず「型+条件」だけにする（ボディは別途シーケンスとして扱う）。
function serStmt(s: any): string {
    if (!s) return '?';
    switch (s.type) {
        case 'AssignmentStatement': return `${serExpr(s.left)}=${serExpr(s.right)}`;
        case 'SetStatement':        return `Set ${serExpr(s.left)}=${serExpr(s.right)}`;
        case 'CallStatement':       return `Call ${serExpr(s.expression)}`;
        case 'VariableDeclaration':
            return `Dim[${(s.declarations ?? []).map((d: any) => d.objectType ?? '').join(',')}]`;
        case 'ConstDeclaration':    return `Const=${serExpr(s.value)}`;
        case 'IfStatement':         return `If(${serExpr(s.condition)})`;
        case 'ForStatement':        return `For $ID=${serExpr(s.start)} To ${serExpr(s.end)}`;
        case 'ForEachStatement':    return `ForEach $ID In ${serExpr(s.collection)}`;
        case 'WhileStatement':      return `While(${serExpr(s.condition)})`;
        case 'DoWhileStatement':    return `Do${s.conditionType ?? ''}(${serExpr(s.condition)})`;
        case 'SelectCaseStatement': return `Select ${serExpr(s.expression)}`;
        case 'WithStatement':       return `With ${serExpr(s.expression)}`;
        case 'OnErrorStatement':    return `OnError`;
        case 'ExitStatement':       return `Exit ${s.kind ?? ''}`;
        case 'GoToStatement':       return `GoTo $ID`;
        case 'ReturnStatement':     return `Return`;
        default:                    return s.type ?? '?';
    }
}

// すべてのフラットなステートメントシーケンスを収集する（制御構造ボディも再帰的に含む）
interface StmtSeq {
    stmts: any[];
    file: string;
    procName: string;
}

function collectStmtSeqs(stmts: any[], file: string, procName: string): StmtSeq[] {
    if (!Array.isArray(stmts) || stmts.length === 0) return [];
    const result: StmtSeq[] = [{ stmts, file, procName }];
    for (const s of stmts) {
        if (Array.isArray(s.consequent))  result.push(...collectStmtSeqs(s.consequent, file, procName));
        if (Array.isArray(s.alternate))   result.push(...collectStmtSeqs(s.alternate, file, procName));
        if (Array.isArray(s.body))        result.push(...collectStmtSeqs(s.body, file, procName));
        if (s.elseIfClauses) for (const c of s.elseIfClauses)
            result.push(...collectStmtSeqs(c.consequent ?? [], file, procName));
        if (s.cases) for (const c of s.cases)
            result.push(...collectStmtSeqs(c.body ?? [], file, procName));
    }
    return result;
}

function detectDuplicateBlocks(
    analyses: FileAnalysis[],
    minStmts = 3,
    minCount = 2,
    maxStmts = 12,
): DuplicateBlock[] {
    // N-gram → 出現箇所リスト
    const index = new Map<string, DuplicateOccurrence[]>();

    for (const a of analyses) {
        for (const [procNameLower, info] of a.definedProcs) {
            const procName = info.proc.name?.name ?? procNameLower;
            const seqs = collectStmtSeqs(info.proc.body ?? [], a.report.filePath, procName);
            for (const { stmts } of seqs) {
                const serialized = stmts.map(serStmt);
                for (let len = minStmts; len <= Math.min(maxStmts, stmts.length); len++) {
                    for (let i = 0; i <= stmts.length - len; i++) {
                        const key = serialized.slice(i, i + len).join('\n');
                        const startLine = stmts[i]?.loc?.start.line ?? -1;
                        const endLine   = stmts[i + len - 1]?.loc?.end.line ?? -1;
                        if (startLine === -1) continue;
                        if (!index.has(key)) index.set(key, []);
                        index.get(key)!.push({ file: a.report.filePath, procName, startLine, endLine });
                    }
                }
            }
        }
    }

    // 同一 proc 内でオーバーラップする出現を除去（スライディングウィンドウのノイズ対策）
    // 同一 file+proc で startLine が前の endLine 以下のものは除く（先着優先）
    function deduplicateOverlapping(occs: DuplicateOccurrence[]): DuplicateOccurrence[] {
        const sorted = [...occs].sort((a, b) =>
            a.file.localeCompare(b.file) || a.procName.localeCompare(b.procName) || a.startLine - b.startLine
        );
        const lastEnd = new Map<string, number>();
        return sorted.filter(occ => {
            const k = `${occ.file}::${occ.procName}`;
            const prev = lastEnd.get(k) ?? -Infinity;
            if (occ.startLine > prev) {
                lastEnd.set(k, occ.endLine);
                return true;
            }
            return false;
        });
    }

    // minCount 以上の N-gram を収集（dedup 後）
    const candidates: Array<{ key: string; occurrences: DuplicateOccurrence[]; stmtCount: number }> = [];
    for (const [key, occs] of index) {
        const deduped = deduplicateOverlapping(occs);
        if (deduped.length < minCount) continue;
        const stmtCount = key.split('\n').length;
        candidates.push({ key, occurrences: deduped, stmtCount });
    }

    // 長いマッチの部分集合になっている短いマッチを除去（maximal match のみ残す）
    // 同じ出現箇所セットを持つ短いパターンはスキップ
    candidates.sort((a, b) => b.stmtCount - a.stmtCount);
    const kept: typeof candidates = [];
    const coveredKeys = new Set<string>();
    for (const c of candidates) {
        // この候補の各出現が、より長い候補に完全に包含されているか確認
        const allCovered = c.occurrences.every(occ => {
            // 同一ファイル・プロシージャ・行範囲を包含する longer match が存在するか
            return kept.some(longer =>
                longer.stmtCount > c.stmtCount &&
                longer.occurrences.some(lo =>
                    lo.file === occ.file &&
                    lo.procName === occ.procName &&
                    lo.startLine <= occ.startLine &&
                    lo.endLine >= occ.endLine
                )
            );
        });
        if (!allCovered && !coveredKeys.has(c.key)) {
            kept.push(c);
            coveredKeys.add(c.key);
        }
    }

    // DuplicateBlock に変換、インパクト順（stmtCount × occurrences）でソート
    return kept
        .map(c => ({
            stmtCount: c.stmtCount,
            length: Math.max(...c.occurrences.map(o => o.endLine - o.startLine + 1)),
            occurrences: c.occurrences,
            shape: c.key.split('\n')[0],  // 先頭ステートメントのパターンを代表として表示
        }))
        .sort((a, b) => (b.stmtCount * b.occurrences.length) - (a.stmtCount * a.occurrences.length))
        .slice(0, 30);  // 上位30件
}

// ---------------------------------------------------------
// Excel/VBA 定数抽出と const.ts 出力
// ---------------------------------------------------------

const XL_VB_PATTERN = /^(xl[A-Z]|vb[A-Z]|mso[A-Z])/;

const KNOWN_XL_CONSTANTS: Record<string, number | string> = {
    // Direction
    xlUp: -4162, xlDown: -4121, xlToLeft: -4159, xlToRight: -4161,
    // PasteType
    xlPasteAll: -4104, xlPasteValues: -4163, xlPasteFormulas: -4123,
    xlPasteFormats: -4122, xlPasteComments: -4144, xlPasteAllExceptBorders: 7,
    xlPasteColumnWidths: 8, xlPasteFormulasAndNumberFormats: 11,
    xlPasteValuesAndNumberFormats: 12, xlPasteValidation: 6,
    // General
    xlValues: -4163, xlFormulas: -4123, xlComments: -4144, xlNone: -4142,
    xlAutomatic: -4105, xlManual: -4135,
    xlFirst: 2, xlLast: 1, xlNext: 1, xlPrevious: 2,
    xlYes: 1, xlNo: 2, xlGuess: 0,
    xlWhole: 1, xlPart: 2,
    xlByRows: 1, xlByColumns: 2,
    xlAscending: 1, xlDescending: 2,
    xlOr: 2, xlAnd: 1,
    // CellType
    xlCellTypeBlanks: 4, xlCellTypeComments: -4144, xlCellTypeConstants: 2,
    xlCellTypeFormulas: -4123, xlCellTypeLastCell: 11, xlCellTypeVisible: 12,
    // MsgBox buttons
    vbOKOnly: 0, vbOKCancel: 1, vbAbortRetryIgnore: 2,
    vbYesNoCancel: 3, vbYesNo: 4, vbRetryCancel: 5,
    // MsgBox icons
    vbCritical: 16, vbQuestion: 32, vbExclamation: 48, vbInformation: 64,
    // MsgBox default button / modal
    vbDefaultButton1: 0, vbDefaultButton2: 256, vbDefaultButton3: 512,
    vbApplicationModal: 0, vbSystemModal: 4096,
    // MsgBox return values
    vbOK: 1, vbCancel: 2, vbAbort: 3, vbRetry: 4, vbIgnore: 5, vbYes: 6, vbNo: 7,
    // Boolean / VarType
    vbTrue: -1, vbFalse: 0,
    vbEmpty: 0, vbNull: 1, vbInteger: 2, vbLong: 3, vbSingle: 4, vbDouble: 5,
    vbCurrency: 6, vbDate: 7, vbString: 8, vbObject: 9, vbError: 10,
    vbBoolean: 11, vbVariant: 12, vbDataObject: 13, vbDecimal: 14, vbByte: 17,
    vbArray: 8192,
    // String constants
    vbCrLf: '\r\n', vbCr: '\r', vbLf: '\n', vbTab: '\t',
    vbNullChar: '\x00', vbNewLine: '\n', vbBack: '\b',
    // CompareMethod
    vbBinaryCompare: 0, vbTextCompare: 1, vbDatabaseCompare: 2,
    // DayOfWeek
    vbSunday: 1, vbMonday: 2, vbTuesday: 3, vbWednesday: 4,
    vbThursday: 5, vbFriday: 6, vbSaturday: 7, vbUseSystemDayOfWeek: 0,
    // FirstWeekOfYear
    vbUseSystem: 0, vbFirstJan1: 1, vbFirstFourDays: 2, vbFirstFullWeek: 3,
    // FileAttribute
    vbNormal: 0, vbReadOnly: 1, vbHidden: 2, vbSystem: 4,
    vbVolume: 8, vbDirectory: 16, vbArchive: 32, vbAlias: 64,
};

function collectExcelConstantRefs(ast: any): Set<string> {
    const refs = new Set<string>();
    function visit(node: any): void {
        if (!node || typeof node !== 'object') return;
        if (node.type === 'Identifier' && typeof node.name === 'string') {
            if (XL_VB_PATTERN.test(node.name)) refs.add(node.name);
        }
        for (const key of Object.keys(node)) {
            if (key === 'loc') continue;
            const v = node[key];
            if (Array.isArray(v)) v.forEach(visit);
            else if (v && typeof v === 'object' && v.type) visit(v);
        }
    }
    for (const s of ast.body) visit(s);
    return refs;
}

export function findIoSideEffects(body: any[]): number {
    const IO_NAMES = new Set(['msgbox', 'inputbox']);
    let count = 0;
    function visit(node: any): void {
        if (!node || typeof node !== 'object') return;
        if (node.type === 'CallExpression') {
            const callee = node.callee;
            if (callee?.type === 'Identifier' && IO_NAMES.has(callee.name.toLowerCase())) {
                count++;
            }
            if (callee?.type === 'MemberExpression') {
                const obj = callee.object?.name?.toLowerCase();
                const prop = (callee.property?.name ?? callee.property)?.toString().toLowerCase();
                if (obj === 'debug' && prop === 'print') count++;
            }
        }
        for (const key of Object.keys(node)) {
            if (key === 'loc') continue;
            const v = node[key];
            if (Array.isArray(v)) v.forEach(visit);
            else if (v && typeof v === 'object' && v.type) visit(v);
        }
    }
    if (isStatementArray(body)) for (const s of body) visit(s);
    return count;
}

export function vbaTypeToTs(vbaType: string | null): string {
    if (!vbaType) return 'any';
    switch (vbaType.toLowerCase()) {
        case 'long': case 'integer': case 'byte': case 'single': case 'double': case 'currency': return 'number';
        case 'string': return 'string';
        case 'boolean': return 'boolean';
        default: return 'any';
    }
}

function emitTestTemplates(analyses: FileAnalysis[], outputDir: string): void {
    fs.mkdirSync(outputDir, { recursive: true });
    let fileCount = 0;

    for (const a of analyses) {
        const funcs = a.report.procedures.filter(p => p.kind === 'Function');
        if (funcs.length === 0) continue;

        const vbaPath = a.report.filePath;
        const baseName = path.basename(vbaPath, path.extname(vbaPath));
        const relVbaPath = path.relative(outputDir, vbaPath).replace(/\\/g, '/');
        const relTestLibPath = path.relative(outputDir, path.join(__dirname, 'test-runner')).replace(/\\/g, '/');

        const lines: string[] = [
            '// Auto-generated by vba-analyzer --gen-test-dir',
            `// Source: ${vbaPath}`,
            `import { VBARunner } from '${relTestLibPath}';`,
            '',
            `const runner = new VBARunner('${relVbaPath}');`,
            '',
        ];

        for (const p of funcs) {
            const isPure = p.excelObjectsUsed.length === 0 && p.ioSideEffectCount === 0;
            const testFn = isPure ? 'test' : 'test.skip';

            const reasons: string[] = [];
            if (p.excelObjectsUsed.length > 0) reasons.push(`Excel access (${p.excelObjectsUsed.join(', ')})`);
            if (p.ioSideEffectCount > 0) reasons.push(`I/O side effects (${p.ioSideEffectCount}件)`);

            const argList = p.parameters
                .map(pm => `/* ${pm.name}: ${vbaTypeToTs(pm.paramType)} */`)
                .join(', ');

            if (!isPure) {
                lines.push(`// ⚠️  ${p.name}: impure — ${reasons.join(', ')}`);
            }
            lines.push(`describe('${p.name}', () => {`);
            const retTs = vbaTypeToTs(p.returnType);
            lines.push(`  ${testFn}('基本動作', () => {`);
            lines.push(`    const result = runner.run('${p.name}', [${argList}]) as ${retTs};`);
            lines.push(`    expect(result).toBe(/* expected: ${retTs} */);`);
            lines.push(`  });`);
            lines.push(`});`);
            lines.push('');
        }

        const outPath = path.join(outputDir, `${baseName}.test.ts`);
        fs.writeFileSync(outPath, lines.join('\n') + '\n');
        fileCount++;
        process.stderr.write(`  → ${outPath}\n`);
    }

    if (fileCount > 0) {
        process.stderr.write(`✅ ${fileCount} ファイルのテストテンプレートを出力しました\n`);
    } else {
        process.stderr.write('Function が見つかりませんでした。\n');
    }
}

function emitConstTs(analyses: FileAnalysis[], outputDir: string): void {
    const allRefs = new Set<string>();
    for (const a of analyses) {
        for (const name of a.xlVbConstantRefs) allRefs.add(name);
    }
    if (allRefs.size === 0) {
        process.stderr.write('xl*/vb* 定数が見つかりませんでした。\n');
        return;
    }

    const known: Array<[string, number | string]> = [];
    const unknown: string[] = [];
    for (const name of [...allRefs].sort()) {
        if (name in KNOWN_XL_CONSTANTS) {
            known.push([name, KNOWN_XL_CONSTANTS[name]]);
        } else {
            unknown.push(name);
        }
    }

    const lines: string[] = [
        '// Auto-generated by vba-analyzer --gen-test-dir',
        '// VBA/Excel constants referenced in VBA source.',
        '// Review values before use in tests.',
        '',
    ];
    for (const [name, value] of known) {
        lines.push(`export const ${name} = ${JSON.stringify(value)};`);
    }
    if (unknown.length > 0) {
        if (known.length > 0) lines.push('');
        lines.push('// TODO: values below are not in the known constants table — set actual values');
        for (const name of unknown) {
            lines.push(`export const ${name} = 0; // TODO`);
        }
    }

    // VBARunner.setConstants() に渡せるオブジェクトをまとめて export
    lines.push('');
    lines.push('/** vbaRunner.setConstants(allConstants) で一括注入するためのオブジェクト */');
    lines.push('export const allConstants: Record<string, number | string> = {');
    for (const [name, value] of known) {
        lines.push(`    ${name},`);
    }
    for (const name of unknown) {
        lines.push(`    ${name},`);
    }
    lines.push('};');

    fs.mkdirSync(outputDir, { recursive: true });
    const outPath = path.join(outputDir, 'const.ts');
    fs.writeFileSync(outPath, lines.join('\n') + '\n');
    process.stderr.write(`✅ ${allRefs.size} 件の定数を出力: ${outPath}\n`);
}

// ---------------------------------------------------------
// 単一ファイル解析
// ---------------------------------------------------------

// ---------------------------------------------------------
// Commented-out code detection
// ---------------------------------------------------------

// [キーワード正規表現, スコア重み] のリスト
const COMMENTED_CODE_PATTERNS: Array<[RegExp, number, string]> = [
    [/\bSub\b/gi,      3, 'Sub'],
    [/\bFunction\b/gi, 3, 'Function'],
    [/\bClass\b/gi,    3, 'Class'],
    [/\bProperty\b/gi, 3, 'Property'],
    [/\bType\b/gi,     2, 'Type'],
    [/\bDim\b/gi,      1, 'Dim'],
    [/\bConst\b/gi,    1, 'Const'],
    [/\bSet\b/gi,      1, 'Set'],
    [/\bReDim\b/gi,    1, 'ReDim'],
    [/\bIf\b/gi,       1, 'If'],
    [/\bFor\b/gi,      1, 'For'],
    [/\bDo\b/gi,       1, 'Do'],
    [/\bWhile\b/gi,    1, 'While'],
    [/\bWith\b/gi,     1, 'With'],
    [/\bSelect\b/gi,   1, 'Select'],
    [/\bCase\b/gi,     1, 'Case'],
    [/\bCall\b/gi,     1, 'Call'],
    [/\bEnd\b/gi,      1, 'End'],
    [/\bExit\b/gi,     1, 'Exit'],
    [/\bGoTo\b/gi,     1, 'GoTo'],
];

const CONTEXT_LINES = 3;

function detectCommentedCode(sourceText: string, filePath: string): CommentedCodeBlock[] {
    const lines = sourceText.split('\n');
    const results: CommentedCodeBlock[] = [];
    let blockStart = -1;
    const blockRawLines: string[] = [];
    const blockStrippedLines: string[] = [];

    const flush = (endIdx: number) => {
        if (blockStrippedLines.length >= 3) {
            const contextBefore = lines.slice(Math.max(0, blockStart - CONTEXT_LINES), blockStart);
            const contextAfter  = lines.slice(endIdx + 1, Math.min(lines.length, endIdx + 1 + CONTEXT_LINES));
            const block = scoreCommentBlock(
                [...blockStrippedLines], [...blockRawLines],
                contextBefore, contextAfter,
                filePath, blockStart, endIdx,
            );
            if (block) results.push(block);
        }
        blockStart = -1;
        blockRawLines.length = 0;
        blockStrippedLines.length = 0;
    };

    for (let i = 0; i < lines.length; i++) {
        const trimmed = lines[i].trimStart();
        if (trimmed.startsWith("'")) {
            if (blockStart === -1) blockStart = i;
            blockRawLines.push(trimmed);                          // ' を保持（表示用）
            blockStrippedLines.push(trimmed.slice(1).replace(/^ /, '')); // ' を除去（スコアリング用）
        } else {
            flush(i - 1);
        }
    }
    flush(lines.length - 1);

    return results.sort((a, b) => b.score - a.score);
}

function scoreCommentBlock(
    strippedLines: string[],
    rawLines: string[],
    contextBefore: string[],
    contextAfter: string[],
    filePath: string,
    startIdx: number,
    endIdx: number,
): CommentedCodeBlock | null {
    const content = strippedLines.join('\n');
    let totalScore = 0;
    const detectedKeywords: Record<string, number> = {};

    for (const [re, weight, label] of COMMENTED_CODE_PATTERNS) {
        const matches = content.match(re);
        if (matches) {
            detectedKeywords[label] = matches.length;
            totalScore += matches.length * weight;
        }
    }

    // 代入文を含む行（比較・不等号を除く）
    const assignLines = strippedLines.filter(l => /[^!=<>]=[^=]/.test(l)).length;
    totalScore += assignLines * 0.5;

    // メソッド/プロパティアクセスを含む行
    const dotLines = strippedLines.filter(l => /\w\.\w/.test(l)).length;
    totalScore += dotLines * 0.3;

    const score = Math.round(totalScore * 10) / 10;
    if (score < 2) return null;

    const confidence: 'high' | 'medium' | 'low' =
        score >= 8 ? 'high' : score >= 4 ? 'medium' : 'low';

    return {
        file: filePath,
        startLine: startIdx + 1,
        endLine: endIdx + 1,
        lineCount: strippedLines.length,
        score,
        confidence,
        detectedKeywords,
        strippedContent: content,
        rawContent: rawLines.join('\n'),
        contextBefore,
        contextAfter,
    };
}

interface FileAnalysis {
    report: FileReport;
    // クロスファイル分析に使う中間データ
    definedProcs: Map<string, { proc: any; kind: string; scope: string }>;
    definedTypes: Set<string>;
    callsByProc: Map<string, Set<string>>;  // procName -> called identifiers
    xlVbConstantRefs: Set<string>;          // xl*/vb*/mso* 定数の参照名
    commentedCodeBlocks: CommentedCodeBlock[];
}

function analyzeFile(filePath: string): FileAnalysis {
    const src = fs.readFileSync(filePath, 'utf-8');
    const lines = src.split('\n');
    const warnings: string[] = [];

    let ast: any;
    try {
        const tokens = new Lexer(preprocess(src)).tokenize();
        ast = new Parser(tokens).parse();
    } catch (e: any) {
        warnings.push(`Parse error: ${e.message}`);
        return {
            report: { filePath, totalLines: lines.length, procedureCount: 0, procedures: [], prefixClusters: [], warnings },
            definedProcs: new Map(),
            definedTypes: new Set(),
            callsByProc: new Map(),
            xlVbConstantRefs: new Set(),
            commentedCodeBlocks: detectCommentedCode(src, filePath),
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
        const magicLits = findMagicLiteralsInCalls(proc.body);
        const byRefAssignments = findByRefAssignments(proc);
        const ioSideEffectCount = findIoSideEffects(proc.body);
        const parameters: Array<{ name: string; paramType: string | null; isByVal: boolean }> =
            (proc.parameters ?? []).map((p: any) => ({
                name: p.name as string,
                paramType: p.paramType ?? null,
                isByVal: p.isByVal ?? false,
            }));
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
            magicLiteralsInCalls: magicLits,
            byRefAssignments,
            parameters,
            returnType: proc.returnType ?? null,
            ioSideEffectCount,
            referenceCount: 0,  // 後段のワークスペース解析で埋める
            hardcodedSheetCount: magicLits.filter(m =>
                (m.callee.toLowerCase() === 'sheets' || m.callee.toLowerCase() === 'worksheets') && m.argIndex === 0
            ).length,
            hardcodedAddressCount: magicLits.filter(m =>
                m.callee.toLowerCase() === 'range' || m.callee.toLowerCase() === 'cells'
            ).length,
        };
    });

    const prefixClusters = detectPrefixClusters(ast.body);
    const xlVbConstantRefs = collectExcelConstantRefs(ast);

    return {
        report: { filePath, totalLines: lines.length, procedureCount: procedures.length, procedures, prefixClusters, warnings },
        definedProcs,
        definedTypes,
        callsByProc,
        xlVbConstantRefs,
        commentedCodeBlocks: detectCommentedCode(src, filePath),
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

    const duplicateBlocks = detectDuplicateBlocks(analyses);

    const commentedCodeBlocks = analyses
        .flatMap(a => a.commentedCodeBlocks)
        .sort((a, b) => b.score - a.score);

    return {
        files: analyses.map(a => a.report),
        entryPointCandidates,
        deadCodeCandidates,
        excelMockTargets,
        callGraph,
        duplicateBlocks,
        commentedCodeBlocks,
    };
}

// ---------------------------------------------------------
// 凝集度判定
// ---------------------------------------------------------

// 各指標のペナルティスコア合計で HIGH/MEDIUM/LOW を判定する。
//
// 変数数:           0-7 → 0点  8-14 → 1点  15-29 → 2点  30+ → 3点
// 固定シート参照:   0-1 → 0点  2    → 1点  3-4  → 2点  5+  → 3点
// 固定アドレス引数: 0-2 → 0点  3-4  → 1点  5-7  → 2点  8+  → 3点
//
// 合計スコア: 0-1 → HIGH ✅   2-3 → MED ⚠️   4+ → LOW ❌
function cohesionJudge(p: ProcedureMetrics): { label: string; reasons: string[] } {
    let score = 0;
    const reasons: string[] = [];

    // 変数の数
    if      (p.localDeclCount >= 30) { score += 3; reasons.push(`変数多(${p.localDeclCount}個)`); }
    else if (p.localDeclCount >= 15) { score += 2; reasons.push(`変数多(${p.localDeclCount}個)`); }
    else if (p.localDeclCount >= 8)  { score += 1; }

    // 固定シート参照の種類数
    if      (p.hardcodedSheetCount >= 5) { score += 3; reasons.push(`シート参照多(${p.hardcodedSheetCount}種)`); }
    else if (p.hardcodedSheetCount >= 3) { score += 2; reasons.push(`シート参照多(${p.hardcodedSheetCount}種)`); }
    else if (p.hardcodedSheetCount >= 2) { score += 1; }

    // 固定アドレス引数の種類数
    if      (p.hardcodedAddressCount >= 8) { score += 3; reasons.push(`アドレス多(${p.hardcodedAddressCount}種)`); }
    else if (p.hardcodedAddressCount >= 5) { score += 2; reasons.push(`アドレス多(${p.hardcodedAddressCount}種)`); }
    else if (p.hardcodedAddressCount >= 3) { score += 1; }

    if (score >= 3) return { label: 'LOW  ❌', reasons };
    if (score >= 1) return { label: 'MED  ⚠️ ', reasons };
    return { label: 'HIGH ✅', reasons: [] };
}

// 行数の良し悪し: <30 ✅  30-99 ⚠️  100+ ❌
function lineEmoji(n: number): string {
    if (n >= 100) return '❌';
    if (n >= 30)  return '⚠️ ';
    return '✅';
}
// ネスト深さの良し悪し: 0-2 ✅  3-4 ⚠️  5+ ❌
function nestEmoji(n: number): string {
    if (n >= 5) return '❌';
    if (n >= 3) return '⚠️ ';
    return '✅';
}
// Excel アクセス数の良し悪し: 0-4 ✅  5-9 ⚠️  10+ ❌
function excelEmoji(n: number): string {
    if (n >= 10) return '❌';
    if (n >= 5)  return '⚠️ ';
    return '✅';
}
// 重複ブロックの種類数の良し悪し: 0 ✅  1-2 ⚠️  3+ ❌
function dupEmoji(n: number): string {
    if (n >= 3) return '❌';
    if (n >= 1) return '⚠️ ';
    return '✅';
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
        out.push(`  ${scopeLabel}[${p.kind}] ${p.name}  (L${p.startLine}-L${p.endLine}, refs=${p.referenceCount})`);

        // 指標サマリー（常に表示、各指標に良し悪しの絵文字付き）
        out.push(`    行数 ${p.lineCount}${lineEmoji(p.lineCount)} / ネスト ${p.maxNestDepth}${nestEmoji(p.maxNestDepth)} / Excel ${p.excelAccessCount}${excelEmoji(p.excelAccessCount)}`);

        // 凝集度
        const cohesion = cohesionJudge(p);
        const cohesionReasons = cohesion.reasons.length ? `  ← ${cohesion.reasons.join(', ')}` : '';
        out.push(`    凝集度: ${cohesion.label}  (Dim ${p.localDeclCount} / シート参照 ${p.hardcodedSheetCount}種 / アドレス ${p.hardcodedAddressCount}種)${cohesionReasons}`);

        if (p.excelObjectsUsed.length > 0) {
            out.push(`    🧪 モック必要候補: ${p.excelObjectsUsed.join(', ')}`);
        }
        if (p.assignmentBlocks.length) {
            out.push(`    ⚠️  連続代入ブロック（関数抽出候補）:`);
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
            out.push(`    ⚠️  繰り返し数値リテラル（定数化候補）: ${top.map(n => `${n.value}(×${n.occurrences})`).join(', ')}`);
        }
        if (p.byRefAssignments.length) {
            out.push(`    ❌ ByRef パラメータへの代入（UDT 戻り値リファクタリング候補）:`);
            for (const b of p.byRefAssignments) {
                const typeStr = b.paramType ? ` As ${b.paramType}` : '';
                const linesStr = b.lines.map(l => `L${l}`).join(', ');
                out.push(`      ${b.paramName}${typeStr}: ${b.lines.length}件 [${linesStr}]`);
            }
        }
        if (p.magicLiteralsInCalls.length) {
            out.push(`    ⚠️  即値引数（定数化候補）: ${p.magicLiteralsInCalls.length}種`);
            for (const m of p.magicLiteralsInCalls.slice(0, 10)) {
                const val = typeof m.value === 'string' ? `"${m.value}"` : String(m.value);
                const linesStr = m.lines.map(l => `L${l}`).join(', ');
                const pname = paramName(m.callee, m.argIndex);
                out.push(`      ${m.callee}(${pname}=${val}): ${m.lines.length}件 [${linesStr}]`);
            }
            if (p.magicLiteralsInCalls.length > 10) {
                out.push(`      ... 他 ${p.magicLiteralsInCalls.length - 10} 種`);
            }
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

    if (w.duplicateBlocks.length > 0) {
        out.push('');
        out.push('========================================');
        out.push('♻️  重複ブロック候補（関数抽出候補）');
        out.push('========================================');
        out.push('  以下は AST レベルで同一パターンが複数箇所に現れるブロックです。');
        out.push('  変数名は正規化済み（$ID）のため、名前が異なっても構造が同じなら一致します。');
        out.push('  ※ 制御構造のボディは別シーケンスとして評価するため、');
        out.push('     "If 内の3行" と "For 内の3行" が同じ構造なら別々に報告されます。');
        out.push('');
        for (const b of w.duplicateBlocks) {
            out.push(`  [${b.stmtCount}文×${b.occurrences.length}箇所]  ${b.shape}`);
            for (const o of b.occurrences) {
                out.push(`    ${o.procName} (${path.basename(o.file)}: L${o.startLine}-L${o.endLine})`);
            }
        }
    }

    return out.join('\n');
}

function formatCommentedCodeBlocks(blocks: CommentedCodeBlock[]): string {
    if (blocks.length === 0) return '(コメントアウトされたコードは検出されませんでした)';
    const out: string[] = [];
    out.push('========================================');
    out.push('💬 コメントアウトされたコード候補');
    out.push('========================================');
    out.push('  連続するコメント行の中に VBA キーワードが多く含まれるブロックです。');
    out.push('  AI に渡すなど、人間が最終判断してください。');
    out.push('');

    const pad = (n: number) => String(n).padStart(5);

    for (const b of blocks) {
        const fname = path.basename(b.file);
        const kwStr = Object.entries(b.detectedKeywords)
            .map(([k, n]) => `${k}×${n}`)
            .join(', ');
        out.push(`  [${b.confidence.toUpperCase()}] ${fname} L${b.startLine}-L${b.endLine}  (${b.lineCount}行, score=${b.score})`);
        if (kwStr) out.push(`    キーワード: ${kwStr}`);
        out.push('    ---');

        // コンテキスト（前）
        const beforeStart = b.startLine - b.contextBefore.length;
        for (let i = 0; i < b.contextBefore.length; i++) {
            out.push(`    ${pad(beforeStart + i)}: ${b.contextBefore[i]}`);
        }

        // コメントブロック本体（' を保持）
        const rawLines = b.rawContent.split('\n');
        for (let i = 0; i < rawLines.length; i++) {
            out.push(`  > ${pad(b.startLine + i)}: ${rawLines[i]}`);
        }

        // コンテキスト（後）
        for (let i = 0; i < b.contextAfter.length; i++) {
            out.push(`    ${pad(b.endLine + 1 + i)}: ${b.contextAfter[i]}`);
        }

        out.push('');
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
        console.log('  --json                JSON 形式で出力（プログラム連携用）');
        console.log('  --summary-only        ワークスペース要約（エントリーポイント候補・モック必要箇所）のみ表示');
        console.log('  --outline             Workspace Outline（AI向けコンテキスト圧縮）を出力');
        console.log('  --commented-code      コメントアウトされたコード候補のみ表示');
        console.log('  --gen-test-dir <dir>  テスト用ソースを指定ディレクトリに生成（const.ts 等）');
        process.exit(args.length === 0 ? 1 : 0);
    }
    const wantJson = args.includes('--json');
    const summaryOnly = args.includes('--summary-only');
    const wantOutline = args.includes('--outline');
    const wantCommentedCode = args.includes('--commented-code');
    const genTestDirIdx = args.indexOf('--gen-test-dir');
    const genTestDir = genTestDirIdx !== -1 ? args[genTestDirIdx + 1] : null;
    if (genTestDirIdx !== -1 && (!genTestDir || genTestDir.startsWith('--'))) {
        console.error('--gen-test-dir には出力ディレクトリを指定してください');
        process.exit(1);
    }
    const target = args.find(a => !a.startsWith('--') && a !== genTestDir);
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

    if (genTestDir) {
        emitConstTs(analyses, genTestDir);
        emitTestTemplates(analyses, genTestDir);
    }

    if (wantJson) {
        console.log(JSON.stringify(workspace, null, 2));
        return;
    }

    if (wantOutline) {
        console.log(formatWorkspaceOutline(workspace));
        return;
    }

    if (wantCommentedCode) {
        console.log(formatCommentedCodeBlocks(workspace.commentedCodeBlocks));
        return;
    }

    if (!summaryOnly) {
        for (const r of workspace.files) console.log(formatFileReport(r));
    }
    console.log(formatWorkspaceSummary(workspace));
}

if (process.argv[1]?.includes('vba-analyzer')) main();
