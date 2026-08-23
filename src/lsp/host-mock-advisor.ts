import {
    Program,
    Statement,
    Expression,
    Identifier,
    MemberExpression,
    CallExpression,
} from '../engine/parser';
import { buildScopedSymbolTable } from './symbol-table';

export interface HostMockDiagnostic {
    range: {
        start: { line: number; character: number };
        end: { line: number; character: number };
    };
    severity: 3;
    message: string;
    source: 'vba-mock-advisor';
}

const MOCK_GUIDE_URL = 'https://github.com/jca02266/vba-runner/blob/main/docs/guides/MOCK_GUIDE.md';

const HOST_IDENTIFIERS = new Map<string, string>([
    // Excel application/workbook globals
    ['application', 'Excel/Access Application'],
    ['thisworkbook', 'Excel ThisWorkbook'],
    ['activeworkbook', 'Excel ActiveWorkbook'],
    ['activesheet', 'Excel ActiveSheet'],
    ['workbooks', 'Excel Workbooks'],
    ['worksheets', 'Excel Worksheets'],
    ['sheets', 'Excel Sheets'],
    ['range', 'Excel Range'],
    ['cells', 'Excel Cells'],
    ['rows', 'Excel Rows'],
    ['columns', 'Excel Columns'],
    ['selection', 'Excel Selection'],
    ['charts', 'Excel Charts'],
    // Access application globals
    ['currentdb', 'Access CurrentDb'],
    ['currentproject', 'Access CurrentProject'],
    ['docmd', 'Access DoCmd'],
    ['forms', 'Access Forms'],
    ['reports', 'Access Reports'],
    ['screen', 'Access Screen'],
]);

function diagnosticFor(id: Identifier, kind: string): HostMockDiagnostic {
    const loc = id.loc;
    const line = (loc?.start.line ?? 1) - 1;
    const character = (loc?.start.column ?? 1) - 1;
    const name = id.name;
    return {
        range: {
            start: { line, character },
            end: { line, character: character + name.length },
        },
        severity: 3,
        source: 'vba-mock-advisor',
        message: `${kind} identifier '${name}' depends on an Excel/Access host. `
            + `Create a mock in __mocks__ before running with vba-runner. `
            + `See ${MOCK_GUIDE_URL}. `
            + `This diagnostic alone does not contain the complete member list; `
            + `include the full current VBA source when giving it to an AI. `
            + `AI prompt example: "Read the mock guide at ${MOCK_GUIDE_URL}. `
            + `Inspect the complete VBA source included with this diagnostic, `
            + `enumerate every member used on '${name}', and create the smallest `
            + `__mocks__ VBA/JS stub that preserves those calls. Do not guess `
            + `members from this diagnostic alone; add a focused test."`,
    };
}

function visitExpression(
    expr: Expression | undefined,
    declared: Set<string>,
    out: HostMockDiagnostic[],
): void {
    if (!expr) return;
    switch (expr.type) {
        case 'Identifier': {
            const id = expr as Identifier;
            const kind = HOST_IDENTIFIERS.get(id.name.toLowerCase());
            if (kind && !declared.has(id.name.toLowerCase())) out.push(diagnosticFor(id, kind));
            return;
        }
        case 'MemberExpression': {
            const member = expr as MemberExpression;
            // The property name is not a host reference; inspect only the object.
            visitExpression(member.object, declared, out);
            return;
        }
        case 'CallExpression': {
            const call = expr as CallExpression;
            visitExpression(call.callee, declared, out);
            for (const arg of call.args) visitExpression(arg, declared, out);
            return;
        }
        case 'BinaryExpression':
        case 'LogicalExpression': {
            const binary = expr as any;
            visitExpression(binary.left, declared, out);
            visitExpression(binary.right, declared, out);
            return;
        }
        case 'UnaryExpression':
        case 'ParenthesizedExpression':
            visitExpression((expr as any).argument ?? (expr as any).expression, declared, out);
            return;
        case 'NamedArgument':
            visitExpression((expr as any).value, declared, out);
            return;
        case 'DictionaryAccessExpression':
            visitExpression((expr as any).object, declared, out);
            visitExpression((expr as any).property, declared, out);
            return;
        default:
            return;
    }
}

function visitStatement(stmt: Statement, declared: Set<string>, out: HostMockDiagnostic[]): void {
    if (stmt.type === 'ProcedureDeclaration') {
        const proc = stmt as any;
        for (const bodyStmt of proc.body ?? []) visitStatement(bodyStmt, declared, out);
        return;
    }
    if (stmt.type === 'ClassDeclaration') {
        const cls = stmt as any;
        for (const field of cls.fields ?? []) visitStatement(field, declared, out);
        for (const proc of cls.procedures ?? []) visitStatement(proc, declared, out);
        return;
    }
    if (stmt.type === 'VariableDeclaration' || stmt.type === 'ConstDeclaration') {
        for (const declaration of (stmt as any).declarations ?? []) {
            visitExpression(declaration.initializer, declared, out);
            for (const bound of declaration.bounds ?? []) {
                visitExpression(bound.lower, declared, out);
                visitExpression(bound.upper, declared, out);
            }
        }
        return;
    }

    // Walk expression-valued statement fields without treating declaration
    // names, labels, or source locations as identifier references.
    for (const [key, value] of Object.entries(stmt as any)) {
        if (key === 'type' || key === 'loc' || key === 'name' || key === 'label') continue;
        if (Array.isArray(value)) {
            for (const item of value) {
                if (item?.type === 'Identifier' || item?.type?.endsWith('Expression')) {
                    visitExpression(item, declared, out);
                } else if (item?.type) {
                    visitStatement(item, declared, out);
                }
            }
        } else if (value?.type === 'Identifier' || value?.type?.endsWith('Expression')) {
            visitExpression(value, declared, out);
        } else if (value?.type) {
            visitStatement(value, declared, out);
        }
    }
}

/** Return host-object mock recommendations for a parsed VBA document. */
export function collectHostMockDiagnostics(program: Program): HostMockDiagnostic[] {
    const table = buildScopedSymbolTable(program.body);
    const declared = new Set<string>(table.moduleSymbols.keys());
    for (const procedure of table.procedures) {
        for (const name of procedure.localSymbols.keys()) declared.add(name);
    }
    const out: HostMockDiagnostic[] = [];
    for (const stmt of program.body) visitStatement(stmt, declared, out);
    return out;
}
