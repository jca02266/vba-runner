import { LintDiagnostic } from '../engine/vba-lint';
import {
    Statement,
    VariableDeclaration,
    ProcedureDeclaration,
    ClassDeclaration,
    TypeDeclaration,
    EnumDeclaration,
} from '../engine/parser';

// VBA 組み込みプリミティブ型 — これらは型定義不要
const VBA_PRIMITIVE_TYPES = new Set([
    'string', 'long', 'integer', 'double', 'boolean', 'date', 'object',
    'variant', 'byte', 'single', 'currency', 'longlongptr', 'longlong',
    'longptr', 'error', 'empty', 'nothing', 'null', 'any',
]);

/** AST から全ユーザー定義型名（Class / Type / Enum）を収集して返す。 */
export function collectUserDefinedTypeNames(statements: Statement[]): Set<string> {
    const names = new Set<string>();
    for (const stmt of statements) {
        if (stmt.type === 'ClassDeclaration') {
            names.add((stmt as ClassDeclaration).name.toLowerCase());
        } else if (stmt.type === 'TypeDeclaration') {
            names.add((stmt as TypeDeclaration).name.toLowerCase());
        } else if (stmt.type === 'EnumDeclaration') {
            names.add((stmt as EnumDeclaration).name.name.toLowerCase());
        }
    }
    return names;
}

/**
 * VBA016: 未知の型名を使っている宣言を検出する。
 * `knownTypeNames` に含まれない型名を `As TypeName` で使っている場合に警告を出す。
 */
export function checkUnknownTypes(
    statements: Statement[],
    knownTypeNames: Set<string>,
): LintDiagnostic[] {
    const diags: LintDiagnostic[] = [];

    function warn(typeName: string, loc: any): void {
        const start  = loc?.start;
        const end    = loc?.end;
        if (!start) return;
        const line    = (start.line    ?? 1) - 1;
        const column  = (start.column  ?? 1) - 1;
        const endLine = (end?.line     ?? start.line) - 1;
        const endCol  = (end?.column   ?? start.column + typeName.length) - 1;
        diags.push({
            code: 'VBA016',
            severity: 2, // Warning
            message: `Unknown type '${typeName}'`,
            l10nKey: "Unknown type '{0}'",
            l10nArgs: [typeName],
            line,
            column,
            endLine,
            endColumn: endCol,
        });
    }

    function isKnown(typeName: string): boolean {
        const lower = typeName.toLowerCase();
        if (VBA_PRIMITIVE_TYPES.has(lower)) return true;
        if (knownTypeNames.has(lower)) return true;
        // モジュール修飾を除いた短縮名でも検索: "Excel.Range" → "range"
        const short = lower.replace(/^[a-z_][a-z0-9_]*\./, '');
        return knownTypeNames.has(short);
    }

    function scanStmts(stmts: Statement[]): void {
        for (const stmt of stmts) {
            if (stmt.type === 'VariableDeclaration') {
                const vd = stmt as VariableDeclaration;
                for (const d of vd.declarations) {
                    if (d.objectType && !isKnown(d.objectType)) {
                        warn(d.objectType, d.name.loc ?? stmt.loc);
                    }
                }
            } else if (stmt.type === 'ProcedureDeclaration') {
                const proc = stmt as ProcedureDeclaration;
                for (const param of proc.parameters) {
                    if (param.paramType && !isKnown(param.paramType)) {
                        warn(param.paramType, param.loc ?? proc.loc);
                    }
                }
                if (proc.isFunction && proc.returnType && !isKnown(proc.returnType)) {
                    warn(proc.returnType, proc.loc);
                }
                scanStmts(proc.body);
            } else if (stmt.type === 'ClassDeclaration') {
                scanStmts((stmt as ClassDeclaration).body);
            }
        }
    }

    scanStmts(statements);
    return diags;
}
