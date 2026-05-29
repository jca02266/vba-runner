/**
 * GoTo ↔ LabelStatement のナビゲーション支援
 *
 * - GoTo のラベル名にカーソルがある → LabelStatement の定義位置を返す
 * - LabelStatement にカーソルがある → そのラベルを参照する GoTo 一覧を返す
 */

import { Statement, GoToStatement, LabelStatement } from '../engine/parser';

export interface LabelLocation {
    line: number;   // 0-based
    character: number;
    endCharacter: number;
}

interface LabelDef {
    name: string;         // lowercase
    loc: LabelLocation;
}

interface GoToRef {
    labelName: string;    // lowercase
    loc: LabelLocation;   // GoTo キーワードの次のラベル名部分
}

// ─── AST walker ──────────────────────────────────────────────────────────────

function collectLabels(stmts: Statement[], defs: LabelDef[], gotos: GoToRef[]): void {
    for (const stmt of stmts) {
        if (stmt.type === 'LabelStatement') {
            const ls = stmt as LabelStatement;
            const loc = (ls as any).loc;
            if (loc) {
                defs.push({
                    name: ls.label.toLowerCase(),
                    loc: {
                        line: loc.start.line - 1,
                        character: loc.start.column - 1,
                        endCharacter: loc.start.column - 1 + ls.label.length,
                    },
                });
            }
        } else if (stmt.type === 'GoToStatement') {
            const gs = stmt as GoToStatement;
            const loc = (gs as any).loc;
            if (loc) {
                // GoTo keyword is 4 chars; label starts after "GoTo "
                const labelChar = loc.start.column - 1 + 5; // "GoTo " = 5 chars
                gotos.push({
                    labelName: gs.label.toLowerCase(),
                    loc: {
                        line: loc.start.line - 1,
                        character: labelChar,
                        endCharacter: labelChar + gs.label.length,
                    },
                });
            }
        }

        // Recurse into children
        recurse(stmt, defs, gotos);
    }
}

function recurse(stmt: Statement, defs: LabelDef[], gotos: GoToRef[]): void {
    const s = stmt as any;
    const bodies: Statement[][] = [];

    if (Array.isArray(s.body))       bodies.push(s.body);
    if (Array.isArray(s.consequent)) bodies.push(s.consequent);
    if (Array.isArray(s.alternate))  bodies.push(s.alternate);
    if (Array.isArray(s.elseBody))   bodies.push(s.elseBody);
    if (s.cases) for (const c of s.cases) if (Array.isArray(c.body)) bodies.push(c.body);
    if (s.procedures) bodies.push(s.procedures);
    if (s.fields) bodies.push(s.fields);

    for (const body of bodies) collectLabels(body, defs, gotos);
}

// ─── 公開API ─────────────────────────────────────────────────────────────────

function buildIndex(statements: Statement[]): { defs: LabelDef[]; gotos: GoToRef[] } {
    const defs: LabelDef[] = [];
    const gotos: GoToRef[] = [];
    collectLabels(statements, defs, gotos);
    return { defs, gotos };
}

function hitTest(loc: LabelLocation, line: number, character: number): boolean {
    return loc.line === line && loc.character <= character && character <= loc.endCharacter;
}

/**
 * カーソル位置が GoTo のラベル名上にある場合、対応する LabelStatement の位置を返す。
 */
export function findLabelDefinition(
    statements: Statement[],
    line: number,
    character: number,
    uri: string,
): { uri: string; range: { start: { line: number; character: number }; end: { line: number; character: number } } } | null {
    const { defs, gotos } = buildIndex(statements);

    const hit = gotos.find(g => hitTest(g.loc, line, character));
    if (!hit) return null;

    const def = defs.find(d => d.name === hit.labelName);
    if (!def) return null;

    return {
        uri,
        range: {
            start: { line: def.loc.line, character: def.loc.character },
            end:   { line: def.loc.line, character: def.loc.endCharacter },
        },
    };
}

/**
 * カーソル位置が LabelStatement 上にある場合、そのラベルを参照する GoTo 一覧を返す。
 */
export function findGoToReferences(
    statements: Statement[],
    line: number,
    character: number,
    uri: string,
    includeDeclaration: boolean,
): Array<{ uri: string; range: { start: { line: number; character: number }; end: { line: number; character: number } } }> {
    const { defs, gotos } = buildIndex(statements);

    const hitDef = defs.find(d => hitTest(d.loc, line, character));
    if (!hitDef) return [];

    const refs: Array<{ uri: string; range: { start: { line: number; character: number }; end: { line: number; character: number } } }> = [];

    if (includeDeclaration) {
        refs.push({
            uri,
            range: {
                start: { line: hitDef.loc.line, character: hitDef.loc.character },
                end:   { line: hitDef.loc.line, character: hitDef.loc.endCharacter },
            },
        });
    }

    for (const g of gotos) {
        if (g.labelName !== hitDef.name) continue;
        refs.push({
            uri,
            range: {
                start: { line: g.loc.line, character: g.loc.character },
                end:   { line: g.loc.line, character: g.loc.endCharacter },
            },
        });
    }

    return refs;
}

/**
 * カーソル位置がラベル関連（GoTo のラベル名 or LabelStatement）かどうかを判定する。
 * definition/references プロバイダーが label ロジックを使うか判断するために使う。
 */
export function isOnLabel(statements: Statement[], line: number, character: number): boolean {
    const { defs, gotos } = buildIndex(statements);
    return defs.some(d => hitTest(d.loc, line, character)) ||
           gotos.some(g => hitTest(g.loc, line, character));
}
