import {
    ClassDeclaration,
    EventDeclaration,
    ProcedureDeclaration,
    VariableDeclaration,
    ConstDeclaration,
    Statement,
} from '../engine/parser';

/** 'public' → 'Public'、'private' → 'Private' など先頭大文字化 */
function cap(s: string): string { return s.charAt(0).toUpperCase() + s.slice(1); }

/** 定数値の AST ノードを表示用テキストに変換（リテラル以外は空文字を返す） */
function constLiteralText(expr: any): string {
    if (!expr) return '';
    if (expr.type === 'NumberLiteral')  return String(expr.value);
    if (expr.type === 'StringLiteral')  return `"${expr.value}"`;
    if (expr.type === 'BooleanLiteral') return expr.value ? 'True' : 'False';
    if (expr.type === 'Identifier')     return expr.name; // True / False / Nothing
    if (expr.type === 'UnaryExpression' && expr.operator === '-') {
        const inner = constLiteralText(expr.operand ?? expr.right);
        return inner ? `-${inner}` : '';
    }
    return '';
}

export type SymbolKind =
    | 'module-var' | 'local-var' | 'param' | 'for' | 'for-each'
    | 'procedure' | 'class' | 'const' | 'event';

export interface SymbolEntry {
    name: string;
    displayText: string;
    kind: SymbolKind;
    range: {
        start: { line: number; character: number };
        end: { line: number; character: number };
    };
}

export interface SymbolLookupResult {
    entry: SymbolEntry;
    /** Enclosing procedure name; null when symbol is at module level */
    procName: string | null;
}

export interface ProcedureScope {
    name: string;
    /** 0-based line range of the entire procedure (Sub ... End Sub) */
    range: {
        start: { line: number; character: number };
        end: { line: number; character: number };
    };
    localSymbols: Map<string, SymbolEntry>;
}

export interface ScopedSymbolTable {
    /** Module-level declarations: procedures, module Dims, Consts, Classes */
    moduleSymbols: Map<string, SymbolEntry>;
    /** Per-procedure scopes with local declarations */
    procedures: ProcedureScope[];
}

// ─── text helpers ────────────────────────────────────────────────────────────

export function getWordAtPosition(text: string, line: number, character: number): string | null {
    const lines = text.split('\n');
    if (line >= lines.length) return null;
    const lineText = lines[line];

    let start = character;
    let end = character;
    while (start > 0 && /[a-zA-Z0-9_]/.test(lineText[start - 1])) start--;
    while (end < lineText.length && /[a-zA-Z0-9_]/.test(lineText[end])) end++;

    if (start === end) return null;
    return lineText.slice(start, end);
}

// ─── scope-aware builder ─────────────────────────────────────────────────────

export function buildScopedSymbolTable(statements: Statement[]): ScopedSymbolTable {
    const moduleSymbols = new Map<string, SymbolEntry>();
    const procedures: ProcedureScope[] = [];
    collectScopedSymbols(statements, moduleSymbols, procedures);
    return { moduleSymbols, procedures };
}

function collectScopedSymbols(
    statements: Statement[],
    moduleSymbols: Map<string, SymbolEntry>,
    procedures: ProcedureScope[],
): void {
    for (const stmt of statements) {
        if (stmt.type === 'ProcedureDeclaration') {
            const proc = stmt as ProcedureDeclaration;
            if (!proc.loc) continue;

            const procEntry = makeProcEntry(proc);
            moduleSymbols.set(proc.name.name.toLowerCase(), procEntry);

            const localSymbols = new Map<string, SymbolEntry>();

            // Parameters
            for (const param of proc.parameters) {
                const pname = param.name as unknown as string;
                if (!pname) continue;
                const ptype = param.paramType || 'Variant';
                const modifier = param.isByVal ? 'ByVal ' : '';
                // param has no loc in AST; best-effort: position them at the procedure header line
                const procLine = proc.loc.start.line - 1;
                localSymbols.set(pname.toLowerCase(), {
                    name: pname,
                    displayText: `${modifier}${pname} As ${ptype}`,
                    kind: 'param',
                    range: {
                        start: { line: procLine, character: 0 },
                        end: { line: procLine, character: pname.length },
                    },
                });
            }

            // Local declarations (body)
            collectLocalDeclarations(proc.body, localSymbols);

            procedures.push({
                name: proc.name.name,
                range: {
                    start: { line: proc.loc.start.line - 1, character: 0 },
                    end: { line: proc.loc.end.line - 1, character: Number.MAX_SAFE_INTEGER },
                },
                localSymbols,
            });

        } else if (stmt.type === 'VariableDeclaration') {
            addVariableDeclaration(stmt as VariableDeclaration, moduleSymbols, 'module-var');

        } else if (stmt.type === 'ConstDeclaration') {
            addConstDeclaration(stmt as ConstDeclaration, moduleSymbols);

        } else if (stmt.type === 'ClassDeclaration') {
            const cls = stmt as ClassDeclaration;
            if (!cls.loc) continue;

            const clsLine = cls.loc.start.line - 1;
            const nameStart = (cls.loc.start.column - 1) + 6; // 'Class '
            moduleSymbols.set(cls.name.toLowerCase(), {
                name: cls.name,
                displayText: `Class ${cls.name}`,
                kind: 'class',
                range: { start: { line: clsLine, character: nameStart }, end: { line: clsLine, character: nameStart + cls.name.length } },
            });
            collectScopedSymbols(cls.body, moduleSymbols, procedures);

        } else if (stmt.type === 'EventDeclaration') {
            const evt = stmt as EventDeclaration;
            if (!evt.loc) continue;
            const evtLine = evt.loc.start.line - 1;
            const nameStart = (evt.loc.start.column - 1) + 6; // 'Event '
            const evtName = evt.name.name;
            const params = evt.parameters
                .map(p => `${p.name} As ${p.paramType || 'Variant'}`)
                .join(', ');
            moduleSymbols.set(evtName.toLowerCase(), {
                name: evtName,
                displayText: `Event ${evtName}(${params})`,
                kind: 'event',
                range: { start: { line: evtLine, character: nameStart }, end: { line: evtLine, character: nameStart + evtName.length } },
            });
        }
    }
}

/** Collect Dim/Const inside a Sub/Function body (does not recurse into nested Sub). */
function collectLocalDeclarations(body: Statement[], out: Map<string, SymbolEntry>): void {
    for (const stmt of body) {
        if (stmt.type === 'VariableDeclaration') {
            addVariableDeclaration(stmt as VariableDeclaration, out);
        } else if (stmt.type === 'ConstDeclaration') {
            addConstDeclaration(stmt as ConstDeclaration, out);
        } else if (stmt.type === 'ForStatement') {
            const f = stmt as any;
            // For loop variable: use its identifier loc if available
            if (f.identifier?.loc) {
                const id = f.identifier;
                const ln = id.loc.start.line - 1;
                const col = id.loc.start.column - 1;
                out.set(id.name.toLowerCase(), {
                    name: id.name,
                    displayText: id.name,
                    kind: 'for',
                    range: { start: { line: ln, character: col }, end: { line: ln, character: col + id.name.length } },
                });
            }
            collectLocalDeclarations(f.body ?? [], out);
        } else if (stmt.type === 'ForEachStatement') {
            const f = stmt as any;
            if (f.variable?.loc) {
                const id = f.variable;
                const ln = id.loc.start.line - 1;
                const col = id.loc.start.column - 1;
                out.set(id.name.toLowerCase(), {
                    name: id.name,
                    displayText: id.name,
                    kind: 'for-each',
                    range: { start: { line: ln, character: col }, end: { line: ln, character: col + id.name.length } },
                });
            }
            collectLocalDeclarations((f.body ?? []), out);
        } else if (stmt.type === 'IfStatement') {
            const s = stmt as any;
            collectLocalDeclarations(s.consequent ?? [], out);
            if (Array.isArray(s.alternate)) collectLocalDeclarations(s.alternate, out);
            else if (s.alternate) collectLocalDeclarations([s.alternate], out);
        } else if (stmt.type === 'DoWhileStatement' || stmt.type === 'WhileStatement' || stmt.type === 'WithStatement') {
            collectLocalDeclarations((stmt as any).body ?? [], out);
        } else if (stmt.type === 'SelectCaseStatement') {
            const sc = stmt as any;
            for (const c of sc.cases ?? []) collectLocalDeclarations(c.body ?? [], out);
            if (sc.elseBody) collectLocalDeclarations(sc.elseBody, out);
        }
    }
}

function addVariableDeclaration(
    decl: VariableDeclaration,
    out: Map<string, SymbolEntry>,
    kind: SymbolKind = 'local-var',
): void {
    if (!decl.loc) return;
    // Module-level keywords indicate a module-level variable
    const isModuleLevel = decl.scope === 'private' || decl.scope === 'public' || decl.scope === 'friend';
    const effectiveKind: SymbolKind = isModuleLevel ? 'module-var' : kind;
    for (const d of decl.declarations) {
        if (d.name.loc) {
            const ln = d.name.loc.start.line - 1;
            const col = d.name.loc.start.column - 1;
            const varType = d.objectType || 'Variant';
            const keyword = decl.scope ? cap(decl.scope) : 'Dim';
            out.set(d.name.name.toLowerCase(), {
                name: d.name.name,
                displayText: `${keyword} ${d.name.name} As ${varType}`,
                kind: effectiveKind,
                range: { start: { line: ln, character: col }, end: { line: ln, character: col + d.name.name.length } },
            });
        } else {
            // fallback: compute from statement loc + keyword offset
            const declLine = decl.loc!.start.line - 1;
            const declCol  = decl.loc!.start.column - 1;
            const keyword = decl.scope ? cap(decl.scope) : 'Dim';
            const nameStart = declCol + keyword.length + 1;
            const varType = d.objectType || 'Variant';
            out.set(d.name.name.toLowerCase(), {
                name: d.name.name,
                displayText: `${keyword} ${d.name.name} As ${varType}`,
                kind: effectiveKind,
                range: { start: { line: declLine, character: nameStart }, end: { line: declLine, character: nameStart + d.name.name.length } },
            });
        }
    }
}

function addConstDeclaration(
    decl: ConstDeclaration,
    out: Map<string, SymbolEntry>,
    kind: SymbolKind = 'const',
): void {
    if (!decl.loc) return;
    const constName = decl.name.name;
    const constScope: string | undefined = (decl as any).scope;
    const scopePrefix = constScope ? `${cap(constScope)} ` : '';
    const typeStr = decl.objectType ? ` As ${decl.objectType}` : '';
    const valStr  = constLiteralText(decl.value);
    const displayText = `${scopePrefix}Const ${constName}${typeStr}${valStr ? ` = ${valStr}` : ''}`;
    if (decl.name.loc) {
        const ln = decl.name.loc.start.line - 1;
        const col = decl.name.loc.start.column - 1;
        out.set(constName.toLowerCase(), {
            name: constName,
            displayText,
            kind,
            range: { start: { line: ln, character: col }, end: { line: ln, character: col + constName.length } },
        });
    } else {
        const declLine = decl.loc.start.line - 1;
        const declCol  = decl.loc.start.column - 1;
        const nameStart = declCol + (scopePrefix.length + 6); // 'Const '
        out.set(constName.toLowerCase(), {
            name: constName,
            displayText,
            kind,
            range: { start: { line: declLine, character: nameStart }, end: { line: declLine, character: nameStart + constName.length } },
        });
    }
}

function makeProcEntry(proc: ProcedureDeclaration): SymbolEntry {
    const procLine = proc.loc!.start.line - 1;

    let nameRange: { start: { line: number; character: number }; end: { line: number; character: number } };
    if (proc.name.loc) {
        const ln = proc.name.loc.start.line - 1;
        const col = proc.name.loc.start.column - 1;
        nameRange = { start: { line: ln, character: col }, end: { line: ln, character: col + proc.name.name.length } };
    } else {
        const procCol = proc.loc!.start.column - 1;
        let keywordLen = 3;
        if (proc.isFunction) { keywordLen = 8; }
        if (proc.isProperty) {
            keywordLen = 8 + 1 + (proc.propertyType || 'get').length;
        }
        const nameStart = procCol + keywordLen + 1;
        nameRange = { start: { line: procLine, character: nameStart }, end: { line: procLine, character: nameStart + proc.name.name.length } };
    }

    const params = proc.parameters
        .map(p => {
            const pname = p.name as unknown as string;
            const ptype = p.paramType || 'Variant';
            const byref = p.isByVal ? 'ByVal ' : '';
            return `${byref}${pname} As ${ptype}`;
        })
        .join(', ');
    let keyword = 'Sub';
    if (proc.isFunction) keyword = 'Function';
    if (proc.isProperty) keyword = `Property ${proc.propertyType ?? 'Get'}`;
    const scopePrefix = proc.scope ? `${cap(proc.scope)} ` : '';
    let sig = `${scopePrefix}${keyword} ${proc.name.name}(${params})`;
    if (proc.isFunction || proc.isProperty) sig += ` As ${proc.returnType || 'Variant'}`;

    return { name: proc.name.name, displayText: sig, kind: 'procedure', range: nameRange };
}

// ─── legacy flat builder (kept for any remaining callers) ────────────────────

export function buildSymbolTable(statements: Statement[]): Map<string, SymbolEntry> {
    const { moduleSymbols, procedures } = buildScopedSymbolTable(statements);
    // Merge all scopes into a flat map (last write wins — same behaviour as before)
    const flat = new Map<string, SymbolEntry>(moduleSymbols);
    for (const proc of procedures) {
        for (const [k, v] of proc.localSymbols) flat.set(k, v);
    }
    return flat;
}

// ─── scope lookup helper ─────────────────────────────────────────────────────

/**
 * Find the procedure scope that contains (line, character) (0-based).
 * Returns null if the cursor is at module level.
 */
export function findEnclosingScope(procedures: ProcedureScope[], line: number): ProcedureScope | null {
    for (const scope of procedures) {
        if (line >= scope.range.start.line && line <= scope.range.end.line) {
            return scope;
        }
    }
    return null;
}

/**
 * Look up a symbol name respecting scope.
 * Returns the entry from the innermost scope that declares it, or null.
 */
export function lookupSymbol(
    name: string,
    cursorLine: number,
    table: ScopedSymbolTable,
): SymbolEntry | null {
    return lookupSymbolWithContext(name, cursorLine, table)?.entry ?? null;
}

/**
 * Look up a symbol with scope context (for rich hover display).
 * Returns the entry plus the enclosing procedure name (null = module level).
 */
export function lookupSymbolWithContext(
    name: string,
    cursorLine: number,
    table: ScopedSymbolTable,
): SymbolLookupResult | null {
    const lower = name.toLowerCase();
    const enclosing = findEnclosingScope(table.procedures, cursorLine);
    if (enclosing) {
        const local = enclosing.localSymbols.get(lower);
        if (local) return { entry: local, procName: enclosing.name };
    }
    const mod = table.moduleSymbols.get(lower);
    if (mod) return { entry: mod, procName: null };
    return null;
}
