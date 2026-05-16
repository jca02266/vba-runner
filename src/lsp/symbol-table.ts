import {
    ClassDeclaration,
    EventDeclaration,
    ProcedureDeclaration,
    VariableDeclaration,
    ConstDeclaration,
    Statement,
} from '../compiler/parser';

export interface SymbolEntry {
    name: string;
    displayText: string; // shown in hover tooltip
    range: {
        start: { line: number; character: number };
        end: { line: number; character: number };
    };
}

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

export function buildSymbolTable(statements: Statement[]): Map<string, SymbolEntry> {
    const symbols = new Map<string, SymbolEntry>();
    collectSymbols(statements, symbols);
    return symbols;
}

function collectSymbols(statements: Statement[], symbols: Map<string, SymbolEntry>): void {
    for (const stmt of statements) {
        if (stmt.type === 'ProcedureDeclaration') {
            const proc = stmt as ProcedureDeclaration;
            if (!proc.loc) continue;

            const procLine = proc.loc.start.line - 1;
            const procCol  = proc.loc.start.column - 1;

            let keyword = 'Sub';
            let keywordLen = 3;
            if (proc.isFunction) { keyword = 'Function'; keywordLen = 8; }
            if (proc.isProperty) {
                keyword = 'Property';
                keywordLen = 8;
                const pt = proc.propertyType || 'get';
                keywordLen += 1 + pt.length; // ' Get' etc.
            }
            const nameStart = procCol + keywordLen + 1;
            const nameEnd   = nameStart + proc.name.name.length;

            // Build signature for hover display
            const params = proc.parameters
                .map(p => {
                    const pname = p.name as unknown as string;
                    const ptype = p.paramType || 'Variant';
                    const byref = p.isByVal ? 'ByVal ' : '';
                    return `${byref}${pname} As ${ptype}`;
                })
                .join(', ');
            let sig = `${keyword} ${proc.name.name}(${params})`;
            if (proc.isFunction || proc.isProperty) {
                sig += ' As Variant'; // return type not tracked in AST simply
            }

            symbols.set(proc.name.name.toLowerCase(), {
                name: proc.name.name,
                displayText: sig,
                range: { start: { line: procLine, character: nameStart }, end: { line: procLine, character: nameEnd } },
            });

            // Parameters (name is plain string in AST)
            for (const param of proc.parameters) {
                const pname = param.name as unknown as string;
                if (!pname) continue;
                const ptype = param.paramType || 'Variant';
                symbols.set(pname.toLowerCase(), {
                    name: pname,
                    displayText: `(parameter) ${pname} As ${ptype}`,
                    range: { start: { line: procLine, character: nameStart }, end: { line: procLine, character: nameEnd } },
                });
            }

            if (proc.body) collectSymbols(proc.body, symbols);

        } else if (stmt.type === 'VariableDeclaration') {
            const decl = stmt as VariableDeclaration;
            if (!decl.loc) continue;

            const declLine = decl.loc.start.line - 1;
            const declCol  = decl.loc.start.column - 1;
            const scope = decl.scope;
            const keyword = scope ?? 'Dim';
            const nameStart = declCol + keyword.length + 1;

            for (const d of decl.declarations) {
                const varName = d.name.name;
                const varType = d.objectType || 'Variant';
                symbols.set(varName.toLowerCase(), {
                    name: varName,
                    displayText: `${keyword} ${varName} As ${varType}`,
                    range: { start: { line: declLine, character: nameStart }, end: { line: declLine, character: nameStart + varName.length } },
                });
            }

        } else if (stmt.type === 'ConstDeclaration') {
            const decl = stmt as ConstDeclaration;
            if (!decl.loc) continue;

            const declLine  = decl.loc.start.line - 1;
            const declCol   = decl.loc.start.column - 1;
            const nameStart = declCol + 6; // 'Const '
            const constName = decl.name.name;

            symbols.set(constName.toLowerCase(), {
                name: constName,
                displayText: `Const ${constName}`,
                range: { start: { line: declLine, character: nameStart }, end: { line: declLine, character: nameStart + constName.length } },
            });

        } else if (stmt.type === 'ClassDeclaration') {
            const cls = stmt as ClassDeclaration;
            if (!cls.loc) continue;

            const clsLine   = cls.loc.start.line - 1;
            const nameStart = (cls.loc.start.column - 1) + 6; // 'Class '

            symbols.set(cls.name.toLowerCase(), {
                name: cls.name,
                displayText: `Class ${cls.name}`,
                range: { start: { line: clsLine, character: nameStart }, end: { line: clsLine, character: nameStart + cls.name.length } },
            });

            if (cls.body) collectSymbols(cls.body, symbols);

        } else if (stmt.type === 'EventDeclaration') {
            const evt = stmt as EventDeclaration;
            if (!evt.loc) continue;

            const evtLine   = evt.loc.start.line - 1;
            const nameStart = (evt.loc.start.column - 1) + 6; // 'Event '
            const evtName   = evt.name.name;

            const params = evt.parameters
                .map(p => {
                    const pname = p.name as unknown as string;
                    const ptype = p.paramType || 'Variant';
                    return `${pname} As ${ptype}`;
                })
                .join(', ');

            symbols.set(evtName.toLowerCase(), {
                name: evtName,
                displayText: `Event ${evtName}(${params})`,
                range: { start: { line: evtLine, character: nameStart }, end: { line: evtLine, character: nameStart + evtName.length } },
            });
        }
    }
}
