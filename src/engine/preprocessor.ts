export type CompilerConstants = Record<string, number | string | boolean>;

export const DEFAULT_COMPILER_CONSTANTS: CompilerConstants = {
    VBA7: 0,
    Win64: 0,
    Win32: -1,
    Mac: 0,
};

export function preprocess(source: string, constants: CompilerConstants = {}): string {
    const merged: Record<string, any> = {};
    for (const k of Object.keys(DEFAULT_COMPILER_CONSTANTS)) {
        merged[k.toLowerCase()] = DEFAULT_COMPILER_CONSTANTS[k];
    }
    for (const k of Object.keys(constants)) {
        merged[k.toLowerCase()] = constants[k];
    }
    const localConsts: Record<string, any> = {};

    const resolve = (name: string): any => {
        const key = name.toLowerCase();
        if (key in localConsts) return localConsts[key];
        if (key in merged) return merged[key];
        return 0;
    };

    const lines = source.split(/\r?\n/);
    const result: string[] = [];

    type IfState = { included: boolean; done: boolean };
    const stack: IfState[] = [];

    const isIncluded = (): boolean => stack.every(s => s.included);

    for (const line of lines) {
        const trimmed = line.trimStart();

        // Match cc directives: optional whitespace then #keyword
        const ccMatch = trimmed.match(/^#\s*([a-zA-Z]+)(.*)/);
        if (ccMatch) {
            const directive = ccMatch[1].toLowerCase();
            const rest = ccMatch[2].trim();

            if (directive === 'const') {
                const m = rest.match(/^(\w+)\s*=\s*(.+)$/);
                if (m) {
                    localConsts[m[1].toLowerCase()] = evaluateCCExpr(m[2], resolve);
                }
                result.push('');
                continue;
            }

            if (directive === 'if') {
                const m = rest.match(/^(.*?)\s+then\s*(?:'.*)?$/i);
                const cond = m ? isTruthy(evaluateCCExpr(m[1], resolve)) : false;
                const parentOk = isIncluded();
                stack.push({ included: parentOk && cond, done: parentOk && cond });
                result.push('');
                continue;
            }

            if (directive === 'elseif') {
                if (stack.length > 0) {
                    const state = stack[stack.length - 1];
                    const parentOk = stack.slice(0, -1).every(s => s.included);
                    if (!state.done && parentOk) {
                        const m = rest.match(/^(.*?)\s+then\s*(?:'.*)?$/i);
                        const cond = m ? isTruthy(evaluateCCExpr(m[1], resolve)) : false;
                        state.included = cond;
                        if (cond) state.done = true;
                    } else {
                        state.included = false;
                    }
                    result.push('');
                    continue;
                }
            }

            if (directive === 'else') {
                if (stack.length > 0) {
                    const state = stack[stack.length - 1];
                    const parentOk = stack.slice(0, -1).every(s => s.included);
                    state.included = parentOk && !state.done;
                    result.push('');
                    continue;
                }
            }

            if (directive === 'endif' || (directive === 'end' && rest.toLowerCase().startsWith('if'))) {
                if (stack.length > 0) stack.pop();
                result.push('');
                continue;
            }
        }

        result.push(isIncluded() ? line : '');
    }

    return result.join('\n');
}

function isTruthy(value: any): boolean {
    if (typeof value === 'boolean') return value;
    if (typeof value === 'number') return value !== 0;
    if (typeof value === 'string') return value !== '';
    return Boolean(value);
}

function tokenizeCC(expr: string): string[] {
    const tokens: string[] = [];
    let i = 0;
    while (i < expr.length) {
        if (/\s/.test(expr[i])) { i++; continue; }
        if (expr[i] === '"') {
            let j = i + 1;
            while (j < expr.length && expr[j] !== '"') j++;
            tokens.push(expr.slice(i, j + 1));
            i = j + 1;
            continue;
        }
        if (/\d/.test(expr[i])) {
            let j = i;
            while (j < expr.length && /[\d.]/.test(expr[j])) j++;
            tokens.push(expr.slice(i, j));
            i = j;
            continue;
        }
        if (/[a-zA-Z_]/.test(expr[i])) {
            let j = i;
            while (j < expr.length && /\w/.test(expr[j])) j++;
            tokens.push(expr.slice(i, j));
            i = j;
            continue;
        }
        if (i + 1 < expr.length && ['<>', '<=', '>='].includes(expr.slice(i, i + 2))) {
            tokens.push(expr.slice(i, i + 2));
            i += 2;
            continue;
        }
        tokens.push(expr[i]);
        i++;
    }
    return tokens;
}

function evaluateCCExpr(expr: string, resolve: (name: string) => any): any {
    const tokens = tokenizeCC(expr.trim());
    let pos = 0;

    const peek = () => tokens[pos] ?? '';
    const consume = () => tokens[pos++] ?? '';

    function parseOr(): any {
        let left = parseAnd();
        while (peek().toLowerCase() === 'or') {
            consume();
            const right = parseAnd();
            left = (isTruthy(left) || isTruthy(right)) ? -1 : 0;
        }
        return left;
    }

    function parseAnd(): any {
        let left = parseNot();
        while (peek().toLowerCase() === 'and') {
            consume();
            const right = parseNot();
            left = (isTruthy(left) && isTruthy(right)) ? -1 : 0;
        }
        return left;
    }

    function parseNot(): any {
        if (peek().toLowerCase() === 'not') {
            consume();
            return isTruthy(parseNot()) ? 0 : -1;
        }
        return parseComparison();
    }

    function parseComparison(): any {
        const left = parseAddSub();
        const op = peek();
        if (['=', '<>', '<', '>', '<=', '>='].includes(op)) {
            consume();
            const right = parseAddSub();
            switch (op) {
                case '=':  return left == right ? -1 : 0;
                case '<>': return left != right ? -1 : 0;
                case '<':  return left < right ? -1 : 0;
                case '>':  return left > right ? -1 : 0;
                case '<=': return left <= right ? -1 : 0;
                case '>=': return left >= right ? -1 : 0;
            }
        }
        return left;
    }

    function parseAddSub(): any {
        let left = parseMulDiv();
        while (peek() === '+' || peek() === '-') {
            const op = consume();
            const right = parseMulDiv();
            left = op === '+' ? Number(left) + Number(right) : Number(left) - Number(right);
        }
        return left;
    }

    function parseMulDiv(): any {
        let left = parseUnary();
        while (peek() === '*' || peek() === '/') {
            const op = consume();
            const right = parseUnary();
            left = op === '*' ? Number(left) * Number(right) : Number(left) / Number(right);
        }
        return left;
    }

    function parseUnary(): any {
        if (peek() === '-') { consume(); return -Number(parseAtom()); }
        if (peek() === '+') { consume(); return Number(parseAtom()); }
        return parseAtom();
    }

    function parseAtom(): any {
        const tok = peek();
        if (tok === '(') {
            consume();
            const val = parseOr();
            if (peek() === ')') consume();
            return val;
        }
        if (tok.toLowerCase() === 'true')  { consume(); return -1; }
        if (tok.toLowerCase() === 'false') { consume(); return 0; }
        if (/^-?\d+$/.test(tok))           { consume(); return parseInt(tok, 10); }
        if (/^\d+(\.\d+)?$/.test(tok))     { consume(); return parseFloat(tok); }
        if (/^".*"$/.test(tok))            { consume(); return tok.slice(1, -1); }
        if (/^[a-zA-Z_]\w*$/.test(tok))    { consume(); return resolve(tok); }
        consume();
        return 0;
    }

    return parseOr();
}

/**
 * Strip the Excel-exported VBA file header from source text.
 *
 * Excel exports VBA files with a header block that is not valid VBA syntax:
 *
 *   .cls  — "VERSION 1.0 CLASS / BEGIN...END"
 *   .frm  — "VERSION 5.00 / Begin {GUID} FormName...End"
 *   .bas  — no VERSION header (starts with Attribute VB_Name, handled by parser)
 *
 * The VERSION...END block is blanked out to preserve line numbers.
 * Attribute statements that follow are valid VBA and are left intact.
 * Nested Begin...End blocks inside .frm controls are indented, so their
 * "End" lines do not match the column-0 "END" that terminates the outer block.
 */
export function stripVBAFileHeader(source: string): string {
    const lines = source.split('\n');
    if (!lines[0]?.trimEnd().toUpperCase().startsWith('VERSION')) return source;
    const result = [...lines];
    let i = 0;
    result[i] = '';
    i++;
    while (i < result.length) {
        const trimmed = result[i].trimEnd().toUpperCase();
        result[i] = '';
        i++;
        if (trimmed === 'END') break;
    }
    return result.join('\n');
}
