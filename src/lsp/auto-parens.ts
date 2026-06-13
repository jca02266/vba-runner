/**
 * VBE-style auto-parentheses and auto-end-block helpers.
 *
 * autoParensEdit: insert "()" after a Sub/Function/Property name if missing.
 * getEndKeyword:  return "End Sub" / "End Function" / "End Property" for a declaration line.
 * needsEndBlock:  scan forward to check if a matching End block already exists.
 */

/** Pattern for a proc declaration line that is missing its parameter list. */
const PROC_NO_PARENS = /^(\s*(?:(?:Public|Private|Friend|Static|Global)\s+)*(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+\w+)\s*$/i;

/** Pattern to detect any proc declaration (with or without parens). */
const PROC_DECL = /^\s*(?:(?:Public|Private|Friend|Static|Global)\s+)*(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+\w+/i;

/**
 * Returns a { insertCol } describing where to insert "()" on `lineText`,
 * or null if the line already has parentheses or is not a proc declaration.
 */
export function autoParensEdit(lineText: string): { insertCol: number } | null {
    const m = PROC_NO_PARENS.exec(lineText);
    if (!m) return null;
    const insertCol = m[1].trimEnd().length; // right after the procedure name
    return { insertCol };
}

/**
 * Returns the matching end keyword for a proc declaration line:
 *   "End Sub" | "End Function" | "End Property" | null
 */
export function getEndKeyword(lineText: string): string | null {
    const t = lineText.trimStart();
    if (/^(?:(?:Public|Private|Friend|Static|Global)\s+)*Sub\s+\w+/i.test(t)) return 'End Sub';
    if (/^(?:(?:Public|Private|Friend|Static|Global)\s+)*Function\s+\w+/i.test(t)) return 'End Function';
    if (/^(?:(?:Public|Private|Friend|Static|Global)\s+)*Property\s+(?:Get|Let|Set)\s+\w+/i.test(t)) return 'End Property';
    return null;
}

/**
 * Scan forward from `startLine` (using `getLine`) to determine whether
 * there is already a matching End block for `endKeyword`.
 *
 * Returns true  → End block NOT found → caller should insert one.
 * Returns false → End block already present → no action needed.
 *
 * Nested Sub/Function/Property declarations increment a depth counter so
 * their own End blocks don't satisfy ours.
 */
export function needsEndBlock(
    getLine: (n: number) => string | undefined,
    startLine: number,
    endKeyword: string,
    maxLines = 200,
): boolean {
    const endRe = new RegExp(`^\\s*${endKeyword.replace(' ', '\\s+')}\\b`, 'i');
    let depth = 0;
    for (let i = startLine; i < startLine + maxLines; i++) {
        const line = getLine(i);
        if (line === undefined) break;
        const trimmed = line.trimStart();
        if (!trimmed || trimmed.startsWith("'")) continue;
        if (PROC_DECL.test(trimmed)) {
            depth++;
        } else if (endRe.test(line)) {
            if (depth === 0) return false; // found matching end
            depth--;
        }
    }
    return true; // no matching end found within scan range
}
