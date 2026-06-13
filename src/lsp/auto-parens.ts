/**
 * VBE-style auto-parentheses: when the user presses Enter after a Sub/Function/Property
 * declaration that lacks parentheses, insert "()" after the procedure name.
 *
 * Examples:
 *   "Sub Foo"              → "Sub Foo()"
 *   "Private Function Bar" → "Private Function Bar()"
 *   "Property Get Value"   → "Property Get Value()"
 *   "Sub Foo()"            → no change (already has parens)
 *   "Sub Foo(x As Long)"   → no change (already has parens)
 */

/** Pattern for a proc declaration line that is missing its parameter list. */
const PROC_NO_PARENS = /^(\s*(?:(?:Public|Private|Friend|Static|Global)\s+)*(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+\w+)\s*$/i;

/**
 * Returns a { lineIndex, insertCol } describing where to insert "()"
 * on `lineText`, or null if the line doesn't need it.
 */
export function autoParensEdit(lineText: string): { insertCol: number } | null {
    const m = PROC_NO_PARENS.exec(lineText);
    if (!m) return null;
    const insertCol = m[1].trimEnd().length; // right after the procedure name
    return { insertCol };
}
