/**
 * VBE-style auto-parentheses and auto-end-block helpers.
 *
 * autoParensEdit:  insert "()" after a Sub/Function/Property name if missing.
 * getBlockEnd:     return end keyword + scan patterns for a block-opening line.
 * needsBodyIndent: return true for block-continuing lines (Else, ElseIf, Case).
 * needsEndBlock:   scan forward to check if a matching end block already exists.
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

/** Describes the closing keyword and scan patterns for a block-opening line. */
export interface BlockEnd {
    /** Text to insert, e.g. "End Sub", "End If", "Next i". */
    insertKeyword: string;
    /** Regex that matches the closing keyword (e.g. /^\s*End\s+Sub\b/i). */
    closePattern: RegExp;
    /** Regex that matches the same kind of opening keyword for depth tracking. */
    openPattern: RegExp;
}

/**
 * Returns a BlockEnd for the given block-opening line, or null if the line
 * does not start a block that needs an explicit end keyword.
 *
 * Supported:
 *   Sub / Function / Property Get|Let|Set  → End Sub / End Function / End Property
 *   For variable = start To end [Step n]   → Next variable
 *   For Each variable In collection        → Next variable
 *   If ... Then           (multi-line)     → End If
 *   With object                            → End With
 *   Do [While|Until cond]                  → Loop
 *   While condition                        → Wend
 *   Select Case expression                 → End Select
 */
export function getBlockEnd(lineText: string): BlockEnd | null {
    const t = lineText.trimStart();

    // Sub / Function / Property
    if (/^(?:(?:Public|Private|Friend|Static|Global)\s+)*Sub\s+\w+/i.test(t))
        return { insertKeyword: 'End Sub', closePattern: /^\s*End\s+Sub\b/i, openPattern: PROC_DECL };
    if (/^(?:(?:Public|Private|Friend|Static|Global)\s+)*Function\s+\w+/i.test(t))
        return { insertKeyword: 'End Function', closePattern: /^\s*End\s+Function\b/i, openPattern: PROC_DECL };
    if (/^(?:(?:Public|Private|Friend|Static|Global)\s+)*Property\s+(?:Get|Let|Set)\s+\w+/i.test(t))
        return { insertKeyword: 'End Property', closePattern: /^\s*End\s+Property\b/i, openPattern: PROC_DECL };

    // For Each variable In collection  (check before For variable = ...)
    const forEachMatch = /^For\s+Each\s+(\w+)\s+In\b/i.exec(t);
    if (forEachMatch)
        return { insertKeyword: `Next ${forEachMatch[1]}`, closePattern: /^\s*Next\b/i, openPattern: /^\s*For\b/i };

    // For variable = start To end [Step step]
    const forMatch = /^For\s+(\w+)\s*=/i.exec(t);
    if (forMatch)
        return { insertKeyword: `Next ${forMatch[1]}`, closePattern: /^\s*Next\b/i, openPattern: /^\s*For\b/i };

    // If ... Then  (multi-line only: nothing after Then except an optional comment)
    if (/^If\b.*\bThen(\s*'.*)?$/i.test(t))
        return { insertKeyword: 'End If', closePattern: /^\s*End\s+If\b/i, openPattern: /^\s*If\b.*\bThen(\s*'.*)?$/i };

    // With object
    if (/^With\b/i.test(t))
        return { insertKeyword: 'End With', closePattern: /^\s*End\s+With\b/i, openPattern: /^\s*With\b/i };

    // Do [While|Until condition]  — must come before standalone While
    if (/^Do\b/i.test(t))
        return { insertKeyword: 'Loop', closePattern: /^\s*Loop\b/i, openPattern: /^\s*Do\b/i };

    // While condition ... Wend  (standalone, not Do While)
    if (/^While\b/i.test(t))
        return { insertKeyword: 'Wend', closePattern: /^\s*Wend\b/i, openPattern: /^\s*While\b/i };

    // Select Case expression
    if (/^Select\s+Case\b/i.test(t))
        return { insertKeyword: 'End Select', closePattern: /^\s*End\s+Select\b/i, openPattern: /^\s*Select\s+Case\b/i };

    return null;
}

/**
 * Returns true when the line is a block-continuing keyword that should cause
 * Enter to indent the next line without inserting a new end keyword.
 *
 * Handled: Else, ElseIf / Else If ... Then, Case (inside Select Case)
 */
export function needsBodyIndent(lineText: string): boolean {
    const t = lineText.trimStart();
    // Bare Else (multi-line form: nothing after Else except optional comment)
    if (/^Else\s*('.*)?$/i.test(t)) return true;
    // ElseIf or Else If ... Then
    if (/^Else\s*If\b.*\bThen(\s*'.*)?$/i.test(t)) return true;
    // Case x / Case Else / Case Is ... (inside Select Case)
    if (/^Case\b/i.test(t)) return true;
    return false;
}

/**
 * Scan forward from `startLine` (using `getLine`) to determine whether
 * there is already a matching end block.
 *
 * Returns true  → end block NOT found → caller should insert one.
 * Returns false → end block already present → no action needed.
 *
 * Nested blocks of the same kind (matched by `openPattern`) increment a depth
 * counter so their own close keywords don't satisfy the outer block.
 */
export function needsEndBlock(
    getLine: (n: number) => string | undefined,
    startLine: number,
    closePattern: RegExp,
    openPattern: RegExp,
    maxLines = 200,
): boolean {
    let depth = 0;
    for (let i = startLine; i < startLine + maxLines; i++) {
        const line = getLine(i);
        if (line === undefined) break;
        const trimmed = line.trimStart();
        if (!trimmed || trimmed.startsWith("'")) continue;
        if (openPattern.test(trimmed)) {
            depth++;
        } else if (closePattern.test(trimmed)) {
            if (depth === 0) return false; // found matching end
            depth--;
        }
    }
    return true; // no matching end found within scan range
}
