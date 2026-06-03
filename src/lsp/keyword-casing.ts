// VBA keywords with their canonical casing, mirroring the VBE auto-correction
// behavior. When a keyword is typed in any other case (e.g. `if`, `WITHEVENTS`),
// it is rewritten to the casing listed here.
const CANONICAL_KEYWORDS: string[] = [
    // Control flow
    'If', 'Then', 'Else', 'ElseIf', 'End', 'Select', 'Case',
    'For', 'Each', 'In', 'To', 'Step', 'Next',
    'While', 'Wend', 'Do', 'Loop', 'Until',
    'With', 'Exit', 'GoTo', 'GoSub', 'Return', 'Resume', 'On', 'Error',
    // Declarations
    'Sub', 'Function', 'Property', 'Get', 'Let', 'Set', 'Call',
    'Class', 'Type', 'Enum',
    'Public', 'Private', 'Friend', 'Global', 'Static',
    'Dim', 'Const', 'ReDim', 'Preserve', 'Erase',
    'Implements', 'Option', 'Explicit', 'Compare', 'Binary', 'Text', 'Module', 'Base',
    'Declare', 'Lib', 'Alias', 'PtrSafe',
    'Event', 'RaiseEvent', 'WithEvents',
    'ByVal', 'ByRef', 'Optional', 'ParamArray',
    // Expressions / operators
    'New', 'Me', 'TypeOf', 'AddressOf',
    'Is', 'Not', 'And', 'Or', 'Xor', 'Mod', 'Like', 'Eqv', 'Imp',
    'As', 'Nothing', 'True', 'False', 'Empty', 'Null',
    // Types
    'Boolean', 'Byte', 'Integer', 'Long', 'LongLong', 'LongPtr',
    'Single', 'Double', 'Currency', 'Decimal', 'Date', 'String', 'Variant', 'Object',
    // File I/O statements
    'Open', 'Close', 'Print', 'Write', 'Input', 'Output', 'Append', 'Random',
    'Lock', 'Unlock',
    // Misc
    'Stop', 'Attribute', 'Rem', 'Collection',
];

const MAP = new Map<string, string>(CANONICAL_KEYWORDS.map((k) => [k.toLowerCase(), k]));

/**
 * Returns the canonical casing if `word` is a VBA keyword written in a
 * different case; otherwise returns undefined (not a keyword, or already
 * canonical).
 */
export function canonicalKeyword(word: string): string | undefined {
    const canonical = MAP.get(word.toLowerCase());
    return canonical && canonical !== word ? canonical : undefined;
}

/**
 * True if column `col` in `line` falls inside a string literal or a line
 * comment. VBA strings cannot span lines, so a single-line scan is sufficient.
 */
export function isInStringOrComment(line: string, col: number): boolean {
    let inString = false;
    for (let i = 0; i < col && i < line.length; i++) {
        const ch = line[i];
        if (ch === '"') {
            inString = !inString;
        } else if (ch === "'" && !inString) {
            return true; // rest of the line is a comment
        }
    }
    return inString;
}
