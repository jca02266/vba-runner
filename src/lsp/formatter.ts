import { Lexer, Token, TokenType } from '../engine/lexer';
import { TextEdit } from './rename-provider';

export interface FormatterOptions {
    indentSize?: number;
    indentChar?: string;
    /** 'pascal' = VBA standard casing, false = skip rule */
    keywordCase?: 'pascal' | false;
}

// Canonical Pascal-case for each keyword token type
const KEYWORD_CANONICAL = new Map<TokenType, string>([
    [TokenType.KeywordDim, 'Dim'],
    [TokenType.KeywordSub, 'Sub'],
    [TokenType.KeywordFunction, 'Function'],
    [TokenType.KeywordProperty, 'Property'],
    [TokenType.KeywordGet, 'Get'],
    [TokenType.KeywordLet, 'Let'],
    [TokenType.KeywordSet, 'Set'],
    [TokenType.KeywordEnd, 'End'],
    [TokenType.KeywordIf, 'If'],
    [TokenType.KeywordThen, 'Then'],
    [TokenType.KeywordElse, 'Else'],
    [TokenType.KeywordElseIf, 'ElseIf'],
    [TokenType.KeywordFor, 'For'],
    [TokenType.KeywordTo, 'To'],
    [TokenType.KeywordStep, 'Step'],
    [TokenType.KeywordNext, 'Next'],
    [TokenType.KeywordEach, 'Each'],
    [TokenType.KeywordIn, 'In'],
    [TokenType.KeywordDo, 'Do'],
    [TokenType.KeywordLoop, 'Loop'],
    [TokenType.KeywordWhile, 'While'],
    [TokenType.KeywordWend, 'Wend'],
    [TokenType.KeywordUntil, 'Until'],
    [TokenType.KeywordWith, 'With'],
    [TokenType.KeywordSelect, 'Select'],
    [TokenType.KeywordCase, 'Case'],
    [TokenType.KeywordType, 'Type'],
    [TokenType.KeywordEnum, 'Enum'],
    [TokenType.KeywordClass, 'Class'],
    [TokenType.KeywordPublic, 'Public'],
    [TokenType.KeywordPrivate, 'Private'],
    [TokenType.KeywordFriend, 'Friend'],
    [TokenType.KeywordStatic, 'Static'],
    [TokenType.KeywordConst, 'Const'],
    [TokenType.KeywordAs, 'As'],
    [TokenType.KeywordNew, 'New'],
    [TokenType.KeywordByVal, 'ByVal'],
    [TokenType.KeywordByRef, 'ByRef'],
    [TokenType.KeywordOptional, 'Optional'],
    [TokenType.KeywordParamArray, 'ParamArray'],
    [TokenType.KeywordCall, 'Call'],
    [TokenType.KeywordExit, 'Exit'],
    [TokenType.KeywordReturn, 'Return'],
    [TokenType.KeywordGoTo, 'GoTo'],
    [TokenType.KeywordGoSub, 'GoSub'],
    [TokenType.KeywordOn, 'On'],
    [TokenType.KeywordError, 'Error'],
    [TokenType.KeywordResume, 'Resume'],
    [TokenType.KeywordReDim, 'ReDim'],
    [TokenType.KeywordErase, 'Erase'],
    [TokenType.KeywordNot, 'Not'],
    [TokenType.KeywordAnd, 'And'],
    [TokenType.KeywordOr, 'Or'],
    [TokenType.KeywordXor, 'Xor'],
    [TokenType.KeywordEqv, 'Eqv'],
    [TokenType.KeywordImp, 'Imp'],
    [TokenType.KeywordMod, 'Mod'],
    [TokenType.KeywordIs, 'Is'],
    [TokenType.KeywordLike, 'Like'],
    [TokenType.KeywordNothing, 'Nothing'],
    [TokenType.KeywordEmpty, 'Empty'],
    [TokenType.KeywordNull, 'Null'],
    [TokenType.KeywordMe, 'Me'],
    [TokenType.KeywordOption, 'Option'],
    [TokenType.KeywordExplicit, 'Explicit'],
    [TokenType.KeywordBase, 'Base'],
    [TokenType.KeywordCompare, 'Compare'],
    [TokenType.KeywordStop, 'Stop'],
    [TokenType.KeywordOpen, 'Open'],
    [TokenType.KeywordClose, 'Close'],
    [TokenType.KeywordPrint, 'Print'],
    [TokenType.KeywordInput, 'Input'],
    [TokenType.KeywordOutput, 'Output'],
    [TokenType.KeywordAppend, 'Append'],
    [TokenType.KeywordRandom, 'Random'],
    [TokenType.KeywordAccess, 'Access'],
    [TokenType.KeywordRead, 'Read'],
    [TokenType.KeywordWrite, 'Write'],
    [TokenType.KeywordLock, 'Lock'],
    [TokenType.KeywordShared, 'Shared'],
    [TokenType.KeywordDeclare, 'Declare'],
    [TokenType.KeywordLib, 'Lib'],
    [TokenType.KeywordAlias, 'Alias'],
    [TokenType.KeywordPtrSafe, 'PtrSafe'],
    [TokenType.KeywordImplements, 'Implements'],
    [TokenType.KeywordEvent, 'Event'],
    [TokenType.KeywordRaiseEvent, 'RaiseEvent'],
    [TokenType.KeywordWithEvents, 'WithEvents'],
    [TokenType.KeywordTypeOf, 'TypeOf'],
    [TokenType.KeywordAddressOf, 'AddressOf'],
    [TokenType.KeywordCollection, 'Collection'],
    [TokenType.KeywordLSet, 'LSet'],
    [TokenType.KeywordRSet, 'RSet'],
    [TokenType.KeywordMid, 'Mid'],
    [TokenType.KeywordKill, 'Kill'],
    [TokenType.KeywordSeek, 'Seek'],
    [TokenType.KeywordReset, 'Reset'],
    [TokenType.KeywordUnlock, 'Unlock'],
    [TokenType.KeywordWidth, 'Width'],
    [TokenType.KeywordLine, 'Line'],
    [TokenType.KeywordPut, 'Put'],
    [TokenType.KeywordAttribute, 'Attribute'],
    [TokenType.KeywordModule, 'Module'],
    [TokenType.KeywordSpc, 'Spc'],
    [TokenType.KeywordTab, 'Tab'],
    [TokenType.KeywordAppActivate, 'AppActivate'],
    [TokenType.KeywordSendKeys, 'SendKeys'],
    [TokenType.KeywordBinary, 'Binary'],
    [TokenType.KeywordText, 'Text'],
]);

interface TokenChange {
    startCol: number; // 0-based column in original line
    endCol: number;
    text: string;
}

// --- identifier definition collector ---------------------------------------

const PARAM_INTRO = new Set([
    TokenType.OperatorLParen,
    TokenType.OperatorComma,
    TokenType.KeywordByRef,
    TokenType.KeywordByVal,
    TokenType.KeywordOptional,
    TokenType.KeywordParamArray,
]);

function collectParamNames(tokens: Token[], startIdx: number, define: (name: string) => void) {
    // Find opening paren
    let i = startIdx;
    while (i < tokens.length && tokens[i].type !== TokenType.OperatorLParen) {
        if (i > startIdx + 2) return;
        i++;
    }
    i++; // skip (
    let depth = 1;
    while (i < tokens.length && depth > 0) {
        const t = tokens[i];
        if (t.type === TokenType.OperatorLParen)  { depth++; i++; continue; }
        if (t.type === TokenType.OperatorRParen)  { depth--; i++; continue; }
        if (depth === 1 && t.type === TokenType.Identifier) {
            const prev = tokens[i - 1];
            if (prev && PARAM_INTRO.has(prev.type)) define(t.value);
        }
        i++;
    }
}

interface ProcScope {
    startLine: number;
    endLine: number;
    defs: Map<string, string>; // proc-local: params + Dim/Const inside the body
}

interface ScopedDefinitions {
    moduleDefs: Map<string, string>; // proc names, module-level Dim/Const, Type/Enum members
    procs: ProcScope[];
}

/** Scan tokens for declaration sites, split into module-level and per-procedure scopes. */
function collectDefinitions(source: string): ScopedDefinitions {
    const moduleDefs = new Map<string, string>();
    const procs: ProcScope[] = [];
    let tokens: Token[];
    try {
        tokens = new Lexer(source).tokenize()
            .filter(t => t.type !== TokenType.Newline && t.type !== TokenType.EOF);
    } catch {
        return { moduleDefs, procs };
    }

    const defModule = (name: string) => {
        if (!moduleDefs.has(name.toLowerCase())) moduleDefs.set(name.toLowerCase(), name);
    };

    let currentProc: { startLine: number; defs: Map<string, string> } | null = null;
    const defLocal = (name: string) => {
        const map = currentProc?.defs ?? moduleDefs;
        if (!map.has(name.toLowerCase())) map.set(name.toLowerCase(), name);
    };

    let inTypeOrEnum = false;
    let prevLine = -1;

    for (let i = 0; i < tokens.length; i++) {
        const t = tokens[i];
        const next = tokens[i + 1];

        // Inside Type/Enum block: first Identifier on each line is a member name
        if (inTypeOrEnum) {
            if (t.type === TokenType.KeywordEnd) {
                if (next?.type === TokenType.KeywordType || next?.type === TokenType.KeywordEnum) {
                    inTypeOrEnum = false;
                }
                prevLine = t.line;
                continue;
            }
            if (t.type === TokenType.Identifier && t.line !== prevLine) {
                defModule(t.value);
            }
            prevLine = t.line;
            continue;
        }

        // Skip access modifiers
        if (ACCESS_MODIFIERS.has(t.type)) continue;

        // End Sub/Function/Property → close current procedure scope
        if (t.type === TokenType.KeywordEnd) {
            if (next?.type === TokenType.KeywordSub ||
                next?.type === TokenType.KeywordFunction ||
                next?.type === TokenType.KeywordProperty) {
                if (currentProc) {
                    procs.push({ startLine: currentProc.startLine, endLine: next.line, defs: currentProc.defs });
                    currentProc = null;
                }
            }
            if (next?.type === TokenType.KeywordType || next?.type === TokenType.KeywordEnum) {
                inTypeOrEnum = false;
            }
            continue;
        }

        switch (t.type) {
            case TokenType.KeywordDim:
            case TokenType.KeywordConst:
            case TokenType.KeywordStatic: {
                if (next?.type === TokenType.Identifier) defLocal(next.value);
                break;
            }
            case TokenType.KeywordReDim: {
                const skip = next?.value?.toLowerCase() === 'preserve' ? tokens[i + 2] : next;
                if (skip?.type === TokenType.Identifier) defLocal(skip.value);
                break;
            }
            case TokenType.KeywordSub:
            case TokenType.KeywordFunction: {
                if (next?.type === TokenType.Identifier) {
                    defModule(next.value); // proc name belongs to module scope
                    currentProc = { startLine: t.line, defs: new Map() };
                    collectParamNames(tokens, i + 2, name => defLocal(name));
                }
                break;
            }
            case TokenType.KeywordProperty: {
                let nameIdx = i + 1;
                const verb = tokens[nameIdx]?.value?.toLowerCase();
                if (verb === 'get' || verb === 'let' || verb === 'set') nameIdx++;
                if (tokens[nameIdx]?.type === TokenType.Identifier) {
                    defModule(tokens[nameIdx].value);
                    currentProc = { startLine: t.line, defs: new Map() };
                    collectParamNames(tokens, nameIdx + 1, name => defLocal(name));
                }
                break;
            }
            case TokenType.KeywordType:
            case TokenType.KeywordEnum: {
                if (next?.type === TokenType.Identifier) defModule(next.value);
                inTypeOrEnum = true;
                break;
            }
        }
    }

    return { moduleDefs, procs };
}

// ---------------------------------------------------------------------------

export function format(source: string, options: FormatterOptions = {}): TextEdit[] {
    const indentSize = options.indentSize ?? 4;
    const indentChar = options.indentChar ?? ' ';
    const doKeywordCase = options.keywordCase !== false;
    const unit = indentChar.repeat(indentSize);

    const physicalLines = source.split('\n');

    // Phase 1: collect keyword-case and identifier-case changes per line index (0-based)
    const scopedDefs = collectDefinitions(source);
    const keywordChangesByLine = new Map<number, TokenChange[]>();
    {
        let tokens: Token[];
        try {
            tokens = new Lexer(source).tokenize();
        } catch {
            tokens = [];
        }
        for (const token of tokens) {
            let canonical: string | undefined;
            if (doKeywordCase) {
                canonical = KEYWORD_CANONICAL.get(token.type);
            }
            if (!canonical && token.type === TokenType.Identifier) {
                const lv = (token.value as string).toLowerCase();
                // Look up in the innermost scope that contains this line, then fall back to module
                const proc = scopedDefs.procs.find(
                    p => token.line >= p.startLine && token.line <= p.endLine
                );
                canonical = proc
                    ? (proc.defs.get(lv) ?? scopedDefs.moduleDefs.get(lv))
                    : scopedDefs.moduleDefs.get(lv);
            }
            if (!canonical) continue;
            const value = token.value as string;
            if (value === canonical) continue;
            const lineIdx = token.line - 1;
            if (!keywordChangesByLine.has(lineIdx)) keywordChangesByLine.set(lineIdx, []);
            keywordChangesByLine.get(lineIdx)!.push({
                startCol: token.column - 1,
                endCol: token.column - 1 + value.length,
                text: canonical,
            });
        }
    }

    // Phase 2: process line by line for indentation + combine keyword changes
    const edits: TextEdit[] = [];
    let level = 0;
    let selectDepth = 0;
    // Accumulate physical lines that form one logical statement (line continuations)
    let continuationAccum: string[] = [];

    for (let i = 0; i < physicalLines.length; i++) {
        const rawLine = physicalLines[i];
        const hasCR = rawLine.endsWith('\r');
        const line = hasCR ? rawLine.slice(0, -1) : rawLine;
        const trimmed = line.trim();

        // Empty lines: skip indentation (don't generate edits for blank lines)
        if (trimmed === '') continue;

        // Continuation lines: preserve their indentation
        const prevLine = i > 0 ? physicalLines[i - 1].replace(/\r$/, '').trimEnd() : '';
        const isContinuation = prevLine.endsWith(' _') || prevLine.endsWith('\t_');

        // Track accumulated lines for block-structure analysis across continuations
        if (!isContinuation) {
            continuationAccum = [trimmed];
        } else {
            continuationAccum.push(trimmed);
        }
        const isLastContinuation = isContinuation &&
            !trimmed.endsWith(' _') && !trimmed.endsWith('\t_');

        // Tokenize for block structure analysis
        const lineTokens = tokenizeLine(trimmed);
        const first = lineTokens[0]?.type;
        const second = lineTokens[1]?.type;

        // Label lines (e.g. `errHandler:`) are always at column 0 per VBA convention
        const isLabel = lineTokens.length >= 2 &&
            lineTokens[0].type === TokenType.Identifier &&
            lineTokens[1].type === TokenType.OperatorColon;

        let expectedLevel = level;

        if (!isContinuation && !isLabel) {
            const startsMultiLine = trimmed.endsWith(' _') || trimmed.endsWith('\t_');
            const { closerBefore, openerAfter } = classifyLine(lineTokens, selectDepth);

            if (closerBefore) {
                level = Math.max(0, level - 1);
                if (first === TokenType.KeywordEnd && second === TokenType.KeywordSelect) {
                    selectDepth = Math.max(0, selectDepth - 1);
                }
            }
            expectedLevel = level;

            // For multi-line statements, defer opener detection to the last continuation line
            if (openerAfter && !startsMultiLine) {
                level++;
                if (first === TokenType.KeywordSelect && second === TokenType.KeywordCase) {
                    selectDepth++;
                }
            }
        } else if (isLastContinuation && !isLabel) {
            // Classify the complete logical line to detect block openers.
            // closerBefore was already applied at the first physical line of this statement,
            // so only openerAfter is processed here to avoid double-applying closerBefore
            // (e.g. ElseIf cond _ / And cond2 Then).
            const logicalLine = continuationAccum
                .map(l => l.replace(/[ \t]+_[ \t]*$/, ''))
                .join(' ');
            const logTokens = tokenizeLine(logicalLine);
            const logFirst = logTokens[0]?.type;
            const logSecond = logTokens[1]?.type;
            const { openerAfter } = classifyLine(logTokens, selectDepth);

            // expectedLevel stays as-is (continuation line keeps its existing indent)
            if (openerAfter) {
                level++;
                if (logFirst === TokenType.KeywordSelect && logSecond === TokenType.KeywordCase) {
                    selectDepth++;
                }
            }
        }

        // Build the new line: expected indent + keyword-fixed content
        const expectedIndent = isContinuation ? '' : isLabel ? '' : unit.repeat(expectedLevel);
        const actualIndent = leadingWhitespace(line);
        const contentStart = actualIndent.length;

        // Apply keyword changes to the content (positions are in the original full line)
        const changes = keywordChangesByLine.get(i) ?? [];
        let content = line.slice(contentStart);

        if (changes.length > 0) {
            // Adjust positions: subtract contentStart offset
            const contentChanges = changes
                .map(c => ({ start: c.startCol - contentStart, end: c.endCol - contentStart, text: c.text }))
                .filter(c => c.start >= 0)
                .sort((a, b) => b.start - a.start); // reverse order for safe replacement

            for (const c of contentChanges) {
                content = content.slice(0, c.start) + c.text + content.slice(c.end);
            }
        }

        const newLine = (isContinuation ? actualIndent : expectedIndent) + content + (hasCR ? '\r' : '');

        if (newLine !== rawLine) {
            edits.push({
                range: {
                    start: { line: i, character: 0 },
                    end: { line: i, character: rawLine.length },
                },
                newText: newLine,
            });
        }
    }

    return edits;
}

// --- block structure classification -----------------------------------------

const BLOCK_OPEN = new Set([
    TokenType.KeywordSub,
    TokenType.KeywordFunction,
    TokenType.KeywordProperty,
    TokenType.KeywordFor,
    TokenType.KeywordDo,
    TokenType.KeywordWhile,
    TokenType.KeywordWith,
    TokenType.KeywordType,
    TokenType.KeywordEnum,
    TokenType.KeywordClass,
]);

const BLOCK_CLOSE_SINGLE = new Set([
    TokenType.KeywordNext,
    TokenType.KeywordWend,
    TokenType.KeywordLoop,
]);

interface LineKind {
    closerBefore: boolean;
    openerAfter: boolean;
}

const ACCESS_MODIFIERS = new Set([
    TokenType.KeywordPublic,
    TokenType.KeywordPrivate,
    TokenType.KeywordFriend,
    TokenType.KeywordStatic,
]);

function classifyLine(tokens: Token[], selectDepth: number): LineKind {
    if (tokens.length === 0) return { closerBefore: false, openerAfter: false };

    // Skip leading access modifiers (Public/Private/Friend/Static) to find the keyword
    let offset = 0;
    while (offset < tokens.length && ACCESS_MODIFIERS.has(tokens[offset].type)) {
        offset++;
    }
    const first = tokens[offset]?.type ?? tokens[0].type;
    const second = tokens[offset + 1]?.type;

    // End <X>
    if (first === TokenType.KeywordEnd && second !== undefined) {
        const isEndProc = second === TokenType.KeywordSub ||
                          second === TokenType.KeywordFunction ||
                          second === TokenType.KeywordProperty;
        const isEndBlock = second === TokenType.KeywordIf ||
                           second === TokenType.KeywordWith ||
                           second === TokenType.KeywordType ||
                           second === TokenType.KeywordEnum ||
                           second === TokenType.KeywordClass ||
                           second === TokenType.KeywordSelect;
        if (isEndProc || isEndBlock) {
            return { closerBefore: true, openerAfter: false };
        }
    }

    if (BLOCK_CLOSE_SINGLE.has(first)) {
        return { closerBefore: true, openerAfter: false };
    }

    // Else / ElseIf: dedent then reindent
    if (first === TokenType.KeywordElse || first === TokenType.KeywordElseIf) {
        return { closerBefore: true, openerAfter: true };
    }

    // Case (inside Select Case): dedent then reindent
    if (first === TokenType.KeywordCase && selectDepth > 0) {
        return { closerBefore: true, openerAfter: true };
    }

    // Block openers
    if (BLOCK_OPEN.has(first)) {
        return { closerBefore: false, openerAfter: true };
    }

    // Select Case opener
    if (first === TokenType.KeywordSelect && second === TokenType.KeywordCase) {
        return { closerBefore: false, openerAfter: true };
    }

    // If ... Then (block form: nothing after Then, ignoring trailing `:` separators)
    if (first === TokenType.KeywordIf) {
        const thenIdx = tokens.findIndex(t => t.type === TokenType.KeywordThen);
        if (thenIdx !== -1) {
            const afterThen = tokens.slice(thenIdx + 1).filter(t => t.type !== TokenType.OperatorColon);
            if (afterThen.length === 0) {
                return { closerBefore: false, openerAfter: true };
            }
        }
        return { closerBefore: false, openerAfter: false };
    }

    return { closerBefore: false, openerAfter: false };
}

function tokenizeLine(trimmed: string): Token[] {
    try {
        return new Lexer(trimmed).tokenize().filter(
            t => t.type !== TokenType.Newline && t.type !== TokenType.EOF
        );
    } catch {
        return [];
    }
}

function leadingWhitespace(line: string): string {
    const match = line.match(/^[ \t]*/);
    return match ? match[0] : '';
}

// Apply TextEdit list to source string (one edit per line, full-line replacement).
export function applyEdits(source: string, edits: TextEdit[]): string {
    if (edits.length === 0) return source;
    const lines = source.split('\n');
    const sorted = [...edits].sort((a, b) => b.range.start.line - a.range.start.line);
    for (const edit of sorted) {
        const lineIdx = edit.range.start.line;
        const origLine = lines[lineIdx];
        lines[lineIdx] =
            origLine.slice(0, edit.range.start.character) +
            edit.newText +
            origLine.slice(edit.range.end.character);
    }
    return lines.join('\n');
}
