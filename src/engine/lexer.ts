export const enum TokenType {
    Identifier,
    Number,
    String,
    KeywordFor,
    KeywordTo,
    KeywordNext,
    KeywordIf,
    KeywordThen,
    KeywordElseIf,
    KeywordElse,
    KeywordEnd,
    KeywordDo,
    KeywordWhile,
    KeywordWend,
    KeywordLoop,
    KeywordUntil,
    KeywordStop,
    KeywordSub,
    KeywordFunction,
    KeywordProperty,
    KeywordGet,
    KeywordLet,
    KeywordDim,
    KeywordAs,
    KeywordNew,
    KeywordCollection,
    KeywordAnd,
    KeywordOr,
    KeywordNot,
    KeywordOption,
    KeywordExplicit,
    KeywordConst,
    KeywordSet,
    KeywordCall,
    KeywordOn,
    KeywordError,
    KeywordGoTo,
    KeywordGoSub,
    KeywordReturn,
    KeywordLSet,
    KeywordRSet,
    KeywordErase,
    KeywordReDim,
    KeywordStep,
    KeywordEmpty,
    KeywordExit,
    KeywordByRef,
    KeywordByVal,
    KeywordType,
    KeywordNothing,
    KeywordOptional,
    KeywordIs,
    KeywordResume,
    KeywordSelect,
    KeywordCase,
    KeywordEach,
    KeywordIn,
    KeywordPublic,
    KeywordPrivate,
    KeywordGlobal,
    KeywordEnum,
    KeywordFriend,
    KeywordWith,
    KeywordTypeOf,
    KeywordLike,
    KeywordXor,
    KeywordEqv,
    KeywordImp,
    KeywordNull,
    KeywordStatic,
    KeywordClass,
    KeywordMe,
    KeywordCompare,
    KeywordBinary,
    KeywordText,
    KeywordAttribute,
    KeywordDeclare,
    KeywordLib,
    KeywordAlias,
    KeywordPtrSafe,
    KeywordBase,
    KeywordModule,
    KeywordOpen,
    KeywordClose,
    KeywordLine,
    KeywordInput,
    KeywordPrint,
    KeywordPut,
    KeywordOutput,
    KeywordAppend,
    KeywordRandom,
    KeywordAccess,
    KeywordRead,
    KeywordWrite,
    KeywordLock,
    KeywordShared,
    KeywordSpc,
    KeywordTab,
    KeywordKill,
    KeywordSeek,
    KeywordReset,
    KeywordUnlock,
    KeywordParamArray,
    KeywordEvent,
    KeywordRaiseEvent,
    KeywordWithEvents,
    KeywordImplements,
    KeywordAppActivate,
    KeywordSendKeys,
    KeywordMid,
    KeywordWidth,
    KeywordAddressOf,
    KeywordDef,
    OperatorHash,
    OperatorPlus,
    OperatorMinus,
    OperatorMultiply,
    OperatorDivide,
    OperatorIntDivide,
    OperatorAmpersand,
    KeywordMod,
    OperatorPower,
    OperatorEquals,
    OperatorNotEquals,
    OperatorLessThan,
    OperatorGreaterThan,
    OperatorLessThanOrEqual,
    OperatorGreaterThanOrEqual,
    OperatorComma,
    OperatorLParen,
    OperatorRParen,
    OperatorDot,
    OperatorColon,
    OperatorColonEquals,
    OperatorExclamation,
    OperatorSemicolon,
    Date,
    Newline,
    EOF,
    Unknown,
    ForeignName,
}

/** MS-VBAL §3.3.5.2 lexical identifier categories. */
export type VbaKeywordCategory =
    | 'statement-keyword'
    | 'marker-keyword'
    | 'operator-identifier'
    | 'special-form'
    | 'reserved-type-identifier'
    | 'reserved-name'
    | 'literal-identifier'
    | 'rem-keyword'
    | 'reserved-for-implementation-use'
    | 'future-reserved';

/**
 * Spelling-level classification used by tooling and tests.  The lexer still
 * emits the existing Keyword* token types; this table is deliberately
 * separate from tokenization so contextual/parser compatibility is unchanged.
 */
export const VBA_KEYWORD_CATEGORIES: Readonly<Record<string, readonly VbaKeywordCategory[]>> = {
    // §3.3.5.2 statement-keyword
    call: ['statement-keyword'], case: ['statement-keyword', 'marker-keyword'],
    close: ['statement-keyword'], const: ['statement-keyword'], declare: ['statement-keyword'],
    dim: ['statement-keyword'], do: ['statement-keyword'], else: ['statement-keyword', 'marker-keyword'],
    elseif: ['statement-keyword'], end: ['statement-keyword'], enum: ['statement-keyword'],
    erase: ['statement-keyword'], event: ['statement-keyword'], exit: ['statement-keyword'],
    for: ['statement-keyword'], friend: ['statement-keyword'], function: ['statement-keyword'],
    get: ['statement-keyword'], gosub: ['statement-keyword'], goto: ['statement-keyword'],
    if: ['statement-keyword'], implements: ['statement-keyword'], input: ['statement-keyword', 'special-form'],
    let: ['statement-keyword'], lock: ['statement-keyword'], loop: ['statement-keyword'],
    lset: ['statement-keyword'], next: ['statement-keyword'], on: ['statement-keyword'],
    open: ['statement-keyword'], option: ['statement-keyword'], print: ['statement-keyword'],
    private: ['statement-keyword'], public: ['statement-keyword'], put: ['statement-keyword'],
    raiseevent: ['statement-keyword'], redim: ['statement-keyword'], resume: ['statement-keyword'],
    return: ['statement-keyword'], rset: ['statement-keyword'], seek: ['statement-keyword'],
    select: ['statement-keyword'], set: ['statement-keyword'], static: ['statement-keyword'],
    stop: ['statement-keyword'], sub: ['statement-keyword'], type: ['statement-keyword'],
    unlock: ['statement-keyword'], wend: ['statement-keyword'], while: ['statement-keyword'],
    with: ['statement-keyword'], write: ['statement-keyword', 'marker-keyword'],

    // marker-keyword
    any: ['marker-keyword'], as: ['marker-keyword'], byref: ['marker-keyword'],
    byval: ['marker-keyword'], each: ['marker-keyword'], in: ['marker-keyword'],
    new: ['marker-keyword', 'operator-identifier'], optional: ['marker-keyword'],
    paramarray: ['marker-keyword'], preserve: ['marker-keyword'], shared: ['marker-keyword'],
    spc: ['marker-keyword'], tab: ['marker-keyword'], then: ['marker-keyword'],
    to: ['marker-keyword'], until: ['marker-keyword'], withevents: ['marker-keyword'],

    // operator-identifier
    addressof: ['operator-identifier'], and: ['operator-identifier'], eqv: ['operator-identifier'],
    imp: ['operator-identifier'], is: ['operator-identifier'], like: ['operator-identifier'],
    mod: ['operator-identifier'], not: ['operator-identifier'], or: ['operator-identifier'],
    typeof: ['operator-identifier'], xor: ['operator-identifier'],

    // special-form
    array: ['special-form'], circle: ['special-form'], inputb: ['special-form'],
    lbound: ['special-form'], scale: ['special-form', 'reserved-name'],
    ubound: ['special-form'],

    // reserved-type-identifier
    boolean: ['reserved-type-identifier'], byte: ['reserved-type-identifier'],
    currency: ['reserved-type-identifier'], date: ['reserved-type-identifier', 'reserved-name'],
    double: ['reserved-type-identifier'], integer: ['reserved-type-identifier'],
    long: ['reserved-type-identifier'], longlong: ['reserved-type-identifier'],
    longptr: ['reserved-type-identifier'], single: ['reserved-type-identifier'],
    string: ['reserved-type-identifier', 'reserved-name'], variant: ['reserved-type-identifier'],

    // reserved-name
    abs: ['reserved-name'], cbool: ['reserved-name'], cbyte: ['reserved-name'],
    ccur: ['reserved-name'], cdate: ['reserved-name'], cdbl: ['reserved-name'],
    cdec: ['reserved-name'], cint: ['reserved-name'], clng: ['reserved-name'],
    clnglng: ['reserved-name'], clngptr: ['reserved-name'], csng: ['reserved-name'],
    cstr: ['reserved-name'], cvar: ['reserved-name'], cverr: ['reserved-name'],
    debug: ['reserved-name'], doevents: ['reserved-name'], fix: ['reserved-name'],
    int: ['reserved-name'], len: ['reserved-name'], lenb: ['reserved-name'],
    me: ['reserved-name'], pset: ['reserved-name'], sgn: ['reserved-name'],

    // literal-identifier / rem-keyword
    true: ['literal-identifier'], false: ['literal-identifier'], nothing: ['literal-identifier'],
    empty: ['literal-identifier'], null: ['literal-identifier'], rem: ['rem-keyword'],

    // Dedicated lexer tokens which are contextual in the parser.
    access: [], alias: [], appactivate: [], append: [], attribute: [], base: [],
    binary: [], class: [], collection: [], compare: [], error: [], explicit: [],
    kill: [], lib: [], line: [], mid: [], module: [], output: [], property: [],
    ptrsafe: [], random: [], read: [], reset: [], sendkeys: [], step: [], text: [],
    width: [],
};

/** Spellings accepted as identifiers by Parser.isIdentifier(). */
export const VBA_CONTEXTUAL_KEYWORDS: ReadonlySet<string> = new Set(
    Object.entries(VBA_KEYWORD_CATEGORIES)
        .filter(([, categories]) => categories.length === 0)
        .map(([spelling]) => spelling),
);

export interface Token {
    type: TokenType;
    value: string;
    line: number;
    column: number; // 1-based column of the first character of this token
}

export interface LexerDiagnostic {
    message: string;
    line: number;
    column: number;
}

export class LexError extends Error {
    constructor(message: string, public readonly line: number, public readonly column: number) {
        super(`${message} (line ${line})`);
        this.name = 'LexError';
    }
}

// VBA type-declaration characters (MS-VBAL §3.3.5.2)
// '!' (Single) and '^' (LongLong) are also operators (bang member-access / exponentiation).
// They are consumed as identifier type suffixes only when NOT immediately followed by an
// alphanumeric character or '_' — e.g. 'x!' and 'x^' are type suffixes, but 'dict!Key'
// and 'x^2' keep '!' and '^' as operators.  '$' is identifier-only (String).
const NUMERIC_TYPE_SUFFIXES = new Set(['%', '&', '^', '!', '#', '@']);
const IDENTIFIER_TYPE_SUFFIXES = new Set(['%', '&', '^', '!', '#', '@', '$']);

export class Lexer {
    private input: string = '';
    private pos: number = 0;
    private line: number = 1;
    private column: number = 1;
    public readonly diagnostics: LexerDiagnostic[] = [];

    // MS-VBAL §3.3.5: name-start-character = Unicode-Letter (Lu,Ll,Lt,Lm,Lo,Nl) | "_"
    private static readonly reUnicodeNameStart = /^[\p{L}\p{Nl}]$/u;
    // name-continue-character adds Unicode-Combining-Character (Mn,Mc), Digit (Nd), Connector-Punct (Pc)
    private static readonly reUnicodeNameContinue = /^[\p{Mn}\p{Mc}\p{Nd}\p{Pc}]$/u;

    constructor(input: string) {
        this.input = input;
    }

    private peek(): string {
        if (this.pos >= this.input.length) return '\0';
        return this.input[this.pos];
    }

    private advance(): string {
        if (this.pos >= this.input.length) return '\0';
        const char = this.input[this.pos++];
        if (char === '\n') {
            this.line++;
            this.column = 1;
        } else {
            this.column++;
        }
        return char;
    }

    private isWhitespace(char: string): boolean {
        return char === ' ' || char === '\t' || char === '\r';
    }

    private isAlpha(char: string): boolean {
        if ((char >= 'a' && char <= 'z') || (char >= 'A' && char <= 'Z') || char === '_') return true;
        const code = char.codePointAt(0) ?? 0;
        return code > 0x7F && Lexer.reUnicodeNameStart.test(char);
    }

    private isAlphaNumeric(char: string): boolean {
        if ((char >= 'a' && char <= 'z') || (char >= 'A' && char <= 'Z') || char === '_') return true;
        if (char >= '0' && char <= '9') return true;
        const code = char.codePointAt(0) ?? 0;
        return code > 0x7F && (Lexer.reUnicodeNameStart.test(char) || Lexer.reUnicodeNameContinue.test(char));
    }

    private isDigit(char: string): boolean {
        return char >= '0' && char <= '9';
    }

    private skipWhitespace() {
        while (this.isWhitespace(this.peek())) {
            this.advance();
        }
    }

    public getNextToken(): Token {
        while (true) {
            this.skipWhitespace();
            const startLine = this.line;
            const startColumn = this.column;

            if (this.pos >= this.input.length) {
                return { type: TokenType.EOF, value: '', line: startLine, column: startColumn };
            }

            const char = this.peek();
            
            // Handle Date literal #mm/dd/yyyy#
            if (char === '#') {
                // Peek ahead to see if there's a closing # on the same line
                let foundClosingHash = false;
                for (let j = this.pos + 1; j < this.input.length; j++) {
                    if (this.input[j] === '#') {
                        foundClosingHash = true;
                        break;
                    }
                    if (this.input[j] === '\n') break;
                }

                if (foundClosingHash) {
                    // Check if it really looks like a date (digits, /, -, :, whitespace, AM/PM)
                    let potentialDate = '';
                    let k = this.pos + 1;
                    while (k < this.input.length && this.input[k] !== '#' && this.input[k] !== '\n') {
                        potentialDate += this.input[k];
                        k++;
                    }
                    
                    // Simple heuristic: if it contains letters (other than AM/PM/months), it's probably not a date
                    const dateRegex = /^[0-9\/\-\s:apm,]+$/i;
                    // Month names are also allowed: Jan, Feb, etc.
                    const monthsRegex = /jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec/i;
                    // 日付・時刻の区切り（/ - :）を必須にする。これがないと
                    // `Write #1, #2024/03/15#` の先頭 `#1, #` を日付と誤認する（Bug 32-B）
                    const hasDateSeparator = /[\/\-:]/.test(potentialDate);

                    if ((dateRegex.test(potentialDate) && hasDateSeparator) || monthsRegex.test(potentialDate)) {
                        this.advance(); // consume opening #
                        let dateValue = '';
                        while (this.peek() !== '#' && this.peek() !== '\n' && this.peek() !== '\0') {
                            dateValue += this.advance();
                        }
                        if (this.peek() === '#') {
                            this.advance(); // consume closing #
                            return { type: TokenType.Date, value: dateValue, line: startLine, column: startColumn };
                        }
                    }
                }
                // Fallback to OperatorHash
                this.advance();
                return { type: TokenType.OperatorHash, value: '#', line: startLine, column: startColumn };
            }

            // Line continuation: _ must be immediately followed by \n or \r\n (no trailing spaces)
            if (char === '_') {
                const next = this.pos + 1 < this.input.length ? this.input[this.pos + 1] : '\0';
                const afterCr = next === '\r' && this.pos + 2 < this.input.length ? this.input[this.pos + 2] : '\0';
                if (next === '\n' || (next === '\r' && afterCr === '\n')) {
                    this.advance(); // consume '_'
                    if (this.peek() === '\r') this.advance(); // consume \r
                    if (this.peek() === '\n') this.advance(); // consume \n (advance() updates this.line)
                    continue;
                }
                // _ followed by spaces (with optional comment) then newline: invalid — must be _ immediately before newline
                let lookahead = this.pos + 1;
                while (lookahead < this.input.length && this.input[lookahead] === ' ') lookahead++;
                const afterSpaces = lookahead < this.input.length ? this.input[lookahead] : '\0';
                if (afterSpaces === '\n' || afterSpaces === '\r' || afterSpaces === "'") {
                    const msg = afterSpaces === "'"
                        ? "行継続文字 '_' の後にコメントは記述できません"
                        : "行継続文字 '_' の直後に空白を置くことはできません";
                    this.advance(); // consume '_'
                    while (this.peek() !== '\n' && this.peek() !== '\0') this.advance(); // consume rest of line
                    if (this.peek() === '\n') this.advance();
                    throw new LexError(msg, startLine, startColumn);
                }
            }

            // Handle single quote comment
            if (char === "'") {
                // Collect comment text to check for trailing _ (which would be useless)
                let commentText = '';
                while (this.peek() !== '\n' && this.peek() !== '\0') {
                    commentText += this.advance();
                }
                if (/[ \t]*_[ \t]*$/.test(commentText)) {
                    this.diagnostics.push({
                        message: "コメント末尾の '_' は行継続として機能しません",
                        line: startLine,
                        column: startColumn + 1 + commentText.lastIndexOf('_'),
                    });
                }
                continue; // Skip the comment and loop again to get the next valid token
            }

            if (char === '\n') {
                this.advance();
                return { type: TokenType.Newline, value: '\n', line: startLine, column: startColumn };
            }

            if (char === '"') {
                this.advance(); // consume opening quote
                let strValue = '';
                while (this.peek() !== '\0') {
                    if (this.peek() === '"') {
                        this.advance(); // consume first quote
                        if (this.peek() === '"') {
                            strValue += '"'; // it's an escaped quote
                            this.advance(); // consume second quote
                        } else {
                            // end of string
                            return { type: TokenType.String, value: strValue, line: startLine, column: startColumn };
                        }
                    } else {
                        strValue += this.advance();
                    }
                }
                return { type: TokenType.String, value: strValue, line: startLine, column: startColumn };
            }

            if (char === '=') {
                this.advance();
                if (this.peek() === '<') {
                    this.advance();
                    return { type: TokenType.OperatorLessThanOrEqual, value: '<=', line: startLine, column: startColumn };
                }
                if (this.peek() === '>') {
                    this.advance();
                    return { type: TokenType.OperatorGreaterThanOrEqual, value: '>=', line: startLine, column: startColumn };
                }
                return { type: TokenType.OperatorEquals, value: '=', line: startLine, column: startColumn };
            }

            if (char === '<') {
                this.advance();
                if (this.peek() === '>') {
                    this.advance();
                    return { type: TokenType.OperatorNotEquals, value: '<>', line: startLine, column: startColumn };
                }
                if (this.peek() === '=') {
                    this.advance();
                    return { type: TokenType.OperatorLessThanOrEqual, value: '<=', line: startLine, column: startColumn };
                }
                return { type: TokenType.OperatorLessThan, value: '<', line: startLine, column: startColumn };
            }

            if (char === '>') {
                this.advance();
                if (this.peek() === '<') {
                    this.advance();
                    return { type: TokenType.OperatorNotEquals, value: '<>', line: startLine, column: startColumn };
                }
                if (this.peek() === '=') {
                    this.advance();
                    return { type: TokenType.OperatorGreaterThanOrEqual, value: '>=', line: startLine, column: startColumn };
                }
                return { type: TokenType.OperatorGreaterThan, value: '>', line: startLine, column: startColumn };
            }

            if (char === '+') {
                this.advance();
                return { type: TokenType.OperatorPlus, value: '+', line: startLine, column: startColumn };
            }

            if (char === '-') {
                this.advance();
                return { type: TokenType.OperatorMinus, value: '-', line: startLine, column: startColumn };
            }

            if (char === '!') {
                this.advance();
                return { type: TokenType.OperatorExclamation, value: '!', line: startLine, column: startColumn };
            }

            if (char === '&') {
                this.advance();
                const next = this.peek().toLowerCase();
                if (next === 'h') {
                    this.advance(); // consume 'h'
                    let hexStr = '';
                    while (/[0-9a-f]/i.test(this.peek())) {
                        hexStr += this.advance();
                    }
                    if (NUMERIC_TYPE_SUFFIXES.has(this.peek())) this.advance();
                    return { type: TokenType.Number, value: '0x' + hexStr, line: startLine, column: startColumn };
                } else if (next === 'o' || this.isDigit(next)) {
                    if (next === 'o') this.advance(); // consume 'o'
                    let octStr = '';
                    while (/[0-7]/.test(this.peek())) {
                        octStr += this.advance();
                    }
                    if (NUMERIC_TYPE_SUFFIXES.has(this.peek())) this.advance();
                    return { type: TokenType.Number, value: '0o' + octStr, line: startLine, column: startColumn };
                }
                return { type: TokenType.OperatorAmpersand, value: '&', line: startLine, column: startColumn };
            }

            if (char === ',') {
                this.advance();
                return { type: TokenType.OperatorComma, value: ',', line: startLine, column: startColumn };
            }

            if (char === '#') {
                this.advance();
                return { type: TokenType.OperatorHash, value: '#', line: startLine, column: startColumn };
            }

            if (char === '(') {
                this.advance();
                return { type: TokenType.OperatorLParen, value: '(', line: startLine, column: startColumn };
            }

            if (char === ')') {
                this.advance();
                return { type: TokenType.OperatorRParen, value: ')', line: startLine, column: startColumn };
            }

            // FOREIGN-NAME: "[" foreign-identifier "]"  (§3.3.5.2)
            if (char === '[') {
                this.advance(); // consume '['
                let foreignStr = '';
                while (this.peek() !== ']' && this.peek() !== '\n' && this.peek() !== '\r' && this.peek() !== '\0') {
                    foreignStr += this.advance();
                }
                if (this.peek() === ']') this.advance(); // consume ']'
                return { type: TokenType.ForeignName, value: foreignStr, line: startLine, column: startColumn };
            }

            if (char === '.') {
                this.advance();
                return { type: TokenType.OperatorDot, value: '.', line: startLine, column: startColumn };
            }

            if (char === ':') {
                this.advance();
                // Check for named argument syntax `:=`
                if (this.peek() === '=') {
                    this.advance();
                    return { type: TokenType.OperatorColonEquals, value: ':=', line: startLine, column: startColumn };
                }
                return { type: TokenType.OperatorColon, value: ':', line: startLine, column: startColumn };
            }

            if (char === ';') {
                this.advance();
                return { type: TokenType.OperatorSemicolon, value: ';', line: startLine, column: startColumn };
            }

            if (this.isDigit(char)) {
                let numStr = '';
                while (this.isDigit(this.peek())) {
                    numStr += this.advance();
                }
                if (this.peek() === '.') {
                    numStr += this.advance();
                    while (this.isDigit(this.peek())) {
                        numStr += this.advance();
                    }
                }
                // Handle scientific notation: [eEdD][+-]?digits.
                // VBA accepts D/d as an exponent marker for Double literals.
                const nextChar = this.peek().toLowerCase();
                if (nextChar === 'e' || nextChar === 'd') {
                    numStr += this.advance(); // consume exponent marker
                    if (this.peek() === '+' || this.peek() === '-') {
                        numStr += this.advance();
                    }
                    while (this.isDigit(this.peek())) {
                        numStr += this.advance();
                    }
                }
                // Check for VBA Type Declaration Suffixes for numbers (MS-VBAL §3.3.5.2)
                if (NUMERIC_TYPE_SUFFIXES.has(this.peek())) {
                    numStr += this.advance();
                }
                return { type: TokenType.Number, value: numStr, line: startLine, column: startColumn };
            }

            if (this.isAlpha(char)) {
                let idStr = '';
                while (this.isAlphaNumeric(this.peek())) {
                    idStr += this.advance();
                }
                // Handle type hint characters at the end of an identifier (MS-VBAL §3.3.5.2)
                // '!' and '^' are also operators (bang member-access / exponentiation).
                // Only consume as type suffix when the next character is NOT alphanumeric or '_',
                // so 'dict!Key' and 'x^2' keep them as operators.
                {
                    const nextCh = this.peek();
                    const charAfterNext = this.pos + 1 < this.input.length ? this.input[this.pos + 1] : '\0';
                    if (IDENTIFIER_TYPE_SUFFIXES.has(nextCh) &&
                            !((nextCh === '!' || nextCh === '^') && (this.isAlphaNumeric(charAfterNext) || charAfterNext === '_'))) {
                        idStr += this.advance();
                    }
                }

                const lowerId = idStr.toLowerCase();
                // Strip type suffix for keyword matching: 'dim$' → 'dim', 'for%' → 'for'
                const lowerBase = lowerId.replace(/[$%&#@]$/, '');
                if (lowerBase === 'rem') {
                    while (this.peek() !== '\n' && this.peek() !== '\0') {
                        this.advance();
                    }
                    // We've consumed the comment, now we need to get the actual next token.
                    // Since we are inside the 'while(true)' loop, we can just continue.
                    continue;
                }

                if (lowerBase === 'for') return { type: TokenType.KeywordFor, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'to') return { type: TokenType.KeywordTo, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'next') return { type: TokenType.KeywordNext, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'if') return { type: TokenType.KeywordIf, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'then') return { type: TokenType.KeywordThen, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'elseif') return { type: TokenType.KeywordElseIf, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'else') return { type: TokenType.KeywordElse, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'end') return { type: TokenType.KeywordEnd, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'do') return { type: TokenType.KeywordDo, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'while') return { type: TokenType.KeywordWhile, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'wend') return { type: TokenType.KeywordWend, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'loop') return { type: TokenType.KeywordLoop, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'until') return { type: TokenType.KeywordUntil, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'gosub') return { type: TokenType.KeywordGoSub, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'return') return { type: TokenType.KeywordReturn, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'lset') return { type: TokenType.KeywordLSet, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'rset') return { type: TokenType.KeywordRSet, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'stop') return { type: TokenType.KeywordStop, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'sub') return { type: TokenType.KeywordSub, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'function') return { type: TokenType.KeywordFunction, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'property') return { type: TokenType.KeywordProperty, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'get') return { type: TokenType.KeywordGet, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'let') return { type: TokenType.KeywordLet, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'dim') return { type: TokenType.KeywordDim, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'as') return { type: TokenType.KeywordAs, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'new') return { type: TokenType.KeywordNew, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'collection') return { type: TokenType.KeywordCollection, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'like') return { type: TokenType.KeywordLike, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'and') return { type: TokenType.KeywordAnd, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'or') return { type: TokenType.KeywordOr, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'xor') return { type: TokenType.KeywordXor, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'eqv') return { type: TokenType.KeywordEqv, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'imp') return { type: TokenType.KeywordImp, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'not') return { type: TokenType.KeywordNot, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'option') return { type: TokenType.KeywordOption, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'explicit') return { type: TokenType.KeywordExplicit, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'const') return { type: TokenType.KeywordConst, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'set') return { type: TokenType.KeywordSet, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'call') return { type: TokenType.KeywordCall, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'on') return { type: TokenType.KeywordOn, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'error') return { type: TokenType.KeywordError, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'goto') return { type: TokenType.KeywordGoTo, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'erase') return { type: TokenType.KeywordErase, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'redim') return { type: TokenType.KeywordReDim, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'step') return { type: TokenType.KeywordStep, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'empty') return { type: TokenType.KeywordEmpty, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'exit') return { type: TokenType.KeywordExit, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'byref') return { type: TokenType.KeywordByRef, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'byval') return { type: TokenType.KeywordByVal, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'mod') return { type: TokenType.KeywordMod, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'type') return { type: TokenType.KeywordType, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'nothing') return { type: TokenType.KeywordNothing, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'optional') return { type: TokenType.KeywordOptional, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'is') return { type: TokenType.KeywordIs, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'resume') return { type: TokenType.KeywordResume, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'select') return { type: TokenType.KeywordSelect, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'case') return { type: TokenType.KeywordCase, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'each') return { type: TokenType.KeywordEach, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'in') return { type: TokenType.KeywordIn, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'public') return { type: TokenType.KeywordPublic, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'private') return { type: TokenType.KeywordPrivate, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'global') return { type: TokenType.KeywordGlobal, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'enum') return { type: TokenType.KeywordEnum, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'typeof') return { type: TokenType.KeywordTypeOf, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'friend') return { type: TokenType.KeywordFriend, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'with') return { type: TokenType.KeywordWith, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'null') return { type: TokenType.KeywordNull, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'static') return { type: TokenType.KeywordStatic, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'class') return { type: TokenType.KeywordClass, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'me') return { type: TokenType.KeywordMe, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'compare') return { type: TokenType.KeywordCompare, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'binary') return { type: TokenType.KeywordBinary, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'text') return { type: TokenType.KeywordText, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'attribute') return { type: TokenType.KeywordAttribute, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'declare') return { type: TokenType.KeywordDeclare, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'lib') return { type: TokenType.KeywordLib, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'alias') return { type: TokenType.KeywordAlias, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'ptrsafe') return { type: TokenType.KeywordPtrSafe, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'open') return { type: TokenType.KeywordOpen, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'close') return { type: TokenType.KeywordClose, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'line') return { type: TokenType.KeywordLine, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'input') return { type: TokenType.KeywordInput, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'print') return { type: TokenType.KeywordPrint, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'put') return { type: TokenType.KeywordPut, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'output') return { type: TokenType.KeywordOutput, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'append') return { type: TokenType.KeywordAppend, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'random') return { type: TokenType.KeywordRandom, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'access') return { type: TokenType.KeywordAccess, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'read') return { type: TokenType.KeywordRead, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'write') return { type: TokenType.KeywordWrite, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'lock') return { type: TokenType.KeywordLock, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'shared') return { type: TokenType.KeywordShared, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'spc') return { type: TokenType.KeywordSpc, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'tab') return { type: TokenType.KeywordTab, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'kill') return { type: TokenType.KeywordKill, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'get') return { type: TokenType.KeywordGet, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'seek') return { type: TokenType.KeywordSeek, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'reset') return { type: TokenType.KeywordReset, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'unlock') return { type: TokenType.KeywordUnlock, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'paramarray') return { type: TokenType.KeywordParamArray, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'module') return { type: TokenType.KeywordModule, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'event') return { type: TokenType.KeywordEvent, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'raiseevent') return { type: TokenType.KeywordRaiseEvent, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'withevents') return { type: TokenType.KeywordWithEvents, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'implements') return { type: TokenType.KeywordImplements, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'appactivate') return { type: TokenType.KeywordAppActivate, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'sendkeys') return { type: TokenType.KeywordSendKeys, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'mid') return { type: TokenType.KeywordMid, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'midb') return { type: TokenType.KeywordMid, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'width') return { type: TokenType.KeywordWidth, value: idStr, line: startLine, column: startColumn };
                if (lowerBase === 'addressof') return { type: TokenType.KeywordAddressOf, value: idStr, line: startLine, column: startColumn };
                if (['defbool','defbyte','defint','deflng','deflnglng','deflngptr',
                     'defsng','defdbl','defcur','defdate','defstr','defobj','defvar'].includes(lowerBase))
                    return { type: TokenType.KeywordDef, value: idStr, line: startLine, column: startColumn };

                return { type: TokenType.Identifier, value: idStr, line: startLine, column: startColumn };
            }

            if (char === '*') {
                this.advance();
                return { type: TokenType.OperatorMultiply, value: '*', line: startLine, column: startColumn };
            }
            if (char === '/') {
                this.advance();
                return { type: TokenType.OperatorDivide, value: '/', line: startLine, column: startColumn };
            }
            if (char === '\\') {
                this.advance();
                return { type: TokenType.OperatorIntDivide, value: '\\', line: startLine, column: startColumn };
            }
            if (char === '^') {
                this.advance();
                return { type: TokenType.OperatorPower, value: '^', line: startLine, column: startColumn };
            }

            // Unknown character
            const unknownChar = this.advance();
            return { type: TokenType.Unknown, value: unknownChar, line: startLine, column: startColumn };
        } // End of while(true) loop
    }

    public tokenize(): Token[] {
        const tokens: Token[] = [];
        let token: Token;
        do {
            token = this.getNextToken();
            tokens.push(token);
        } while (token.type !== TokenType.EOF);
        return tokens;
    }
}
