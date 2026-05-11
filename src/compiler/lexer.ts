export enum TokenType {
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
    Date,
    Newline,
    EOF,
    Unknown
}

export interface Token {
    type: TokenType;
    value: string;
    line: number;
}

export class Lexer {
    private input: string = '';
    private pos: number = 0;
    private line: number = 1;

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
        }
        return char;
    }

    private isWhitespace(char: string): boolean {
        return char === ' ' || char === '\t' || char === '\r';
    }

    private isAlpha(char: string): boolean {
        return (char >= 'a' && char <= 'z') || (char >= 'A' && char <= 'Z') || char === '_';
    }

    private isAlphaNumeric(char: string): boolean {
        return this.isAlpha(char) || (char >= '0' && char <= '9');
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

            if (this.pos >= this.input.length) {
                return { type: TokenType.EOF, value: '', line: this.line };
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
                    this.advance(); // consume opening #
                    let dateValue = '';
                    while (this.peek() !== '#' && this.peek() !== '\n' && this.peek() !== '\0') {
                        dateValue += this.advance();
                    }
                    if (this.peek() === '#') {
                        this.advance(); // consume closing #
                        return { type: TokenType.Date, value: dateValue, line: this.line };
                    }
                }
                // Fallback to OperatorHash
                this.advance();
                return { type: TokenType.OperatorHash, value: '#', line: this.line };
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
            }

            // Handle single quote comment
            if (char === "'") {
                while (this.peek() !== '\n' && this.peek() !== '\0') {
                    this.advance();
                }
                continue; // Skip the comment and loop again to get the next valid token
            }

            if (char === '\n') {
                this.advance();
                return { type: TokenType.Newline, value: '\n', line: this.line - 1 };
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
                            return { type: TokenType.String, value: strValue, line: this.line };
                        }
                    } else {
                        strValue += this.advance();
                    }
                }
                return { type: TokenType.String, value: strValue, line: this.line };
            }

            if (char === '=') {
                this.advance();
                return { type: TokenType.OperatorEquals, value: '=', line: this.line };
            }

            if (char === '<') {
                this.advance();
                if (this.peek() === '>') {
                    this.advance();
                    return { type: TokenType.OperatorNotEquals, value: '<>', line: this.line };
                }
                if (this.peek() === '=') {
                    this.advance();
                    return { type: TokenType.OperatorLessThanOrEqual, value: '<=', line: this.line };
                }
                return { type: TokenType.OperatorLessThan, value: '<', line: this.line };
            }

            if (char === '>') {
                this.advance();
                if (this.peek() === '=') {
                    this.advance();
                    return { type: TokenType.OperatorGreaterThanOrEqual, value: '>=', line: this.line };
                }
                return { type: TokenType.OperatorGreaterThan, value: '>', line: this.line };
            }

            if (char === '+') {
                this.advance();
                return { type: TokenType.OperatorPlus, value: '+', line: this.line };
            }

            if (char === '-') {
                this.advance();
                return { type: TokenType.OperatorMinus, value: '-', line: this.line };
            }

            if (char === '!') {
                this.advance();
                return { type: TokenType.OperatorExclamation, value: '!', line: this.line };
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
                    return { type: TokenType.Number, value: '0x' + hexStr, line: this.line };
                } else if (next === 'o' || this.isDigit(next)) {
                    if (next === 'o') this.advance(); // consume 'o'
                    let octStr = '';
                    while (/[0-7]/.test(this.peek())) {
                        octStr += this.advance();
                    }
                    return { type: TokenType.Number, value: '0o' + octStr, line: this.line };
                }
                return { type: TokenType.OperatorAmpersand, value: '&', line: this.line };
            }

            if (char === ',') {
                this.advance();
                return { type: TokenType.OperatorComma, value: ',', line: this.line };
            }

            if (char === '#') {
                this.advance();
                return { type: TokenType.OperatorHash, value: '#', line: this.line };
            }

            if (char === '(') {
                this.advance();
                return { type: TokenType.OperatorLParen, value: '(', line: this.line };
            }

            if (char === ')') {
                this.advance();
                return { type: TokenType.OperatorRParen, value: ')', line: this.line };
            }

            if (char === '.') {
                this.advance();
                return { type: TokenType.OperatorDot, value: '.', line: this.line };
            }

            if (char === ':') {
                this.advance();
                // Check for named argument syntax `:=`
                if (this.peek() === '=') {
                    this.advance();
                    return { type: TokenType.OperatorColonEquals, value: ':=', line: this.line };
                }
                return { type: TokenType.OperatorColon, value: ':', line: this.line };
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
                // Handle scientific notation: [eE][+-]?digits
                const nextChar = this.peek().toLowerCase();
                if (nextChar === 'e') {
                    numStr += this.advance(); // consume 'e'
                    if (this.peek() === '+' || this.peek() === '-') {
                        numStr += this.advance();
                    }
                    while (this.isDigit(this.peek())) {
                        numStr += this.advance();
                    }
                }
                // Check for VBA Type Declaration Suffixes for numbers (%, &, @, !, #, ^)
                const peekChar = this.peek();
                if (['%', '&', '@', '!', '#', '^'].indexOf(peekChar) !== -1) {
                    numStr += this.advance();
                }
                return { type: TokenType.Number, value: numStr, line: this.line };
            }

            if (this.isAlpha(char)) {
                let idStr = '';
                while (this.isAlphaNumeric(this.peek())) {
                    idStr += this.advance();
                }
                // Handle type hint characters at the end of an identifier
                const typeHints = ['$', '%', '&', '!', '#', '@'];
                if (typeHints.includes(this.peek())) {
                    idStr += this.advance();
                }

                const lowerId = idStr.toLowerCase();
                if (lowerId === 'rem') {
                    while (this.peek() !== '\n' && this.peek() !== '\0') {
                        this.advance();
                    }
                    // We've consumed the comment, now we need to get the actual next token.
                    // Since we are inside the 'while(true)' loop, we can just continue.
                    continue;
                }

                if (lowerId === 'for') return { type: TokenType.KeywordFor, value: idStr, line: this.line };
                if (lowerId === 'to') return { type: TokenType.KeywordTo, value: idStr, line: this.line };
                if (lowerId === 'next') return { type: TokenType.KeywordNext, value: idStr, line: this.line };
                if (lowerId === 'if') return { type: TokenType.KeywordIf, value: idStr, line: this.line };
                if (lowerId === 'then') return { type: TokenType.KeywordThen, value: idStr, line: this.line };
                if (lowerId === 'elseif') return { type: TokenType.KeywordElseIf, value: idStr, line: this.line };
                if (lowerId === 'else') return { type: TokenType.KeywordElse, value: idStr, line: this.line };
                if (lowerId === 'end') return { type: TokenType.KeywordEnd, value: idStr, line: this.line };
                if (lowerId === 'do') return { type: TokenType.KeywordDo, value: idStr, line: this.line };
                if (lowerId === 'while') return { type: TokenType.KeywordWhile, value: idStr, line: this.line };
                if (lowerId === 'wend') return { type: TokenType.KeywordWend, value: idStr, line: this.line };
                if (lowerId === 'loop') return { type: TokenType.KeywordLoop, value: idStr, line: this.line };
                if (lowerId === 'until') return { type: TokenType.KeywordUntil, value: idStr, line: this.line };
                if (lowerId === 'gosub') return { type: TokenType.KeywordGoSub, value: idStr, line: this.line };
                if (lowerId === 'return') return { type: TokenType.KeywordReturn, value: idStr, line: this.line };
                if (lowerId === 'lset') return { type: TokenType.KeywordLSet, value: idStr, line: this.line };
                if (lowerId === 'rset') return { type: TokenType.KeywordRSet, value: idStr, line: this.line };
                if (lowerId === 'stop') return { type: TokenType.KeywordStop, value: idStr, line: this.line };
                if (lowerId === 'sub') return { type: TokenType.KeywordSub, value: idStr, line: this.line };
                if (lowerId === 'function') return { type: TokenType.KeywordFunction, value: idStr, line: this.line };
                if (lowerId === 'property') return { type: TokenType.KeywordProperty, value: idStr, line: this.line };
                if (lowerId === 'get') return { type: TokenType.KeywordGet, value: idStr, line: this.line };
                if (lowerId === 'let') return { type: TokenType.KeywordLet, value: idStr, line: this.line };
                if (lowerId === 'dim') return { type: TokenType.KeywordDim, value: idStr, line: this.line };
                if (lowerId === 'as') return { type: TokenType.KeywordAs, value: idStr, line: this.line };
                if (lowerId === 'new') return { type: TokenType.KeywordNew, value: idStr, line: this.line };
                if (lowerId === 'collection') return { type: TokenType.KeywordCollection, value: idStr, line: this.line };
                if (lowerId === 'like') return { type: TokenType.KeywordLike, value: idStr, line: this.line };
                if (lowerId === 'and') return { type: TokenType.KeywordAnd, value: idStr, line: this.line };
                if (lowerId === 'or') return { type: TokenType.KeywordOr, value: idStr, line: this.line };
                if (lowerId === 'xor') return { type: TokenType.KeywordXor, value: idStr, line: this.line };
                if (lowerId === 'eqv') return { type: TokenType.KeywordEqv, value: idStr, line: this.line };
                if (lowerId === 'imp') return { type: TokenType.KeywordImp, value: idStr, line: this.line };
                if (lowerId === 'not') return { type: TokenType.KeywordNot, value: idStr, line: this.line };
                if (lowerId === 'option') return { type: TokenType.KeywordOption, value: idStr, line: this.line };
                if (lowerId === 'explicit') return { type: TokenType.KeywordExplicit, value: idStr, line: this.line };
                if (lowerId === 'const') return { type: TokenType.KeywordConst, value: idStr, line: this.line };
                if (lowerId === 'set') return { type: TokenType.KeywordSet, value: idStr, line: this.line };
                if (lowerId === 'call') return { type: TokenType.KeywordCall, value: idStr, line: this.line };
                if (lowerId === 'on') return { type: TokenType.KeywordOn, value: idStr, line: this.line };
                if (lowerId === 'error') return { type: TokenType.KeywordError, value: idStr, line: this.line };
                if (lowerId === 'goto') return { type: TokenType.KeywordGoTo, value: idStr, line: this.line };
                if (lowerId === 'erase') return { type: TokenType.KeywordErase, value: idStr, line: this.line };
                if (lowerId === 'redim') return { type: TokenType.KeywordReDim, value: idStr, line: this.line };
                if (lowerId === 'step') return { type: TokenType.KeywordStep, value: idStr, line: this.line };
                if (lowerId === 'empty') return { type: TokenType.KeywordEmpty, value: idStr, line: this.line };
                if (lowerId === 'exit') return { type: TokenType.KeywordExit, value: idStr, line: this.line };
                if (lowerId === 'byref') return { type: TokenType.KeywordByRef, value: idStr, line: this.line };
                if (lowerId === 'byval') return { type: TokenType.KeywordByVal, value: idStr, line: this.line };
                if (lowerId === 'mod') return { type: TokenType.KeywordMod, value: idStr, line: this.line };
                if (lowerId === 'type') return { type: TokenType.KeywordType, value: idStr, line: this.line };
                if (lowerId === 'nothing') return { type: TokenType.KeywordNothing, value: idStr, line: this.line };
                if (lowerId === 'optional') return { type: TokenType.KeywordOptional, value: idStr, line: this.line };
                if (lowerId === 'is') return { type: TokenType.KeywordIs, value: idStr, line: this.line };
                if (lowerId === 'resume') return { type: TokenType.KeywordResume, value: idStr, line: this.line };
                if (lowerId === 'select') return { type: TokenType.KeywordSelect, value: idStr, line: this.line };
                if (lowerId === 'case') return { type: TokenType.KeywordCase, value: idStr, line: this.line };
                if (lowerId === 'each') return { type: TokenType.KeywordEach, value: idStr, line: this.line };
                if (lowerId === 'in') return { type: TokenType.KeywordIn, value: idStr, line: this.line };
                if (lowerId === 'public') return { type: TokenType.KeywordPublic, value: idStr, line: this.line };
                if (lowerId === 'private') return { type: TokenType.KeywordPrivate, value: idStr, line: this.line };
                if (lowerId === 'enum') return { type: TokenType.KeywordEnum, value: idStr, line: this.line };
                if (lowerId === 'typeof') return { type: TokenType.KeywordTypeOf, value: idStr, line: this.line };
                if (lowerId === 'friend') return { type: TokenType.KeywordFriend, value: idStr, line: this.line };
                if (lowerId === 'with') return { type: TokenType.KeywordWith, value: idStr, line: this.line };
                if (lowerId === 'null') return { type: TokenType.KeywordNull, value: idStr, line: this.line };
                if (lowerId === 'static') return { type: TokenType.KeywordStatic, value: idStr, line: this.line };
                if (lowerId === 'class') return { type: TokenType.KeywordClass, value: idStr, line: this.line };
                if (lowerId === 'me') return { type: TokenType.KeywordMe, value: idStr, line: this.line };
                if (lowerId === 'compare') return { type: TokenType.KeywordCompare, value: idStr, line: this.line };
                if (lowerId === 'binary') return { type: TokenType.KeywordBinary, value: idStr, line: this.line };
                if (lowerId === 'text') return { type: TokenType.KeywordText, value: idStr, line: this.line };
                if (lowerId === 'attribute') return { type: TokenType.KeywordAttribute, value: idStr, line: this.line };
                if (lowerId === 'declare') return { type: TokenType.KeywordDeclare, value: idStr, line: this.line };
                if (lowerId === 'lib') return { type: TokenType.KeywordLib, value: idStr, line: this.line };
                if (lowerId === 'alias') return { type: TokenType.KeywordAlias, value: idStr, line: this.line };
                if (lowerId === 'ptrsafe') return { type: TokenType.KeywordPtrSafe, value: idStr, line: this.line };
                if (lowerId === 'open') return { type: TokenType.KeywordOpen, value: idStr, line: this.line };
                if (lowerId === 'close') return { type: TokenType.KeywordClose, value: idStr, line: this.line };
                if (lowerId === 'line') return { type: TokenType.KeywordLine, value: idStr, line: this.line };
                if (lowerId === 'input') return { type: TokenType.KeywordInput, value: idStr, line: this.line };
                if (lowerId === 'print') return { type: TokenType.KeywordPrint, value: idStr, line: this.line };
                if (lowerId === 'put') return { type: TokenType.KeywordPut, value: idStr, line: this.line };
                if (lowerId === 'output') return { type: TokenType.KeywordOutput, value: idStr, line: this.line };
                if (lowerId === 'append') return { type: TokenType.KeywordAppend, value: idStr, line: this.line };
                if (lowerId === 'random') return { type: TokenType.KeywordRandom, value: idStr, line: this.line };
                if (lowerId === 'access') return { type: TokenType.KeywordAccess, value: idStr, line: this.line };
                if (lowerId === 'read') return { type: TokenType.KeywordRead, value: idStr, line: this.line };
                if (lowerId === 'write') return { type: TokenType.KeywordWrite, value: idStr, line: this.line };
                if (lowerId === 'lock') return { type: TokenType.KeywordLock, value: idStr, line: this.line };
                if (lowerId === 'shared') return { type: TokenType.KeywordShared, value: idStr, line: this.line };
                if (lowerId === 'spc') return { type: TokenType.KeywordSpc, value: idStr, line: this.line };
                if (lowerId === 'tab') return { type: TokenType.KeywordTab, value: idStr, line: this.line };
                if (lowerId === 'kill') return { type: TokenType.KeywordKill, value: idStr, line: this.line };
                if (lowerId === 'base') return { type: TokenType.KeywordBase, value: idStr, line: this.line };
                if (lowerId === 'module') return { type: TokenType.KeywordModule, value: idStr, line: this.line };

                return { type: TokenType.Identifier, value: idStr, line: this.line };
            }

            if (char === '*') {
                this.advance();
                return { type: TokenType.OperatorMultiply, value: '*', line: this.line };
            }
            if (char === '/') {
                this.advance();
                return { type: TokenType.OperatorDivide, value: '/', line: this.line };
            }
            if (char === '\\') {
                this.advance();
                return { type: TokenType.OperatorIntDivide, value: '\\', line: this.line };
            }
            if (char === '^') {
                this.advance();
                return { type: TokenType.OperatorPower, value: '^', line: this.line };
            }

            // Unknown character
            const unknownChar = this.advance();
            return { type: TokenType.Unknown, value: unknownChar, line: this.line };
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
