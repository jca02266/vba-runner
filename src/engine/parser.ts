import { Token, TokenType } from './lexer';

export interface ResumeStatement extends Statement {
    type: 'ResumeStatement';
    target: string; // 'Next' or label
}

export interface GoToStatement extends Statement {
    type: 'GoToStatement';
    label: string;
}

export interface StopStatement extends Statement {
    type: 'StopStatement';
}

export interface EndStatement extends Statement {
    type: 'EndStatement';
}

export interface GoSubStatement extends Statement {
    type: 'GoSubStatement';
    label: string;
}

export interface ReturnStatement extends Statement {
    type: 'ReturnStatement';
}

export interface OnGoToSubStatement extends Statement {
    type: 'OnGoToSubStatement';
    expression: Expression;
    isGoSub: boolean;
    labels: string[];
}

export interface Position {
    line: number;   // 1-based
    column: number; // 1-based
}

export interface SourceLocation {
    start: Position;
    end: Position;
}

export interface ASTNode {
    type: string;
    line?: number;        // kept for backward compat (= loc.start.line when loc is set)
    loc?: SourceLocation; // source position (ESTree convention)
}

export interface ParseDiagnostic {
    message: string;
    loc: SourceLocation;
    severity: 'error' | 'warning';
}

export interface Program extends ASTNode {
    type: 'Program';
    body: Statement[];
    diagnostics: ParseDiagnostic[];
}

// eslint-disable-next-line @typescript-eslint/no-empty-object-type
export interface Statement extends ASTNode { }

export interface ForStatement extends Statement {
    type: 'ForStatement';
    identifier: Identifier;
    start: Expression;
    end: Expression;
    step?: Expression;
    body: Statement[];
    nextIdentifier?: Identifier;
}

export interface ForEachStatement extends Statement {
    type: 'ForEachStatement';
    variable: Identifier;
    collection: Expression;
    body: Statement[];
    nextIdentifier?: Identifier;
}

export interface IfStatement extends Statement {
    type: 'IfStatement';
    condition: Expression;
    consequent: Statement[];
    alternate: Statement[] | IfStatement | null; // For Else and ElseIf
}

export interface DoWhileStatement extends Statement {
    type: 'DoWhileStatement';
    conditionType?: 'while' | 'until'; // undefined = infinite loop
    conditionPosition?: 'pre' | 'post'; // undefined = infinite loop
    condition?: Expression;
    body: Statement[];
}

export interface WithStatement extends Statement {
    type: 'WithStatement';
    expression: Expression;
    body: Statement[];
}

export interface WhileStatement extends Statement {
    type: 'WhileStatement';
    condition: Expression;
    body: Statement[];
}

export interface Parameter extends ASTNode {
    type: 'Parameter';
    name: string;
    isByVal: boolean;
    /** ByVal または ByRef が明示されている場合 true */
    hasPassingModifier: boolean;
    isOptional?: boolean;
    isParamArray?: boolean;
    isArray?: boolean;
    paramType?: string;
    defaultValue?: Expression;
}

export interface ProcedureDeclaration extends Statement {
    type: 'ProcedureDeclaration';
    isFunction: boolean;
    isProperty: boolean;
    propertyType?: 'get' | 'let' | 'set';
    name: Identifier;
    parameters: Parameter[];
    returnType?: string;
    /** true when the return type is written as an array (e.g. `As String()`) */
    returnsArray?: boolean;
    body: Statement[];
    scope?: 'public' | 'private' | 'friend';
    isStatic?: boolean;
    moduleName?: string;
    /** 1-based column immediately after the closing ')' of the parameter list */
    paramsEndColumn?: number;
}

export interface ArrayBound {
    lower?: Expression;
    upper: Expression;
}

export interface VariableDeclarator {
    name: Identifier;
    isArray: boolean;
    arrayBounds?: ArrayBound[];
    isNew: boolean;
    isWithEvents: boolean;
    objectType?: string;
    objectTypeLoc?: SourceLocation;
    /** 1-based column immediately after the closing ')' of the array bounds, when isArray is true */
    arrayEndColumn?: number;
}

export interface VariableDeclaration extends Statement {
    type: 'VariableDeclaration';
    declarations: VariableDeclarator[];
    isStatic?: boolean;
    scope?: 'public' | 'private' | 'friend';
}

export interface ConstDeclaration extends Statement {
    type: 'ConstDeclaration';
    name: Identifier;
    value: Expression;
}

export interface SetStatement extends Statement {
    type: 'SetStatement';
    left: Expression;
    right: Expression;
}

export interface OnErrorStatement extends Statement {
    type: 'OnErrorStatement';
    label: string; // "Cleanup", "Resume Next", "0"
}

export interface OptionExplicitStatement extends Statement {
    type: 'OptionExplicitStatement';
}

export interface OptionBaseStatement extends Statement {
    type: 'OptionBaseStatement';
    base: 0 | 1;
}

export interface OptionPrivateModuleStatement extends Statement {
    type: 'OptionPrivateModuleStatement';
}

export interface OptionCompareStatement extends Statement {
    type: 'OptionCompareStatement';
    mode: 'Binary' | 'Text';
}

export interface AttributeStatement extends Statement {
    type: 'AttributeStatement';
    name: string;
    value: Expression;
}

export interface DefDirective extends Statement {
    type: 'DefDirective';
    vbaType: string;         // e.g. "Integer", "String"
    ranges: { from: string; to: string }[];  // letter ranges (lowercased)
}

export interface OpenStatement extends Statement {
    type: 'OpenStatement';
    path: Expression;
    mode: 'Input' | 'Output' | 'Append' | 'Random' | 'Binary';
    access?: 'Read' | 'Write' | 'Read Write';
    lock?: 'Shared' | 'Lock Read' | 'Lock Write' | 'Lock Read Write';
    fileNumber: Expression;
}

export interface LockStatement extends Statement {
    type: 'LockStatement';
    fileNumber: Expression;
    recordRange?: { start: Expression, end?: Expression };
}

export interface UnlockStatement extends Statement {
    type: 'UnlockStatement';
    fileNumber: Expression;
    recordRange?: { start: Expression, end?: Expression };
}

export interface WidthStatement extends Statement {
    type: 'WidthStatement';
    fileNumber: Expression;
    width: Expression;
}

export interface CloseStatement extends Statement {
    type: 'CloseStatement';
    fileNumbers: Expression[]; // Empty means close all
}

export interface PrintStatement extends Statement {
    type: 'PrintStatement';
    fileNumber: Expression;
    expressions: (Expression | 'Spc' | 'Tab' | 'Comma' | 'Semicolon')[];
}

export interface LineInputStatement extends Statement {
    type: 'LineInputStatement';
    fileNumber: Expression;
    variable: Identifier;
}

export interface PutStatement extends Statement {
    type: 'PutStatement';
    fileNumber: Expression;
    recordNumber?: Expression;
    data: Expression;
}

export interface KillStatement extends Statement {
    type: 'KillStatement';
    path: Expression;
}

export interface WriteStatement extends Statement {
    type: 'WriteStatement';
    fileNumber: Expression;
    items: Expression[];
}

export interface InputStatement extends Statement {
    type: 'InputStatement';
    fileNumber: Expression;
    variables: Expression[];
}

export interface GetStatement extends Statement {
    type: 'GetStatement';
    fileNumber: Expression;
    recordNumber?: Expression;
    variable: Expression;
}

export interface SeekStatement extends Statement {
    type: 'SeekStatement';
    fileNumber: Expression;
    position: Expression;
}

export interface ResetStatement extends Statement {
    type: 'ResetStatement';
}

export interface LSetStatement extends Statement {
    type: 'LSetStatement';
    left: Expression;
    right: Expression;
}

export interface RSetStatement extends Statement {
    type: 'RSetStatement';
    left: Expression;
    right: Expression;
}

export interface MidStatement extends Statement {
    type: 'MidStatement';
    target: Expression;
    start: Expression;
    length: Expression | null;
    value: Expression;
    isByte: boolean;
}

export interface EraseStatement extends Statement {
    type: 'EraseStatement';
    name: Identifier;
}

export interface ImplementsDirective extends Statement {
    type: 'ImplementsDirective';
    interfaceName: string;
}

export interface AppActivateStatement extends Statement {
    type: 'AppActivateStatement';
    title: Expression;
    wait?: Expression;
}

export interface SendKeysStatement extends Statement {
    type: 'SendKeysStatement';
    keys: Expression;
    wait?: Expression;
}

export interface EventDeclaration extends Statement {
    type: 'EventDeclaration';
    name: Identifier;
    parameters: Parameter[];
    scope?: 'public' | 'private' | 'friend';
}

export interface RaiseEventStatement extends Statement {
    type: 'RaiseEventStatement';
    eventName: Identifier;
    args: Expression[];
}

export interface DeclareStatement extends Statement {
    type: 'DeclareStatement';
    isPtrSafe: boolean;
    isSub: boolean;
    name: string;
    libName: string;
    aliasName?: string;
    parameters: Parameter[];
    returnType?: string;
}

export interface ReDimStatement extends Statement {
    type: 'ReDimStatement';
    name: Identifier;
    bounds: ArrayBound[]; // Multi-dimensional bounds (e.g. 1 To numDays)
    isPreserve: boolean;
    objectType?: string;
}

export interface AddressOfExpression extends Expression {
    type: 'AddressOfExpression';
    procedureName: Identifier;
    moduleName?: string;  // set when AddressOf Module.Proc form (§5.6.16.8)
}

export interface ErrorStatement extends Statement {
    type: 'ErrorStatement';
    errorNumber: Expression;
}

export interface ExitStatement extends Statement {
    type: 'ExitStatement';
    exitType: 'For' | 'Do' | 'Sub' | 'Function' | 'Property';
}

export interface LabelStatement extends Statement {
    type: 'LabelStatement';
    label: string;
}

export interface EnumMember {
    name: Identifier;
    value?: Expression;
}

export interface EnumDeclaration extends Statement {
    type: 'EnumDeclaration';
    name: Identifier;
    members: EnumMember[];
    scope?: 'public' | 'private';
}

export interface TypeMember {
    name: string;
    memberType: string;
}

export interface TypeDeclaration extends Statement {
    type: 'TypeDeclaration';
    name: string;
    members: TypeMember[];
}

export interface ClassDeclaration extends Statement {
    type: 'ClassDeclaration';
    name: string;
    fields: VariableDeclaration[];
    procedures: ProcedureDeclaration[];
    body: Statement[];
}

export type RangeClause =
    | { kind: 'expression'; value: Expression }
    | { kind: 'to'; start: Expression; end: Expression }
    | { kind: 'comparison'; operator: string; value: Expression };

export interface CaseClause {
    ranges: RangeClause[];
    body: Statement[];
}

export interface SelectCaseStatement extends Statement {
    type: 'SelectCaseStatement';
    expression: Expression;
    cases: CaseClause[];
    elseBody: Statement[] | null;
}

export interface AssignmentStatement extends Statement {
    type: 'AssignmentStatement';
    left: Expression; // Identifier, CallExpression (for arrays), MemberExpression
    right: Expression;
}

export interface CallStatement extends Statement {
    type: 'CallStatement';
    expression: CallExpression;
}

// eslint-disable-next-line @typescript-eslint/no-empty-object-type
export interface Expression extends ASTNode { }

export interface ParenthesizedExpression extends Expression {
    type: 'ParenthesizedExpression';
    expression: Expression;
}

export interface CallExpression extends Expression {
    type: 'CallExpression';
    callee: Expression;
    args: Expression[];
}

export interface NamedArgument extends Expression {
    type: 'NamedArgument';
    name: string;
    value: Expression;
}

export interface MemberExpression extends Expression {
    type: 'MemberExpression';
    object: Expression;
    property: Identifier;
}

export interface DictionaryAccessExpression extends Expression {
    type: 'DictionaryAccessExpression';
    object: Expression;
    property: Identifier;
}

export interface TypeOfIsExpression extends Expression {
    type: 'TypeOfIsExpression';
    expression: Expression;
    typeName: string;
}

export interface NewExpression extends Expression {
    type: 'NewExpression';
    className: string;
}

export interface Identifier extends Expression {
    type: 'Identifier';
    name: string;
    /** true when parsed from FOREIGN-NAME `[...]` syntax (§3.3.5.2) */
    foreign?: boolean;
}

export interface ImplicitWithObjectExpression extends Expression {
    type: 'ImplicitWithObjectExpression';
    property: Identifier;
}

export interface NumberLiteral extends Expression {
    type: 'NumberLiteral';
    value: number;
    /** VBA type declaration suffix: % Integer, & Long, ! Single, # Double, @ Currency, ^ LongLong */
    typeSuffix?: '%' | '&' | '!' | '#' | '@' | '^';
    /** true when the literal was written with a decimal point or exponent (e.g. 1.0, 1E5) */
    isFloat?: true;
}

export interface StringLiteral extends Expression {
    type: 'StringLiteral';
    value: string;
}

export interface DateLiteral extends Expression {
    type: 'DateLiteral';
    value: string;
}

export interface BinaryExpression extends Expression {
    type: 'BinaryExpression';
    operator: string;
    left: Expression;
    right: Expression;
}

export interface UnaryExpression extends Expression {
    type: 'UnaryExpression';
    operator: string;
    argument: Expression;
}

export class ParseError extends Error {
    constructor(message: string, public readonly line: number, public readonly column: number) {
        super(message);
    }
}

export class Parser {
    private tokens: Token[];
    private pos: number = 0;
    private readonly parseAsClass: string | undefined;
    private readonly _diagnostics: ParseDiagnostic[] = [];

    // ---------------------------------------------------------------
    // §3.3.5.2 contextual keyword Sets
    //
    // IDENTIFIER = <any lex-identifier that is not a reserved-identifier>
    // The following keywords are NOT listed in any <reserved-identifier>
    // category and are therefore valid IDENTIFIERs outside the specific
    // syntactic positions where they carry dedicated meaning.
    // ---------------------------------------------------------------

    /** file-mode specifiers: Open...For <mode>  (§5.4.5.1 open-stmt) */
    private static readonly CONTEXTUAL_KW_FILE_MODE = new Set<TokenType>([
        TokenType.KeywordOutput,
        TokenType.KeywordAppend,
        TokenType.KeywordRandom,
        TokenType.KeywordBinary,  // also: Option Compare Binary
    ]);

    /** file-access specifiers: Open...Access <access>  (§5.4.5.1 open-stmt) */
    private static readonly CONTEXTUAL_KW_FILE_ACCESS = new Set<TokenType>([
        TokenType.KeywordAccess,
        TokenType.KeywordRead,
    ]);

    /** option-statement modifiers: Option Compare/Base/Explicit/Module  (§5.2.2) */
    private static readonly CONTEXTUAL_KW_OPTION = new Set<TokenType>([
        TokenType.KeywordText,      // Option Compare Text
        TokenType.KeywordCompare,   // Option Compare ...
        TokenType.KeywordExplicit,  // Option Explicit
        TokenType.KeywordBase,      // Option Base 0|1
        TokenType.KeywordModule,    // Option Private Module
    ]);

    /** declare-statement modifiers: Declare [PtrSafe]...Lib...Alias  (§5.2.3.1) */
    private static readonly CONTEXTUAL_KW_DECLARE = new Set<TokenType>([
        TokenType.KeywordLib,
        TokenType.KeywordAlias,
        TokenType.KeywordPtrSafe,
    ]);

    /** standalone statements absent from <statement-keyword>  (§3.3.5.2) */
    private static readonly CONTEXTUAL_KW_STMT_ABSENT = new Set<TokenType>([
        TokenType.KeywordStep,         // For...Step  (not in <marker-keyword>)
        TokenType.KeywordMid,          // Mid-stmt mode-specifier (§5.4.3.8); also Mid/Mid$/MidB function
        TokenType.KeywordKill,         // Kill <pathname>
        TokenType.KeywordWidth,        // Width #<filenumber>, <width>
        TokenType.KeywordLine,         // Line Input #<filenumber>, <variable>
        TokenType.KeywordReset,        // Reset
        TokenType.KeywordAppActivate,  // AppActivate <title>
        TokenType.KeywordSendKeys,     // SendKeys <keys>
    ]);

    /** Structural keywords that appear in dedicated lexer token types but are NOT
     *  listed in any §3.3.5.2 reserved-identifier category.
     *  They are valid IDENTIFIERs in declaration and expression contexts.
     *  Note: statement-level dispatch (e.g. Class → parseClassDeclaration) still has
     *  priority; assignment `Class = x` is disambiguated by a `!= OperatorEquals` guard. */
    private static readonly CONTEXTUAL_KW_STRUCTURAL = new Set<TokenType>([
        TokenType.KeywordClass,       // Class declaration keyword (not in spec's statement-keyword list)
        TokenType.KeywordCollection,  // Built-in object type name (not in reserved-identifier list)
        TokenType.KeywordError,       // On Error construct element (not standalone reserved)
        TokenType.KeywordProperty,    // Property Get/Set/Let keyword (not in spec's statement-keyword list)
    ]);

    /** Union of all contextual keyword groups above.
     *  Tokens in this Set are valid IDENTIFIERs in Dim declarations,
     *  expression context, and assignment statements. */
    private static readonly CONTEXTUAL_KW = new Set<TokenType>([
        ...Parser.CONTEXTUAL_KW_FILE_MODE,
        ...Parser.CONTEXTUAL_KW_FILE_ACCESS,
        ...Parser.CONTEXTUAL_KW_OPTION,
        ...Parser.CONTEXTUAL_KW_DECLARE,
        ...Parser.CONTEXTUAL_KW_STMT_ABSENT,
        ...Parser.CONTEXTUAL_KW_STRUCTURAL,
    ]);

    /** <statement-keyword> tokens that are <reserved-identifier> per §3.3.5.2 but whose
     *  enum values fall in [KeywordBase, KeywordAddressOf].
     *  These CANNOT be used as module-level procedure names (parseProcedureDeclaration rejects them).
     *  They ARE permitted as member names in expression context (obj.Print, ws.Get, obj.Open) via
     *  parsePrimary, which is why they appear in COMPAT_KW_EXPR. */
    private static readonly STATEMENT_KW_RESERVED = new Set<TokenType>([
        TokenType.KeywordOpen,
        TokenType.KeywordClose,
        TokenType.KeywordInput,
        TokenType.KeywordPrint,
        TokenType.KeywordPut,
        TokenType.KeywordWrite,
        TokenType.KeywordLock,
        TokenType.KeywordSeek,
        TokenType.KeywordUnlock,
        TokenType.KeywordEvent,
        TokenType.KeywordRaiseEvent,
        TokenType.KeywordImplements,
    ]);

    /** <statement-keyword> tokens permitted as member names in parsePrimary (obj.Print, ws.Get,
     *  obj.Open, etc.) per §3.3.5.2 "unrestricted-name" rule.
     *  In parseStatementInner these keywords are all dispatched unconditionally before the
     *  identifier branch, so the COMPAT_KW_EXPR entries there are never reached — they exist
     *  solely for parsePrimary dot-member-access context. */
    private static readonly COMPAT_KW_EXPR = new Set<TokenType>([
        TokenType.KeywordSeek,
        TokenType.KeywordInput,
        TokenType.KeywordPrint,
        TokenType.KeywordPut,
        TokenType.KeywordGet,    // enum < KeywordBase so also rejected by range check in parseProcedureDeclaration
        TokenType.KeywordLock,
        TokenType.KeywordUnlock,
        TokenType.KeywordOpen,
        TokenType.KeywordClose,
        // Non-reserved statement keywords: can be user-defined proc names (e.g. Function Kill()).
        // Included here so "Kill = v", "Reset = v", "Width = v" fall through to identifier branch.
        TokenType.KeywordKill,
        TokenType.KeywordReset,
        TokenType.KeywordWidth,
    ]);

    private readonly errorRecovery: boolean;
    private readonly sourceLines: string[] | undefined;

    /** Returns true if token is a valid IDENTIFIER per §3.3.5.2:
     *  either a plain lex-identifier or a contextual keyword that is not reserved. */
    private isIdentifier(token: Token): boolean {
        return token.type === TokenType.Identifier
            || token.type === TokenType.ForeignName
            || Parser.CONTEXTUAL_KW.has(token.type);
    }

    constructor(tokens: Token[], options: {
        parseAsClass?: string;
        errorRecovery?: boolean;
        sourceLines?: string[];
    } = {}) {
        this.tokens = tokens;
        this.parseAsClass = options.parseAsClass;
        this.errorRecovery = options.errorRecovery ?? false;
        this.sourceLines = options.sourceLines;
    }

    // Keywords can appear as property/class names in VBA (e.g. obj.Property, New Collection)
    private isNameToken(token: Token): boolean {
        return token.type === TokenType.Identifier
            || token.type === TokenType.ForeignName
            || (token.type >= TokenType.KeywordFor && token.type <= TokenType.KeywordAddressOf);
    }

    /**
     * Returns true if the current token stream contains a file-I/O Open statement pattern
     * ("For <file-mode>") on the current logical line.
     * Open "path" For Input|Output|Append|Random|Binary ... As #n  →  true
     * Open()  /  Open = value  /  Open x  (user-defined call)       →  false
     *
     * This syntactic check lets the parser disambiguate without a pre-scan:
     * VBA §3.3.5.3 — user-defined procedures (priority 2) > built-in statement keywords (priority 3).
     */
    private recordError(message: string, token: Token): void {
        const pos: Position = { line: token.line, column: token.column };
        this._diagnostics.push({ message, loc: { start: pos, end: pos }, severity: 'error' });
    }

    private throwMissingRParen(): never {
        const peek = this.peek();
        if (peek.type === TokenType.Newline) {
            const prevToken = this.tokens[Math.max(0, this.pos - 1)];
            if (prevToken && this.isContinuationEndToken(prevToken.type)) {
                this.throwError(
                    `行継続文字 '_' が必要です（'${prevToken.value}' の後で改行されています）`,
                    peek
                );
            }
        }
        this.throwError(`Parse error: Expected ')' at line ${peek.line} `);
    }

    private isContinuationEndToken(type: TokenType): boolean {
        return type === TokenType.OperatorPlus
            || type === TokenType.OperatorMinus
            || type === TokenType.OperatorMultiply
            || type === TokenType.OperatorDivide
            || type === TokenType.OperatorIntDivide
            || type === TokenType.OperatorPower
            || type === TokenType.OperatorAmpersand
            || type === TokenType.OperatorEquals
            || type === TokenType.OperatorNotEquals
            || type === TokenType.OperatorLessThan
            || type === TokenType.OperatorGreaterThan
            || type === TokenType.OperatorLessThanOrEqual
            || type === TokenType.OperatorGreaterThanOrEqual
            || type === TokenType.OperatorComma
            || type === TokenType.OperatorLParen
            || type === TokenType.KeywordAnd
            || type === TokenType.KeywordOr
            || type === TokenType.KeywordXor
            || type === TokenType.KeywordEqv
            || type === TokenType.KeywordImp
            || type === TokenType.KeywordMod
            || type === TokenType.KeywordLike
            || type === TokenType.KeywordIs
            || type === TokenType.KeywordNot;
    }

    private tokenDisplay(value: string): string {
        return value
            .replace(/\n/g, '<改行>')
            .replace(/\r/g, '<CR>')
            .replace(/\t/g, '<タブ>');
    }

    private throwError(message: string, token?: Token): never {
        const peek = this.peek();
        const t = token ?? (peek.type !== TokenType.EOF ? peek : this.tokens[Math.max(0, this.pos - 1)]);
        const snippet = this.sourceLines?.[t.line - 1];
        const fullMessage = snippet !== undefined ? `${message}\n  > ${snippet}` : message;
        throw new ParseError(fullMessage, t.line, t.column);
    }

    private syncToNextTopLevelStatement(): void {
        // After a parse error, skip tokens until we find what looks like
        // the start of a new top-level statement. We always advance at least
        // one token to guarantee forward progress, then skip to the next
        // Newline. Additionally consume any "End Sub/Function/Property"
        // terminators so they don't orphan at the top level and cause
        // spurious cascading errors.
        while (
            this.peek().type !== TokenType.EOF &&
            this.peek().type !== TokenType.Newline
        ) {
            this.advance();
        }
        // Skip past any End Sub / End Function / End Property / End Property
        // that were left un-consumed by the failed parse. Without this they
        // reach the outer parse loop and are misinterpreted as starting a
        // new Sub/Function/Property declaration.
        while (this.isAtEndTerminator()) {
            this.advance(); // consume 'End'
            this.advance(); // consume 'Sub' / 'Function' / etc.
            while (
                this.peek().type !== TokenType.EOF &&
                this.peek().type !== TokenType.Newline
            ) {
                this.advance();
            }
        }
    }

    private peek(offset: number = 0): Token {
        if (this.pos + offset >= this.tokens.length) {
            return this.tokens[this.tokens.length - 1]; // EOF
        }
        return this.tokens[this.pos + offset];
    }
    private advance(): Token {
        const token = this.peek();
        this.pos++;
        return token;
    }

    private parseOptionCompareStatement(): OptionCompareStatement {
        const modeToken = this.advance();
        let mode: 'Binary' | 'Text';
        if (modeToken.type === TokenType.KeywordBinary) {
            mode = 'Binary';
        } else if (modeToken.type === TokenType.KeywordText) {
            mode = 'Text';
        } else {
            this.throwError(`Parse error: Expected 'Binary' or 'Text' after 'Option Compare' at line ${modeToken.line}`);
        }
        return { type: 'OptionCompareStatement', mode };
    }

    private validateParameterOrder(params: Parameter[]): void {
        let seenOptional = false;
        for (const p of params) {
            if (p.isParamArray) break; // ParamArray is always last; no check needed after it
            if (p.isOptional) {
                seenOptional = true;
            } else if (seenOptional) {
                const line = (p as any).loc?.start?.line ?? this.peek().line;
                this.throwError(`Compile error at line ${line}: Non-optional parameter '${p.name}' cannot follow an Optional parameter`);
            }
        }
    }

    private parseParameter(): Parameter {
        let isByVal = false;
        let isOptional = false;

        if (this.match(TokenType.KeywordOptional)) {
            isOptional = true;
        }

        let isParamArray = false;
        if (this.match(TokenType.KeywordParamArray)) {
            isParamArray = true;
        }

        let hasPassingModifier = false;
        if (this.peek().type === TokenType.KeywordByVal || this.peek().type === TokenType.KeywordByRef) {
            hasPassingModifier = true;
            isByVal = this.advance().type === TokenType.KeywordByVal;
        }

        const token = this.peek();
        if (token.type !== TokenType.Identifier && !Parser.CONTEXTUAL_KW.has(token.type) && (token.type < TokenType.KeywordBase || token.type > TokenType.KeywordAddressOf)) {
            this.throwError(`Parse error at line ${token.line}: Expected parameter name (Found ${this.tokenDisplay(token.value)})`);
        }
        const nameToken = this.advance();

        // Consume optional () for arrays or ParamArray
        let isArray = false;
        if (this.match(TokenType.OperatorLParen)) {
            this.consume(TokenType.OperatorRParen, "Expected ')' after parameter name");
            isArray = true;
        }

        let paramType: string | undefined;
        if (this.match(TokenType.KeywordAs)) {
            paramType = this.advance().value; // capture type name (first token)
            // Handle qualified type names: MSForms.ReturnInteger, Module.TypeName, etc.
            if (this.peek().type === TokenType.OperatorDot) {
                this.advance(); // consume '.'
                paramType += '.' + this.advance().value;
            }
        }

        let defaultValue: Expression | undefined;
        if (this.match(TokenType.OperatorEquals)) {
            defaultValue = this.parseExpression();
        }

        return {
            type: 'Parameter', name: nameToken.value, isByVal, hasPassingModifier, isOptional, isParamArray, isArray, paramType, defaultValue,
            loc: { start: { line: nameToken.line, column: nameToken.column }, end: { line: nameToken.line, column: nameToken.column + nameToken.value.length } },
        };
    }

    private parseDefDirective(): DefDirective {
        const defToken = this.advance(); // e.g. 'DefInt'
        const keyword = defToken.value.toLowerCase().replace(/[$%&#@!^]$/, '');
        const typeMap: Record<string, string> = {
            defbool: 'Boolean', defbyte: 'Byte', defint: 'Integer',
            deflng: 'Long', deflnglng: 'LongLong', deflngptr: 'LongPtr',
            defsng: 'Single', defdbl: 'Double', defcur: 'Currency',
            defdate: 'Date', defstr: 'String', defobj: 'Object', defvar: 'Variant',
        };
        const vbaType = typeMap[keyword] ?? 'Variant';
        const ranges: { from: string; to: string }[] = [];
        do {
            const fromTok = this.advance();
            const from = fromTok.value.charAt(0).toLowerCase();
            let to = from;
            if (this.peek().type === TokenType.OperatorMinus) {
                this.advance(); // '-'
                const toTok = this.advance();
                to = toTok.value.charAt(0).toLowerCase();
            }
            ranges.push({ from, to });
        } while (this.match(TokenType.OperatorComma));
        return { type: 'DefDirective', vbaType, ranges };
    }

    private parseAttributeStatement(): AttributeStatement {
        this.advance(); // 'Attribute'
        const name = this.advance().value;
        this.consume(TokenType.OperatorEquals, "Expected '=' after Attribute name");
        const value = this.parseExpression();
        return { type: 'AttributeStatement', name, value };
    }

    private parseOpenStatement(): OpenStatement {
        this.advance(); // 'Open'
        const path = this.parseExpression();
        this.consume(TokenType.KeywordFor, "Expected 'For' in Open statement");
        
        let mode: any = 'Random';
        const modeToken = this.advance();
        switch (modeToken.type) {
            case TokenType.KeywordInput: mode = 'Input'; break;
            case TokenType.KeywordOutput: mode = 'Output'; break;
            case TokenType.KeywordAppend: mode = 'Append'; break;
            case TokenType.KeywordRandom: mode = 'Random'; break;
            case TokenType.KeywordBinary: mode = 'Binary'; break;
            default: this.throwError(`Parse error: Invalid Open mode '${modeToken.value}' at line ${modeToken.line}`);
        }

        let access: any = undefined;
        if (this.match(TokenType.KeywordAccess)) {
            const first = this.advance();
            if (first.type === TokenType.KeywordRead) {
                if (this.match(TokenType.KeywordWrite)) {
                    access = 'Read Write';
                } else {
                    access = 'Read';
                }
            } else if (first.type === TokenType.KeywordWrite) {
                access = 'Write';
            }
        }

        let lock: any = undefined;
        if (this.match(TokenType.KeywordLock)) {
            const first = this.advance();
            if (first.type === TokenType.KeywordShared) {
                lock = 'Shared';
            } else if (first.type === TokenType.KeywordRead) {
                if (this.match(TokenType.KeywordWrite)) {
                    lock = 'Lock Read Write';
                } else {
                    lock = 'Lock Read';
                }
            } else if (first.type === TokenType.KeywordWrite) {
                lock = 'Lock Write';
            }
        }

        this.consume(TokenType.KeywordAs, "Expected 'As' in Open statement");
        this.match(TokenType.OperatorHash); // optional #
        const fileNumber = this.parseExpression();

        return { type: 'OpenStatement', path, mode, access, lock, fileNumber };
    }

    private parseCloseStatement(): CloseStatement {
        this.advance(); // 'Close'
        const fileNumbers: Expression[] = [];
        while (!this.isAtTerminator()) {
            this.match(TokenType.OperatorHash); // optional #
            fileNumbers.push(this.parseExpression());
            if (!this.match(TokenType.OperatorComma)) break;
        }
        return { type: 'CloseStatement', fileNumbers };
    }

    private parsePrintStatement(): PrintStatement {
        this.advance(); // 'Print'
        this.consume(TokenType.OperatorHash, "Expected '#' in Print statement");
        const fileNumber = this.parseExpression();
        this.consume(TokenType.OperatorComma, "Expected ',' after file number in Print statement");

        const expressions: any[] = [];
        while (!this.isAtTerminator()) {
            if (this.match(TokenType.KeywordSpc)) {
                this.consume(TokenType.OperatorLParen, "Expected '(' after Spc");
                expressions.push({ type: 'Spc', val: this.parseExpression() });
                this.consume(TokenType.OperatorRParen, "Expected ')' after Spc");
            } else if (this.match(TokenType.KeywordTab)) {
                if (this.match(TokenType.OperatorLParen)) {
                    expressions.push({ type: 'Tab', val: this.parseExpression() });
                    this.consume(TokenType.OperatorRParen, "Expected ')' after Tab");
                } else {
                    expressions.push('Tab');
                }
            } else if (this.peek().type === TokenType.OperatorComma) {
                this.advance();
                expressions.push('Comma');
            } else if (this.peek().type === TokenType.OperatorSemicolon) {
                this.advance();
                expressions.push('Semicolon');
            } else {
                expressions.push(this.parseExpression());
            }
            
            // Check for trailing separators
            if (this.isAtTerminator()) break;
        }
        return { type: 'PrintStatement', fileNumber, expressions };
    }

    private parseLineInputStatement(): LineInputStatement {
        this.advance(); // 'Line'
        this.consume(TokenType.KeywordInput, "Expected 'Input' after 'Line'");
        this.consume(TokenType.OperatorHash, "Expected '#' in Line Input statement");
        const fileNumber = this.parseExpression();
        this.consume(TokenType.OperatorComma, "Expected ',' after file number");
        const variable = this.parsePrimary();
        if (variable.type !== 'Identifier') {
            this.throwError(`Parse error: Expected variable name in Line Input at line ${this.peek().line}`);
        }
        return { type: 'LineInputStatement', fileNumber, variable: variable as Identifier };
    }

    private parsePutStatement(): PutStatement {
        this.advance(); // 'Put'
        this.consume(TokenType.OperatorHash, "Expected '#' in Put statement");
        const fileNumber = this.parseExpression();
        this.consume(TokenType.OperatorComma, "Expected ',' after file number");
        
        let recordNumber: Expression | undefined = undefined;
        if (this.peek().type !== TokenType.OperatorComma) {
            recordNumber = this.parseExpression();
        }
        this.consume(TokenType.OperatorComma, "Expected second ',' in Put statement");
        const data = this.parseExpression();

        return { type: 'PutStatement', fileNumber, recordNumber, data };
    }

    private parseKillStatement(): KillStatement {
        this.advance(); // 'Kill'
        const path = this.parseExpression();
        return { type: 'KillStatement', path };
    }

    private parseWriteStatement(): WriteStatement {
        this.advance(); // 'Write'
        this.consume(TokenType.OperatorHash, "Expected '#' in Write statement");
        const fileNumber = this.parseExpression();
        
        const items: Expression[] = [];
        if (this.match(TokenType.OperatorComma)) {
            while (!this.isAtTerminator()) {
                items.push(this.parseExpression());
                if (!this.match(TokenType.OperatorComma)) break;
            }
        }
        
        return { type: 'WriteStatement', fileNumber, items };
    }

    private parseInputStatement(): InputStatement {
        this.advance(); // 'Input'
        this.consume(TokenType.OperatorHash, "Expected '#' in Input statement");
        const fileNumber = this.parseExpression();
        this.consume(TokenType.OperatorComma, "Expected ',' in Input statement");
        
        const variables: Expression[] = [];
        while (!this.isAtTerminator()) {
            variables.push(this.parseExpression());
            if (!this.match(TokenType.OperatorComma)) break;
        }
        
        return { type: 'InputStatement', fileNumber, variables };
    }

    private parseGetStatement(): GetStatement {
        this.advance(); // 'Get'
        this.consume(TokenType.OperatorHash, "Expected '#' in Get statement");
        const fileNumber = this.parseExpression();
        this.consume(TokenType.OperatorComma, "Expected ',' in Get statement");
        
        let recordNumber: Expression | undefined = undefined;
        if (this.peek().type !== TokenType.OperatorComma) {
            recordNumber = this.parseExpression();
        }
        this.consume(TokenType.OperatorComma, "Expected ',' in Get statement");
        const variable = this.parseExpression();
        
        return { type: 'GetStatement', fileNumber, recordNumber, variable };
    }

    private parseSeekStatement(): SeekStatement {
        this.advance(); // 'Seek'
        this.consume(TokenType.OperatorHash, "Expected '#' in Seek statement");
        const fileNumber = this.parseExpression();
        this.consume(TokenType.OperatorComma, "Expected ',' in Seek statement");
        const position = this.parseExpression();
        
        return { type: 'SeekStatement', fileNumber, position };
    }

    private parseResetStatement(): ResetStatement {
        this.advance(); // 'Reset'
        return { type: 'ResetStatement' };
    }

    private parseDeclareStatement(): DeclareStatement {
        this.advance(); // 'Declare'
        let isPtrSafe = false;
        if (this.peek().type === TokenType.KeywordPtrSafe) {
            this.advance();
            isPtrSafe = true;
        }

        let isSub = false;
        if (this.peek().type === TokenType.KeywordSub) {
            this.advance();
            isSub = true;
        } else if (this.peek().type === TokenType.KeywordFunction) {
            this.advance();
            isSub = false;
        } else {
            this.throwError(`Parser error: Expected Sub or Function after Declare at line ${this.peek().line}`);
        }

        const name = this.advance().value;

        if (this.peek().type !== TokenType.KeywordLib) {
            this.throwError(`Parser error: Expected Lib after Declare name at line ${this.peek().line}`);
        }
        this.advance();
        const libName = this.advance().value.replace(/^"|"$/g, '');

        let aliasName: string | undefined;
        if (this.peek().type === TokenType.KeywordAlias) {
            this.advance();
            aliasName = this.advance().value.replace(/^"|"$/g, '');
        }

        let parameters: Parameter[] = [];
        if (this.peek().type === TokenType.OperatorLParen) {
            this.advance();
            while (this.peek().type !== TokenType.OperatorRParen && this.peek().type !== TokenType.EOF) {
                parameters.push(this.parseParameter());
                if (this.peek().type === TokenType.OperatorComma) {
                    this.advance();
                }
            }
            this.consume(TokenType.OperatorRParen, "Expected ')' after declare parameters");
            this.validateParameterOrder(parameters);
        }

        let returnType: string | undefined;
        if (!isSub) {
            if (this.peek().type === TokenType.KeywordAs) {
                this.advance();
                returnType = this.advance().value;
            }
        }

        return {
            type: 'DeclareStatement',
            isPtrSafe,
            isSub,
            name,
            libName,
            aliasName,
            parameters,
            returnType
        };
    }

    private isAtEndTerminator(): boolean {
        const token = this.peek();
        if (token.type !== TokenType.KeywordEnd) return false;
        const next = this.peek(1);
        return (
            next.type === TokenType.KeywordSub ||
            next.type === TokenType.KeywordFunction ||
            next.type === TokenType.KeywordIf ||
            next.type === TokenType.KeywordSelect ||
            next.type === TokenType.KeywordWith ||
            next.type === TokenType.KeywordType ||
            next.type === TokenType.KeywordEnum ||
            next.type === TokenType.KeywordProperty ||
            next.type === TokenType.KeywordClass
        );
    }

    private match(expectedType: TokenType): boolean {
        if (this.peek().type === expectedType) {
            this.advance();
            return true;
        }
        return false;
    }

    // Returns true for operators that can ONLY appear in a binary position (never as unary or
    // argument starters). When one of these follows a greedily-parsed CallExpression in a
    // statement context, it signals that the `(args)` was actually part of the argument expression.
    private isBinaryOnlyOperator(type: TokenType): boolean {
        return type === TokenType.OperatorMultiply
            || type === TokenType.OperatorDivide
            || type === TokenType.OperatorIntDivide
            || type === TokenType.OperatorPower
            || type === TokenType.OperatorPlus
            || type === TokenType.OperatorMinus
            || type === TokenType.OperatorAmpersand
            || type === TokenType.OperatorNotEquals
            || type === TokenType.OperatorLessThan
            || type === TokenType.OperatorGreaterThan
            || type === TokenType.OperatorLessThanOrEqual
            || type === TokenType.OperatorGreaterThanOrEqual
            || type === TokenType.KeywordMod;
    }

    // Returns true if there is whitespace between the previous token and the current token (same line).
    // Used to distinguish `Foo(arg)` (no space = function-call postfix) from
    // `Foo (arg)*x` (space = argument expression; `(arg)*x` must be parsed as one expression).
    private hasSpaceBeforeCurrentToken(): boolean {
        if (this.pos <= 0) return false;
        const cur = this.tokens[this.pos];
        const prev = this.tokens[this.pos - 1];
        if (prev.line !== cur.line) return false;
        return cur.column > prev.column + prev.value.length;
    }

    private isAtTerminator(): boolean {
        const type = this.peek().type;
        return type === TokenType.Newline || type === TokenType.EOF || type === TokenType.OperatorColon;
    }

    private consume(expectedType: TokenType, message: string): Token {
        if (this.peek().type === expectedType) {
            return this.advance();
        }
        this.throwError(`Parse error at line ${this.peek().line}: ${message}`);
    }

    private skipNewlines() {
        while (this.match(TokenType.Newline)) { /* skip */ }
    }

    // Build a SourceLocation spanning from start token to end token (inclusive).
    private exprLoc(start: Token, end: Token): SourceLocation {
        return {
            start: { line: start.line, column: start.column },
            end: { line: end.line, column: end.column + end.value.length },
        };
    }

    // Returns true if the token can serve as a name (identifier or any keyword).
    // VBA §5.2.3.3 allows reserved words as UDT/Enum member names (reserved-name-member-dcl).
    // MS-VBAL §3.3.5: name-start は Unicode-Letter (Lo を含む) も許容するため Unicode フラグを使う。
    private isWordToken(t: Token): boolean {
        return /^[\p{L}\p{Nl}_][\p{L}\p{Nl}\p{Mn}\p{Mc}\p{Nd}\p{Pc}_]*$/u.test(t.value);
    }

    // Create an Identifier node with location from the given token.
    private makeIdentifier(token: Token): Identifier {
        return {
            type: 'Identifier',
            name: token.value,
            loc: {
                start: { line: token.line, column: token.column },
                end: { line: token.line, column: token.column + token.value.length },
            },
        };
    }

    public parse(): Program {

        if (this.parseAsClass) {
            const classDecl = this.parseClassBody(this.parseAsClass, false);
            return { type: 'Program', body: [classDecl], diagnostics: this._diagnostics };
        }

        const program: Program = {
            type: 'Program',
            body: [],
            diagnostics: this._diagnostics,
        };

        this.skipNewlines();
        while (this.peek().type !== TokenType.EOF) {
            const startPos = this.pos;
            const startToken = this.peek();
            try {
                const stmt = this.parseStatement();
                if (stmt) {
                    program.body.push(stmt);
                }
            } catch (e) {
                if (!this.errorRecovery) throw e;
                const msg = e instanceof Error ? e.message : String(e);
                if (e instanceof ParseError) {
                    const pos: Position = { line: e.line, column: e.column };
                    this._diagnostics.push({ message: msg, loc: { start: pos, end: pos }, severity: 'error' });
                } else {
                    this.recordError(msg, startToken);
                }
                this.syncToNextTopLevelStatement();
            }
            // Defensive: ensure forward progress even if a parser returns without advancing.
            // When parseStatement() returns null due to isAtEndTerminator() (e.g. "End Sub"
            // orphaned by error recovery), skip the full "End <keyword>" pair so it is not
            // misinterpreted as starting a new procedure declaration.
            if (this.pos === startPos && this.peek().type !== TokenType.EOF) {
                if (this.isAtEndTerminator()) {
                    this.advance(); // consume 'End'
                    this.advance(); // consume 'Sub' / 'Function' / 'Property' / etc.
                } else {
                    this.advance();
                }
            }
            this.skipNewlines();
        }

        return program;
    }

    public parseExpressionPublic(): Expression {
        this.skipNewlines();
        return this.parseExpression();
    }

    // §5.4.1: block-statement = statement EOS
    // checkEOS=false is used only for inline-If bodies where Else/EOF terminate instead of EOS.
    private parseStatement(checkEOS = true): Statement | null {
        this.skipNewlines();
        const startToken = this.peek();
        const stmt = this.parseStatementInner();
        if (stmt !== null) {
            // ラベル文（label:）の直後は EOS 不要 — 同一行に次の文が続く
            const isLabel = (stmt as any).type === 'LabelStatement';
            if (checkEOS && !isLabel && !this.isAtTerminator()) {
                this.throwError(`Parse error: unexpected token '${this.peek().value}' after statement at line ${this.peek().line}`);
            }
            const endToken = this.tokens[this.pos - 1];
            if (startToken.line !== undefined) {
                stmt.line = startToken.line;
                const startPos: Position = { line: startToken.line, column: startToken.column };
                const endPos: Position = endToken && endToken.line !== undefined
                    ? { line: endToken.line, column: endToken.column + endToken.value.length }
                    : startPos;
                stmt.loc = { start: startPos, end: endPos };
            }
        }
        return stmt;
    }

    /**
     * Parse an identifier, method call, or assignment statement.
     * Called from the identifier branch of parseStatementInner AND from the ARCH-1
     * user-proc override path (when a built-in keyword is user-defined as a procedure).
     */
    private parseIdentifierOrCallStatement(): Statement {
        const token = this.peek();

        // Label check: "Identifier:" or contextual-keyword ":"
        if ((token.type === TokenType.Identifier || Parser.CONTEXTUAL_KW.has(token.type)) &&
                this.pos + 1 < this.tokens.length &&
                this.tokens[this.pos + 1].type === TokenType.OperatorColon) {
            const labelName = token.value;
            this.advance(); // consume Identifier / keyword
            this.advance(); // consume ':'
            return { type: 'LabelStatement', label: labelName } as any;
        }

        // Line number label: "42" or "42:"
        if (token.type === TokenType.Number) {
            const labelName = token.value;
            this.advance(); // consume Number
            this.match(TokenType.OperatorColon); // optional colon
            return { type: 'LabelStatement', label: labelName } as any;
        }

        // Unify assignment, array access, method call
        const savedPos = this.pos;
        const expr = this.parsePrimary(); // will parse `foo`, `foo()`, `foo.bar`, `arr(0)` etc

        if (this.match(TokenType.OperatorEquals)) {
            return {
                type: 'AssignmentStatement',
                left: expr,
                right: this.parseExpression()
            } as AssignmentStatement;
        } else {
            // If parsePrimary() greedily consumed `(args)` as a function call postfix but the
            // next token is a binary-only operator (cannot start an argument expression), the
            // `(...)` was actually the start of the argument expression, not a call argument list.
            // Example: `Debug.Print (1+2)*3` — `(1+2)` was mis-parsed as the argument; `*3` is
            // the continuation. Backtrack and re-parse with stopBeforeSpacedLParen=true so that
            // `(1+2)*3` is parsed as a single expression argument.
            if (expr.type === 'CallExpression' && this.isBinaryOnlyOperator(this.peek().type)) {
                this.pos = savedPos;
                const callee = this.parsePrimary(/* stopBeforeSpacedLParen= */ true);
                const args: Expression[] = [this.parseCallArgument()];
                while (this.match(TokenType.OperatorComma)) {
                    args.push(this.parseCallArgument());
                }
                return { type: 'CallStatement', expression: { type: 'CallExpression', callee, args } } as CallStatement;
            }

            // If it's not an assignment, maybe it's a CallStatement with arguments separated by comma
            const args: Expression[] = [];
            // Check if there are args on the same line
            if (
                this.peek().type !== TokenType.Newline &&
                this.peek().type !== TokenType.EOF &&
                this.peek().type !== TokenType.KeywordElse &&
                this.peek().type !== TokenType.KeywordElseIf &&
                this.peek().type !== TokenType.KeywordEnd &&
                this.peek().type !== TokenType.KeywordNext &&
                this.peek().type !== TokenType.KeywordLoop
            ) {
                args.push(this.parseCallArgument());
                while (this.match(TokenType.OperatorComma)) {
                    args.push(this.parseCallArgument());
                }
            }

            if (args.length > 0) {
                return { type: 'CallStatement', expression: { type: 'CallExpression', callee: expr, args } } as CallStatement;
            } else if (expr.type === 'CallExpression') {
                const callExpr = expr as CallExpression;
                // identifier() with empty parens in statement position (no Call keyword)
                // is always a syntax error — use `identifier`, `identifier arg`, or
                // `Call identifier()` instead.
                // FOREIGN-NAME [identifier]() is exempt: used to call COM methods by reserved names.
                if (callExpr.args.length === 0 && callExpr.callee.type === 'Identifier' &&
                        !(callExpr.callee as Identifier).foreign) {
                    const line = callExpr.loc?.start.line ?? this.peek().line;
                    this.throwError(`Parse error: syntax error at line ${line}`);
                }
                return { type: 'CallStatement', expression: callExpr } as CallStatement;
            } else {
                // Call matched without parens e.g. `MainLoop`
                return { type: 'CallStatement', expression: { type: 'CallExpression', callee: expr, args: [] } } as CallStatement;
            }
        }
    }

    private parseStatementInner(): Statement | null {
        const token = this.peek();

        // Label check must come before keyword-specific dispatches.
        // Contextual keywords (Error, Class, Property, etc.) are valid label names when followed by ':'.
        if (Parser.CONTEXTUAL_KW.has(token.type) &&
                this.pos + 1 < this.tokens.length &&
                this.tokens[this.pos + 1].type === TokenType.OperatorColon) {
            const labelName = token.value;
            this.advance(); // consume keyword
            this.advance(); // consume ':'
            return { type: 'LabelStatement', label: labelName } as any;
        }

        if (token.type === TokenType.KeywordPublic || token.type === TokenType.KeywordPrivate || token.type === TokenType.KeywordFriend) {
            const scope = this.advance().value.toLowerCase() as 'public' | 'private' | 'friend';
            const next = this.peek();
            if (next.type === TokenType.KeywordSub || next.type === TokenType.KeywordFunction || next.type === TokenType.KeywordProperty) {
                return this.parseProcedureDeclaration(scope, false);
            }
            if (next.type === TokenType.KeywordConst) {
                const stmt = this.parseConstDeclaration();
                (stmt as any).scope = scope;
                return stmt;
            }
            if (next.type === TokenType.KeywordType) {
                const stmt = this.parseTypeDeclaration();
                (stmt as any).scope = scope;
                return stmt;
            }
            if (next.type === TokenType.KeywordEnum) {
                const stmt = this.parseEnumDeclaration();
                (stmt as any).scope = scope;
                return stmt;
            }
            if (next.type === TokenType.KeywordDeclare) {
                const stmt = this.parseDeclareStatement();
                (stmt as any).scope = scope;
                return stmt;
            }
            if (next.type === TokenType.KeywordEvent) {
                return this.parseEventDeclaration(scope as 'public' | 'private' | 'friend');
            }
            // Public/Private on Dim/Const — handle as variable declaration
            const stmt = this.parseDimStatement(false, true);
            if (stmt) {
                (stmt as any).scope = scope;
            }
            return stmt;
        }

        if (token.type === TokenType.KeywordFor) {
            return this.parseForStatement();
        } else if (token.type === TokenType.KeywordIf) {
            return this.parseIfStatement();
        } else if (token.type === TokenType.KeywordDo) {
            return this.parseDoWhileStatement();
        } else if (token.type === TokenType.KeywordWhile) {
            return this.parseWhileStatement();
        } else if (token.type === TokenType.KeywordSub || token.type === TokenType.KeywordFunction ||
                   (token.type === TokenType.KeywordProperty && this.peek(1).type !== TokenType.OperatorEquals)) {
            return this.parseProcedureDeclaration();
        } else if (token.type === TokenType.KeywordStatic) {
            this.advance(); // consume 'Static'
            const next = this.peek();
            if (next.type === TokenType.KeywordSub || next.type === TokenType.KeywordFunction || next.type === TokenType.KeywordProperty) {
                return this.parseProcedureDeclaration(undefined, true);
            }
            // Static variable declaration inside a procedure
            return this.parseDimStatement(true, true);
        } else if (token.type === TokenType.KeywordDim) {
            return this.parseDimStatement();
        } else if (token.type === TokenType.KeywordConst) {
            return this.parseConstDeclaration();
        } else if (token.type === TokenType.KeywordStop) {
            this.advance(); // consume 'Stop'
            if (!this.isAtTerminator()) {
                this.throwError(`Parse error: unexpected token after 'Stop'`);
            }
            return { type: 'StopStatement' } as StopStatement;
        } else if (token.type === TokenType.KeywordEnd) {
            if (this.isAtEndTerminator()) {
                return null;
            }
            this.advance(); // consume standalone 'End'
            if (!this.isAtTerminator()) {
                this.throwError(`Parse error: unexpected token after 'End'`);
            }
            return { type: 'EndStatement' } as EndStatement;
        } else if (token.type === TokenType.KeywordGoTo) {
            return this.parseGoToStatement();
        } else if (token.type === TokenType.KeywordSet) {
            return this.parseSetStatement();
        } else if (token.type === TokenType.KeywordOn) {
            const next = this.peek(1);
            if (next.type === TokenType.KeywordError) {
                return this.parseOnErrorStatement();
            } else {
                return this.parseOnGoToSubStatement();
            }
        } else if (token.type === TokenType.KeywordError && this.peek(1).type !== TokenType.OperatorEquals) {
            return this.parseErrorStatement();
        } else if (token.type === TokenType.KeywordGoSub) {
            return this.parseGoSubStatement();
        } else if (token.type === TokenType.KeywordReturn) {
            this.advance(); // consume 'Return'
            return { type: 'ReturnStatement' } as ReturnStatement;
        } else if (token.type === TokenType.KeywordMid && this.hasMidAssignmentAhead()) {
            return this.parseMidStatement();
        } else if (token.type === TokenType.KeywordLSet) {
            return this.parseLSetStatement();
        } else if (token.type === TokenType.KeywordRSet) {
            return this.parseRSetStatement();
        } else if (token.type === TokenType.KeywordExit) {
            return this.parseExitStatement();
        } else if (token.type === TokenType.KeywordErase) {
            return this.parseEraseStatement();
        } else if (token.type === TokenType.KeywordReDim) {
            return this.parseReDimStatement();
        } else if (token.type === TokenType.KeywordResume) {
            return this.parseResumeStatement();
        } else if (token.type === TokenType.KeywordImplements) {
            return this.parseImplementsDirective();
        } else if (token.type === TokenType.KeywordAppActivate && this.peek(1).type !== TokenType.OperatorEquals) {
            return this.parseAppActivateStatement();
        } else if (token.type === TokenType.KeywordSendKeys && this.peek(1).type !== TokenType.OperatorEquals) {
            return this.parseSendKeysStatement();
        } else if (token.type === TokenType.KeywordOption) {
            this.advance(); // 'Option'
            if (this.match(TokenType.KeywordCompare)) {
                return this.parseOptionCompareStatement();
            }
            if (this.match(TokenType.KeywordExplicit)) {
                return { type: 'OptionExplicitStatement' } as OptionExplicitStatement;
            }
            if (this.peek().type === TokenType.Identifier && this.peek().value.toLowerCase() === 'base') {
                this.advance(); // 'Base'
                const baseToken = this.advance();
                if (baseToken.value === '0' || baseToken.value === '1') {
                    return { type: 'OptionBaseStatement', base: parseInt(baseToken.value) } as OptionBaseStatement;
                }
                this.throwError(`Parse error: Option Base must be 0 or 1 at line ${baseToken.line}`);
            }
            if (this.match(TokenType.KeywordPrivate)) {
                this.consume(TokenType.KeywordModule, "Expected 'Module' after 'Option Private'");
                return { type: 'OptionPrivateModuleStatement' } as OptionPrivateModuleStatement;
            }
            return null;
        } else if (token.type === TokenType.KeywordDef) {
            return this.parseDefDirective();
        } else if (token.type === TokenType.KeywordAttribute) {
            return this.parseAttributeStatement();
        } else if (token.type === TokenType.KeywordDeclare) {
            return this.parseDeclareStatement();
        } else if (token.type === TokenType.KeywordSelect) {
            return this.parseSelectCaseStatement();
        } else if (token.type === TokenType.KeywordWith) {
            return this.parseWithStatement();
        } else if (token.type === TokenType.KeywordType) {
            return this.parseTypeDeclaration();
        } else if (token.type === TokenType.KeywordEnum) {
            return this.parseEnumDeclaration();
        } else if (token.type === TokenType.KeywordOpen) {
            // Open is reserved-identifier (§3.3.5.2) — always file I/O Open statement.
            return this.parseOpenStatement();
        } else if (token.type === TokenType.KeywordClose) {
            // Close is reserved-identifier (§3.3.5.2) — always file I/O Close statement.
            return this.parseCloseStatement();
        } else if (token.type === TokenType.KeywordLine && this.peek(1).type !== TokenType.OperatorEquals) {
            return this.parseLineInputStatement();
        } else if (token.type === TokenType.KeywordPrint) {
            // Print is reserved-identifier (§3.3.5.2) — always file I/O. Requires "#" per §5.4.5.9.
            return this.parsePrintStatement();
        } else if (token.type === TokenType.KeywordPut) {
            // Put is reserved-identifier — always file I/O Put statement.
            return this.parsePutStatement();
        } else if (token.type === TokenType.KeywordGet &&
                   this.peek(1).type === TokenType.OperatorHash) {
            // Get (enum < KeywordBase) is reserved by range check. "#" guard kept because Get
            // also appears in "Property Get" header which is parsed before reaching this branch.
            return this.parseGetStatement();
        } else if (token.type === TokenType.KeywordInput) {
            // Input is reserved-identifier — always file I/O. Requires "#" per §5.4.5.7.
            return this.parseInputStatement();
        } else if (token.type === TokenType.KeywordWrite) {
            // Write is reserved-identifier — always file I/O. Requires "#" per §5.4.5.10.
            return this.parseWriteStatement();
        } else if (token.type === TokenType.KeywordSeek) {
            // Seek is reserved-identifier — always file I/O Seek statement.
            return this.parseSeekStatement();
        } else if (token.type === TokenType.KeywordReset && this.peek(1).type !== TokenType.OperatorEquals) {
            // Reset is NOT reserved-identifier — "Reset = v" falls to identifier branch.
            return this.parseResetStatement();
        } else if (token.type === TokenType.KeywordKill && this.peek(1).type !== TokenType.OperatorEquals) {
            // Kill is NOT reserved-identifier — "Kill = v" falls to identifier branch.
            return this.parseKillStatement();
        } else if (token.type === TokenType.KeywordEvent) {
            return this.parseEventDeclaration();
        } else if (token.type === TokenType.KeywordRaiseEvent) {
            return this.parseRaiseEventStatement();
        } else if (token.type === TokenType.KeywordLock) {
            return this.parseLockStatement();
        } else if (token.type === TokenType.KeywordUnlock) {
            return this.parseUnlockStatement();
        } else if (token.type === TokenType.KeywordWidth &&
                   this.peek(1).type === TokenType.OperatorHash) {
            // Width is NOT reserved-identifier. width-statement requires "#" per §5.4.5.
            // Without "#", Width falls to identifier branch (user-defined proc or assignment).
            return this.parseWidthStatement();
        } else if (token.type === TokenType.KeywordClass && this.peek(1).type !== TokenType.OperatorEquals) {
            return this.parseClassDeclaration();
        } else if (token.type === TokenType.KeywordCall) {
            this.advance(); // consume 'Call'
            const expr = this.parsePrimary();
            if (expr.type === 'CallExpression') {
                return { type: 'CallStatement', expression: expr } as CallStatement;
            } else if (expr.type === 'Identifier' || expr.type === 'MemberExpression') {
                // Call ProcName  /  Call Module.ProcName  — no arguments
                return { type: 'CallStatement', expression: { type: 'CallExpression', callee: expr, args: [] } } as CallStatement;
            }
            this.throwError(`Parse error: Expected procedure call after 'Call'`);
        } else if (token.type === TokenType.Identifier || token.type === TokenType.ForeignName ||
                   token.type === TokenType.OperatorDot || token.type === TokenType.KeywordMe ||
                   token.type === TokenType.Number || Parser.CONTEXTUAL_KW.has(token.type) ||
                   Parser.COMPAT_KW_EXPR.has(token.type)) {
            return this.parseIdentifierOrCallStatement();
        } else if (token.type === TokenType.Unknown) {
            this.throwError(`Parse error: Unknown token '${this.tokenDisplay(token.value)}' at line ${token.line}`);
        } else {
            // Unknown or unexpected top-level token
            this.advance();
        }
        return null;
    }

    private parseProcedureDeclaration(scope?: 'public' | 'private' | 'friend', isStatic?: boolean, isClassMember = false): ProcedureDeclaration {
        const isFunction = this.peek().type === TokenType.KeywordFunction;
        const isProperty = this.peek().type === TokenType.KeywordProperty;
        this.advance(); // consume Sub, Function, or Property

        let propertyType: 'get' | 'let' | 'set' | undefined;
        if (isProperty) {
            // consume Get, Let, or Set
            const typeToken = this.advance();
            if (typeToken.type === TokenType.KeywordGet) propertyType = 'get';
            else if (typeToken.type === TokenType.KeywordLet) propertyType = 'let';
            else if (typeToken.type === TokenType.KeywordSet) propertyType = 'set';
            else {
                this.throwError(`Parse error: Expected 'Get', 'Let', or 'Set' after 'Property' at line ${typeToken.line}`);
            }
        }

        const idToken = this.advance();
        if (!this.isIdentifier(idToken) && (idToken.type < TokenType.KeywordBase || idToken.type > TokenType.KeywordAddressOf)) {
            this.throwError(`Parse error at line ${idToken.line}: Expected procedure name (Found ${this.tokenDisplay(idToken.value)})`);
        }
        // §3.3.5.2: statement-keyword is reserved-identifier and cannot be used as a module-level
        // procedure name. Class members are accessed via unrestricted-name (obj.Open is valid).
        if (!isClassMember && Parser.STATEMENT_KW_RESERVED.has(idToken.type)) {
            this.throwError(`Compile error at line ${idToken.line}: '${idToken.value}' is a reserved word and cannot be used as a procedure name`);
        }
        const name: Identifier = this.makeIdentifier(idToken);
        const parameters: Parameter[] = [];
        let paramsEndColumn: number | undefined;

        if (this.match(TokenType.OperatorLParen)) {
            if (this.peek().type !== TokenType.OperatorRParen) {
                parameters.push(this.parseParameter());

                while (this.match(TokenType.OperatorComma)) {
                    parameters.push(this.parseParameter());
                }
            }
            const rParen = this.consume(TokenType.OperatorRParen, "Expected ')' after procedure parameters");
            paramsEndColumn = rParen.column + 1; // 1-based column immediately after ')'
            this.validateParameterOrder(parameters);
        }

        // Optional Function return type (e.g. 'As Long', 'As Scripting.Dictionary', 'As String()')
        let returnType: string | undefined;
        let returnsArray = false;
        if (this.match(TokenType.KeywordAs)) {
            returnType = this.advance().value;
            if (this.peek().type === TokenType.OperatorDot) {
                this.advance(); // consume '.'
                returnType += '.' + this.advance().value;
            }
            if (this.peek().type === TokenType.OperatorLParen && this.peek(1).type === TokenType.OperatorRParen) {
                this.advance(); // '('
                this.advance(); // ')'
                returnsArray = true;
            }
        }

        // Trailing Static: Sub Foo() Static
        if (this.peek().type === TokenType.KeywordStatic) {
            this.advance(); // consume trailing 'Static'
            isStatic = true;
        }

        this.skipNewlines();
        const body: Statement[] = [];
        const expectedEndStr = isFunction ? 'Function' : (isProperty ? 'Property' : 'Sub');
        while (!this.isAtEndTerminator() && this.peek().type !== TokenType.EOF) {
            const stmt = this.parseStatement();
            if (stmt) body.push(stmt);
            this.skipNewlines();
        }

        if (this.peek().type === TokenType.KeywordEnd) {
            this.advance(); // consume 'End'
            const endToken = this.advance();
            if (endToken.value.toLowerCase() !== expectedEndStr.toLowerCase()) {
                this.throwError(`Parse error: Expected '${expectedEndStr}' after 'End' at line ${endToken.line}`);
            }
        } else if (this.peek().type === TokenType.EOF) {
            this.throwError(`Parse error: Expected 'End ${expectedEndStr}'`);
        }

        return { type: 'ProcedureDeclaration', isFunction, isProperty, propertyType, name, parameters, returnType, returnsArray, body, scope: scope || 'public', isStatic, paramsEndColumn };
    }

    private parseDimStatement(isStatic: boolean = false, keywordConsumed: boolean = false): VariableDeclaration {
        if (!keywordConsumed && !isStatic) this.advance(); // 'Dim' (already consumed for Static or caller)
        const declarations: VariableDeclarator[] = [];

        while (true) {
            let isWithEvents = false;
            if (this.match(TokenType.KeywordWithEvents)) {
                isWithEvents = true;
            }
            const idToken = this.advance();
            if (!this.isIdentifier(idToken) &&
                (idToken.type < TokenType.KeywordBase || idToken.type > TokenType.KeywordAddressOf)) {
                this.throwError(`Parse error at line ${idToken.line}: Expected variable name (Found ${this.tokenDisplay(idToken.value)})`);
            }
            const name: Identifier = this.makeIdentifier(idToken);

            let isArray = false;
            let arrayBounds: ArrayBound[] | undefined;
            let isNew = false;
            let objectType: string | undefined;
            let objectTypeLoc: SourceLocation | undefined;

            let arrayEndColumn: number | undefined;
            if (this.match(TokenType.OperatorLParen)) {
                isArray = true;
                if (this.peek().type !== TokenType.OperatorRParen) {
                    arrayBounds = [];
                    while (true) {
                        let lower: Expression | undefined;
                        let upper = this.parseExpression();
                        if (this.match(TokenType.KeywordTo)) {
                            lower = upper;
                            upper = this.parseExpression();
                        }
                        arrayBounds.push({ lower, upper });
                        if (!this.match(TokenType.OperatorComma)) {
                            break;
                        }
                    }
                }
                const rParenTok = this.peek();
                if (this.match(TokenType.OperatorRParen)) {
                    arrayEndColumn = rParenTok.column + 1; // 1-based column immediately after ')'
                }
            }

            if (this.match(TokenType.KeywordAs)) {
                if (this.match(TokenType.KeywordNew)) {
                    isNew = true;
                }
                const typeToken = this.peek();
                if (this.isNameToken(typeToken)) {
                    objectType = this.advance().value;
                    let lastTypeTok = typeToken;
                    if (this.peek().type === TokenType.OperatorDot) {
                        this.advance(); // consume '.'
                        lastTypeTok = this.peek();
                        objectType += '.' + this.advance().value;
                    }
                    objectTypeLoc = {
                        start: { line: typeToken.line, column: typeToken.column },
                        end: { line: lastTypeTok.line, column: lastTypeTok.column + lastTypeTok.value.length },
                    };
                }
            }

            declarations.push({ name, isArray, arrayBounds, isNew, isWithEvents, objectType, objectTypeLoc, arrayEndColumn });

            if (this.match(TokenType.OperatorComma)) {
                continue;
            } else {
                break;
            }
        }

        return { type: 'VariableDeclaration', declarations, isStatic };
    }

    private parseConstDeclaration(): ConstDeclaration {
        this.advance(); // 'Const'
        const idToken = this.advance();
        if (!this.isIdentifier(idToken)) this.throwError(`Parse error: Expected identifier after Const at line ${idToken.line}`);
        const name = this.makeIdentifier(idToken);

        // Optional 'As Type'
        if (this.match(TokenType.KeywordAs)) {
            this.advance(); // Ignore Type for now
        }

        if (!this.match(TokenType.OperatorEquals)) this.throwError(`Parse error: Expected '=' in Const at line ${this.peek().line}`);
        const value = this.parseExpression();

        return { type: 'ConstDeclaration', name, value };
    }

    private parseSetStatement(): SetStatement {
        this.advance(); // 'Set'
        const left = this.parsePrimary(); // parse identifier or member access
        if (!this.match(TokenType.OperatorEquals)) this.throwError(`Parse error: Expected '=' in Set statement at line ${this.peek().line}`);
        const right = this.parseExpression();
        return { type: 'SetStatement', left, right };
    }

    private parseOnErrorStatement(): OnErrorStatement {
        this.advance(); // 'On'
        if (!this.match(TokenType.KeywordError)) this.throwError(`Parse error: Expected 'Error' after 'On' at line ${this.peek().line}`);

        let label = '';
        if (this.match(TokenType.KeywordGoTo)) {
            const labelToken = this.advance(); // Identifier or 0
            label = labelToken.value;
        } else {
            // "Resume Next" fallback
            while (!this.isAtTerminator()) {
                label += this.advance().value + ' ';
            }
            label = label.trim();
        }
        return { type: 'OnErrorStatement', label };
    }

    private parseExitStatement(): ExitStatement {
        this.advance(); // 'Exit'
        const typeToken = this.advance();
        let exitType: 'For' | 'Do' | 'Sub' | 'Function' | 'Property';

        if (typeToken.type === TokenType.KeywordFor) {
            exitType = 'For';
        } else if (typeToken.type === TokenType.KeywordDo) {
            exitType = 'Do';
        } else if (typeToken.type === TokenType.KeywordSub) {
            exitType = 'Sub';
        } else if (typeToken.type === TokenType.KeywordFunction) {
            exitType = 'Function';
        } else if (typeToken.type === TokenType.KeywordProperty) {
            exitType = 'Property';
        } else {
            this.throwError(`Parse error: Expected 'For', 'Do', 'Sub', 'Function', or 'Property' after 'Exit' at line ${typeToken.line}`);
        }
        return { type: 'ExitStatement', exitType };
    }

    private parseGoToStatement(): GoToStatement {
        this.advance(); // consume 'GoTo'
        const labelToken = this.advance();
        if (!this.isIdentifier(labelToken) && labelToken.type !== TokenType.Number) {
            this.throwError(`Parse error: Expected identifier or number after 'GoTo' at line ${labelToken.line}`);
        }
        return { type: 'GoToStatement', label: labelToken.value };
    }

    private parseResumeStatement(): ResumeStatement {
        this.advance(); // consume 'Resume'
        let target = '';
        // Consume remaining tokens on the line (e.g., 'Next')
        while (this.peek().type !== TokenType.Newline && this.peek().type !== TokenType.EOF) {
            target += this.advance().value;
        }
        return { type: 'ResumeStatement', target: target.trim() } as ResumeStatement;
    }

    private parseEraseStatement(): EraseStatement {
        this.advance(); // 'Erase'
        const idToken = this.advance();
        const name = this.makeIdentifier(idToken);
        return { type: 'EraseStatement', name };
    }

    private parseReDimStatement(): ReDimStatement {
        this.advance(); // 'ReDim'

        let isPreserve = false;
        // Optional 'Preserve' keyword
        if (this.peek().type === TokenType.Identifier && this.peek().value.toLowerCase() === 'preserve') {
            isPreserve = true;
            this.advance();
        }

        const idToken = this.advance();
        const name = this.makeIdentifier(idToken);
        const bounds: ArrayBound[] = [];

        if (this.match(TokenType.OperatorLParen)) {
            if (this.peek().type !== TokenType.OperatorRParen) {
                while (true) {
                    let lower: Expression | undefined;
                    let upper = this.parseExpression();
                    if (this.match(TokenType.KeywordTo)) {
                        lower = upper;
                        upper = this.parseExpression();
                    }
                    bounds.push({ lower, upper });
                    if (!this.match(TokenType.OperatorComma)) {
                        break;
                    }
                }
            }
            this.match(TokenType.OperatorRParen);
        }

        let objectType: string | undefined;
        if (this.match(TokenType.KeywordAs)) {
            objectType = this.advance().value;
        }

        return { type: 'ReDimStatement', name, bounds, isPreserve, objectType };
    }

    private parseTypeDeclaration(): TypeDeclaration {
        this.advance(); // consume 'Type'
        const nameToken = this.advance();
        if (!this.isIdentifier(nameToken)) {
            this.throwError(`Parse error: Expected identifier after 'Type' at line ${nameToken.line}`);
        }
        const typeName = nameToken.value;
        const members: TypeMember[] = [];

        this.skipNewlines();

        // Parse members until 'End Type'
        while (this.peek().type !== TokenType.KeywordEnd && this.peek().type !== TokenType.EOF) {
            // Each member line: memberName [(bounds)] As memberType
            // VBA §5.2.3.3: reserved words are valid member names (reserved-name-member-dcl)
            const memberNameToken = this.advance();
            if (!this.isWordToken(memberNameToken)) {
                this.throwError(`Parse error: Expected member name in Type at line ${memberNameToken.line}`);
            }

            // Skip optional array bounds: (0 To 31), (), etc.
            if (this.peek().type === TokenType.OperatorLParen) {
                this.advance(); // consume '('
                let depth = 1;
                while (depth > 0 && this.peek().type !== TokenType.EOF && this.peek().type !== TokenType.Newline) {
                    const t = this.advance();
                    if (t.type === TokenType.OperatorLParen) depth++;
                    else if (t.type === TokenType.OperatorRParen) depth--;
                }
            }

            if (!this.match(TokenType.KeywordAs)) {
                this.throwError(`Parse error: Expected 'As' in Type member declaration at line ${this.peek().line}`);
            }

            const memberTypeToken = this.advance();
            members.push({ name: memberNameToken.value, memberType: memberTypeToken.value });

            this.skipNewlines();
        }

        // Consume 'End Type'
        if (this.peek().type === TokenType.KeywordEnd) {
            this.advance(); // consume 'End'
            if (!this.match(TokenType.KeywordType)) {
                this.throwError(`Parse error: Expected 'Type' after 'End' at line ${this.peek().line}`);
            }
        }

        return { type: 'TypeDeclaration', name: typeName, members } as TypeDeclaration;
    }

    private parseEnumDeclaration(): EnumDeclaration {
        this.advance(); // consume 'Enum'
        const nameToken = this.advance();
        if (!this.isIdentifier(nameToken)) {
            this.throwError(`Parse error: Expected identifier after 'Enum' at line ${nameToken.line}`);
        }
        const name: Identifier = this.makeIdentifier(nameToken);
        const members: EnumMember[] = [];

        this.skipNewlines();

        while (this.peek().type !== TokenType.KeywordEnd && this.peek().type !== TokenType.EOF) {
            // VBA §5.2.3.3: reserved words are valid member names (reserved-name-member-dcl)
            const memberNameToken = this.advance();
            if (!this.isWordToken(memberNameToken)) {
                this.throwError(`Parse error: Expected member name in Enum at line ${memberNameToken.line}`);
            }
            const memberName: Identifier = this.makeIdentifier(memberNameToken);
            let value: Expression | undefined;

            if (this.match(TokenType.OperatorEquals)) {
                value = this.parseExpression();
            }

            members.push({ name: memberName, value });
            this.skipNewlines();
        }

        if (this.peek().type === TokenType.KeywordEnd) {
            this.advance(); // consume 'End'
            if (!this.match(TokenType.KeywordEnum)) {
                this.throwError(`Parse error: Expected 'Enum' after 'End' at line ${this.peek().line}`);
            }
        }

        return { type: 'EnumDeclaration', name, members } as EnumDeclaration;
    }

    private parseClassDeclaration(): ClassDeclaration {
        this.advance(); // consume 'Class'
        const nameToken = this.advance();
        if (!this.isIdentifier(nameToken)) {
            this.throwError(`Parse error: Expected class name after 'Class' at line ${nameToken.line}`);
        }
        return this.parseClassBody(nameToken.value, true);
    }

    private parseClassBody(className: string, untilEndClass: boolean): ClassDeclaration {
        const fields: VariableDeclaration[] = [];
        const procedures: ProcedureDeclaration[] = [];
        const body: Statement[] = [];

        this.skipNewlines();

        let clsStartTok: Token | undefined;
        let clsEndTok: Token | undefined;

        while (this.peek().type !== TokenType.EOF) {
            if (untilEndClass && this.peek().type === TokenType.KeywordEnd && this.peek(1).type === TokenType.KeywordClass) {
                break;
            }
            const tok = this.peek();

            // Skip blank lines
            if (tok.type === TokenType.Newline) {
                this.skipNewlines();
                continue;
            }

            if (!clsStartTok && tok.line !== undefined) clsStartTok = tok;

            // Scope modifiers before fields/procedures
            let scope: 'public' | 'private' | 'friend' | undefined;
            if (tok.type === TokenType.KeywordPublic || tok.type === TokenType.KeywordPrivate || tok.type === TokenType.KeywordFriend) {
                scope = tok.value.toLowerCase() as 'public' | 'private' | 'friend';
                this.advance(); // consume scope keyword
            }

            const inner = this.peek();
            if (inner.type === TokenType.KeywordSub || inner.type === TokenType.KeywordFunction || inner.type === TokenType.KeywordProperty) {
                const proc = this.parseProcedureDeclaration(scope, undefined, true);
                if (tok.line !== undefined) {
                    const endTok = this.tokens[this.pos - 1];
                    proc.loc = { start: { line: tok.line, column: tok.column }, end: { line: endTok.line, column: endTok.column + endTok.value.length } };
                    clsEndTok = this.tokens[this.pos - 1];
                }
                proc.moduleName = className;
                procedures.push(proc);
                body.push(proc);
            } else if (inner.type === TokenType.KeywordDim) {
                this.advance(); // consume 'Dim'
                const field = this.parseDimStatement(false, true);
                if (tok.line !== undefined) {
                    const endTok = this.tokens[this.pos - 1];
                    field.loc = { start: { line: tok.line, column: tok.column }, end: { line: endTok.line, column: endTok.column + endTok.value.length } };
                    clsEndTok = this.tokens[this.pos - 1];
                }
                field.scope = scope ?? 'public';
                fields.push(field);
                body.push(field);
            } else if (inner.type === TokenType.KeywordImplements) {
                const impl = this.parseImplementsDirective();
                if (tok.line !== undefined) {
                    const endTok = this.tokens[this.pos - 1];
                    impl.loc = { start: { line: tok.line, column: tok.column }, end: { line: endTok.line, column: endTok.column + endTok.value.length } };
                    clsEndTok = this.tokens[this.pos - 1];
                }
                body.push(impl);
            } else if (inner.type === TokenType.KeywordEvent) {
                const event = this.parseEventDeclaration(scope);
                if (tok.line !== undefined) {
                    const endTok = this.tokens[this.pos - 1];
                    event.loc = { start: { line: tok.line, column: tok.column }, end: { line: endTok.line, column: endTok.column + endTok.value.length } };
                    clsEndTok = this.tokens[this.pos - 1];
                }
                body.push(event);
            } else if (inner.type === TokenType.KeywordConst) {
                const constDecl = this.parseConstDeclaration();
                if (tok.line !== undefined) {
                    const endTok = this.tokens[this.pos - 1];
                    constDecl.loc = { start: { line: tok.line, column: tok.column }, end: { line: endTok.line, column: endTok.column + endTok.value.length } };
                    clsEndTok = this.tokens[this.pos - 1];
                }
                body.push(constDecl);
            } else if (inner.type === TokenType.KeywordStatic) {
                this.advance(); // consume 'Static'
                const field = this.parseDimStatement(true, true);
                if (tok.line !== undefined) {
                    const endTok = this.tokens[this.pos - 1];
                    field.loc = { start: { line: tok.line, column: tok.column }, end: { line: endTok.line, column: endTok.column + endTok.value.length } };
                    clsEndTok = this.tokens[this.pos - 1];
                }
                field.scope = scope ?? 'public';
                fields.push(field);
                body.push(field);
            } else if (scope !== undefined && (this.isIdentifier(inner) || inner.type === TokenType.KeywordWithEvents)) {
                // Public/Private [WithEvents] Name As Type (no Dim keyword)
                const field = this.parseDimStatement(false, true);
                if (tok.line !== undefined) {
                    const endTok = this.tokens[this.pos - 1];
                    field.loc = { start: { line: tok.line, column: tok.column }, end: { line: endTok.line, column: endTok.column + endTok.value.length } };
                    clsEndTok = this.tokens[this.pos - 1];
                }
                field.scope = scope;
                fields.push(field);
                body.push(field);
            } else {
                // Skip unknown tokens gracefully
                this.advance();
            }
            this.skipNewlines();
        }

        if (untilEndClass && this.peek().type === TokenType.KeywordEnd) {
            this.advance();
            if (!this.match(TokenType.KeywordClass)) {
                this.throwError(`Parse error: Expected 'Class' after 'End' at line ${this.peek().line}`);
            }
            clsEndTok = this.tokens[this.pos - 1];
        }

        const clsNode = { type: 'ClassDeclaration', name: className, fields, procedures, body } as ClassDeclaration;
        if (clsStartTok && clsEndTok) {
            clsNode.loc = {
                start: { line: clsStartTok.line, column: clsStartTok.column },
                end: { line: clsEndTok.line, column: clsEndTok.column + clsEndTok.value.length },
            };
        }
        return clsNode;
    }

    private parseForStatement(): ForStatement | ForEachStatement {
        this.advance(); // consume 'For'

        if (this.peek().type === TokenType.KeywordEach) {
            return this.parseForEachStatementBody();
        }

        const idToken = this.advance();
        if (!this.isIdentifier(idToken)) {
            this.throwError(`Parse error: Expected identifier after 'For' at line ${idToken.line} `);
        }
        const identifier: Identifier = this.makeIdentifier(idToken);

        if (!this.match(TokenType.OperatorEquals)) {
            this.throwError(`Parse error: Expected '=' in For statement at line ${this.peek().line} `);
        }

        const startExpr = this.parseExpression();

        if (!this.match(TokenType.KeywordTo)) {
            this.throwError(`Parse error: Expected 'To' in For statement at line ${this.peek().line} `);
        }

        const endExpr = this.parseExpression();

        let stepExpr: Expression | undefined;
        if (this.match(TokenType.KeywordStep)) {
            stepExpr = this.parseExpression();
        }

        this.skipNewlines();

        const body: Statement[] = [];
        while (this.peek().type !== TokenType.KeywordNext && this.peek().type !== TokenType.EOF && !this.isAtEndTerminator()) {
            const stmt = this.parseStatement();
            if (stmt) body.push(stmt);
            this.skipNewlines();
        }

        if (!this.match(TokenType.KeywordNext)) {
            this.throwError(`Parse error: Expected 'Next' at line ${this.peek().line} `);
        }

        let nextIdentifier: Identifier | undefined;
        if (this.isIdentifier(this.peek())) {
            const nextIdToken = this.advance();
            nextIdentifier = this.makeIdentifier(nextIdToken);
            if (nextIdentifier.name.toLowerCase() !== identifier.name.toLowerCase()) {
                this.throwError(
                    `Compile error: Variable reference not valid in 'Next' (expected '${identifier.name}', got '${nextIdentifier.name}')`,
                    nextIdToken
                );
            }
        }

        return {
            type: 'ForStatement',
            identifier,
            start: startExpr,
            end: endExpr,
            step: stepExpr,
            body,
            nextIdentifier
        };
    }



    private parseForEachStatementBody(): ForEachStatement {
        this.advance(); // consume 'Each'

        const varToken = this.advance();
        if (!this.isIdentifier(varToken)) {
            this.throwError(`Parse error: Expected identifier after 'For Each' at line ${varToken.line}`);
        }
        const variable: Identifier = this.makeIdentifier(varToken);

        if (!this.match(TokenType.KeywordIn)) {
            this.throwError(`Parse error: Expected 'In' in For Each statement at line ${this.peek().line}`);
        }

        const collection = this.parseExpression();

        this.skipNewlines();

        const body: Statement[] = [];
        while (this.peek().type !== TokenType.KeywordNext && this.peek().type !== TokenType.EOF && !this.isAtEndTerminator()) {
            const stmt = this.parseStatement();
            if (stmt) body.push(stmt);
            this.skipNewlines();
        }

        if (!this.match(TokenType.KeywordNext)) {
            this.throwError(`Parse error: Expected 'Next' at line ${this.peek().line}`);
        }

        let nextIdentifier: Identifier | undefined;
        if (this.isIdentifier(this.peek())) {
            const nextIdToken = this.advance();
            nextIdentifier = this.makeIdentifier(nextIdToken);
            if (nextIdentifier.name.toLowerCase() !== variable.name.toLowerCase()) {
                this.throwError(
                    `Compile error: Variable reference not valid in 'Next' (expected '${variable.name}', got '${nextIdentifier.name}')`,
                    nextIdToken
                );
            }
        }

        return {
            type: 'ForEachStatement',
            variable,
            collection,
            body,
            nextIdentifier
        };
    }

    private parseIfStatement(): IfStatement {
        this.advance(); // Consume 'If' or 'ElseIf'
        const condition = this.parseExpression();

        if (!this.match(TokenType.KeywordThen)) {
            this.throwError(`Parse error: Expected 'Then' after condition at line ${this.peek().line}`);
        }

        const isMultiLine = this.peek().type === TokenType.Newline;

        const consequent: Statement[] = [];
        let alternate: Statement[] | IfStatement | null = null;

        if (!isMultiLine) {
            // Single-line If: read consequent until Else/Newline/EOF.
            // EOS check disabled: Else terminates the body but is not an EOS token.
            while (
                this.peek().type !== TokenType.Newline &&
                this.peek().type !== TokenType.EOF &&
                this.peek().type !== TokenType.KeywordElse
            ) {
                const stmt = this.parseStatement(false);
                if (stmt) consequent.push(stmt);
            }
            // Handle optional inline Else
            if (this.peek().type === TokenType.KeywordElse) {
                this.advance(); // consume 'Else'
                alternate = [];
                while (this.peek().type !== TokenType.Newline && this.peek().type !== TokenType.EOF) {
                    const stmt = this.parseStatement(false);
                    if (stmt) (alternate as Statement[]).push(stmt);
                }
            }
            return {
                type: 'IfStatement',
                condition,
                consequent,
                alternate
            };
        }

        this.skipNewlines();

        while (
            !this.isAtEndTerminator() &&
            this.peek().type !== TokenType.KeywordElse &&
            this.peek().type !== TokenType.KeywordElseIf &&
            this.peek().type !== TokenType.EOF
        ) {
            const stmt = this.parseStatement();
            if (stmt) consequent.push(stmt);
            this.skipNewlines();
        }

        if (this.peek().type === TokenType.KeywordElseIf) {
            alternate = this.parseIfStatement(); // Recursive ElseIf
        } else if (this.match(TokenType.KeywordElse)) {
            this.skipNewlines();
            alternate = [];
            while (
                !this.isAtEndTerminator() &&
                this.peek().type !== TokenType.EOF
            ) {
                const stmt = this.parseStatement();
                if (stmt) alternate.push(stmt);
                this.skipNewlines();
            }
        }

        // Only end the top-level IF chain with End If
        if (this.peek().type === TokenType.KeywordEnd) {
            this.advance(); // Consume 'End'
            if (!this.match(TokenType.KeywordIf)) {
                this.throwError(`Parse error: Expected 'If' after 'End' at line ${this.peek().line}`);
            }
        }

        return {
            type: 'IfStatement',
            condition,
            consequent,
            alternate
        };
    }

    private parseDoWhileStatement(): DoWhileStatement {
        this.advance(); // consume 'Do'

        // Check for pre-condition: Do While/Until condition
        let conditionType: 'while' | 'until' | undefined;
        let conditionPosition: 'pre' | 'post' | undefined;
        let condition: Expression | undefined;

        if (this.peek().type === TokenType.KeywordWhile) {
            this.advance(); // consume 'While'
            conditionType = 'while';
            conditionPosition = 'pre';
            condition = this.parseExpression();
        } else if (this.peek().type === TokenType.KeywordUntil) {
            this.advance(); // consume 'Until'
            conditionType = 'until';
            conditionPosition = 'pre';
            condition = this.parseExpression();
        }
        // else: no pre-condition (infinite or post-condition)

        this.skipNewlines();

        const body: Statement[] = [];
        while (this.peek().type !== TokenType.KeywordLoop && this.peek().type !== TokenType.EOF && !this.isAtEndTerminator()) {
            const stmt = this.parseStatement();
            if (stmt) body.push(stmt);
            this.skipNewlines();
        }

        if (!this.match(TokenType.KeywordLoop)) {
            this.throwError(`Parse error: Expected 'Loop' at line ${this.peek().line} `);
        }

        // Check for post-condition: Loop While/Until condition
        if (this.peek().type === TokenType.KeywordWhile) {
            this.advance(); // consume 'While'
            conditionType = 'while';
            conditionPosition = 'post';
            condition = this.parseExpression();
        } else if (this.peek().type === TokenType.KeywordUntil) {
            this.advance(); // consume 'Until'
            conditionType = 'until';
            conditionPosition = 'post';
            condition = this.parseExpression();
        }

        return {
            type: 'DoWhileStatement',
            conditionType,
            conditionPosition,
            condition,
            body
        };
    }

    private parseWhileStatement(): WhileStatement {
        this.advance(); // consume 'While'
        const condition = this.parseExpression();
        this.skipNewlines();

        const body: Statement[] = [];
        while (this.peek().type !== TokenType.KeywordWend && this.peek().type !== TokenType.EOF && !this.isAtEndTerminator()) {
            const stmt = this.parseStatement();
            if (stmt) body.push(stmt);
            this.skipNewlines();
        }

        if (!this.match(TokenType.KeywordWend)) {
            this.throwError(`Parse error: Expected 'Wend' at line ${this.peek().line}`);
        }

        return { type: 'WhileStatement', condition, body };
    }

    private parseSelectCaseStatement(): SelectCaseStatement {
        this.advance(); // consume 'Select'
        if (!this.match(TokenType.KeywordCase)) {
            this.throwError(`Parse error: Expected 'Case' after 'Select' at line ${this.peek().line}`);
        }
        const expression = this.parseExpression();
        this.skipNewlines();

        const cases: CaseClause[] = [];
        let elseBody: Statement[] | null = null;

        while (!this.isAtEndTerminator() && this.peek().type !== TokenType.EOF) {
            if (this.peek().type !== TokenType.KeywordCase) {
                this.throwError(`Parse error: Expected 'Case' in Select Case at line ${this.peek().line}`);
            }
            this.advance(); // consume 'Case'

            // Case Else
            if (this.peek().type === TokenType.KeywordElse) {
                this.advance(); // consume 'Else'
                this.skipNewlines();
                elseBody = [];
                while (
                    !this.isAtEndTerminator() &&
                    this.peek().type !== TokenType.KeywordCase &&
                    this.peek().type !== TokenType.EOF
                ) {
                    const stmt = this.parseStatement();
                    if (stmt) elseBody.push(stmt);
                    this.skipNewlines();
                }
                continue;
            }

            // Parse range-clauses (comma-separated)
            const ranges: RangeClause[] = [];
            ranges.push(this.parseRangeClause());
            while (this.match(TokenType.OperatorComma)) {
                ranges.push(this.parseRangeClause());
            }
            this.skipNewlines();

            const body: Statement[] = [];
            while (
                !this.isAtEndTerminator() &&
                this.peek().type !== TokenType.KeywordCase &&
                this.peek().type !== TokenType.EOF
            ) {
                const stmt = this.parseStatement();
                if (stmt) body.push(stmt);
                this.skipNewlines();
            }
            cases.push({ ranges, body });
        }

        // consume 'End Select'
        if (this.peek().type === TokenType.KeywordEnd) {
            this.advance(); // 'End'
            if (!this.match(TokenType.KeywordSelect)) {
                this.throwError(`Parse error: Expected 'Select' after 'End' at line ${this.peek().line}`);
            }
        }

        return { type: 'SelectCaseStatement', expression, cases, elseBody };
    }

    private parseWithStatement(): WithStatement {
        this.advance(); // consume 'With'
        const expression = this.parseExpression();
        this.skipNewlines();

        const body: Statement[] = [];
        while (!this.isAtEndTerminator() && this.peek().type !== TokenType.EOF) {
            const stmt = this.parseStatement();
            if (stmt) body.push(stmt);
            this.skipNewlines();
        }

        if (this.peek().type === TokenType.KeywordEnd) {
            this.advance(); // consume 'End'
            if (!this.match(TokenType.KeywordWith)) {
                this.throwError(`Parse error: Expected 'With' after 'End' at line ${this.peek().line}`);
            }
        }

        return { type: 'WithStatement', expression, body };
    }

    private parseRangeClause(): RangeClause {
        // [Is] comparison-operator expression
        const isKeyword = this.peek().type === TokenType.KeywordIs;
        if (isKeyword) this.advance(); // consume 'Is'

        const compOp = this.peek();
        if (
            compOp.type === TokenType.OperatorEquals ||
            compOp.type === TokenType.OperatorNotEquals ||
            compOp.type === TokenType.OperatorLessThan ||
            compOp.type === TokenType.OperatorGreaterThan ||
            compOp.type === TokenType.OperatorLessThanOrEqual ||
            compOp.type === TokenType.OperatorGreaterThanOrEqual
        ) {
            this.advance(); // consume operator
            const value = this.parseExpression();
            return { kind: 'comparison', operator: compOp.value, value };
        }

        if (isKeyword) {
            this.throwError(`Parse error: Expected comparison operator after 'Is' at line ${this.peek().line}`);
        }

        // expression or start-value To end-value
        const startExpr = this.parseExpression();
        if (this.match(TokenType.KeywordTo)) {
            const endExpr = this.parseExpression();
            return { kind: 'to', start: startExpr, end: endExpr };
        }
        return { kind: 'expression', value: startExpr };
    }

    private parseExpression(): Expression {
        return this.parseLogicalImp();
    }

    private parseLogicalImp(): Expression {
        let left = this.parseLogicalEqv();
        while (this.peek().type === TokenType.KeywordImp) {
            const operator = this.advance().value;
            const right = this.parseLogicalEqv();
            left = { type: 'BinaryExpression', operator, left, right, loc: this.makeBinaryLoc(left, right) } as BinaryExpression;
        }
        return left;
    }

    private parseLogicalEqv(): Expression {
        let left = this.parseLogicalXor();
        while (this.peek().type === TokenType.KeywordEqv) {
            const operator = this.advance().value;
            const right = this.parseLogicalXor();
            left = { type: 'BinaryExpression', operator, left, right, loc: this.makeBinaryLoc(left, right) } as BinaryExpression;
        }
        return left;
    }

    private parseLogicalXor(): Expression {
        let left = this.parseLogicalOr();
        while (this.peek().type === TokenType.KeywordXor) {
            const operator = this.advance().value;
            const right = this.parseLogicalOr();
            left = { type: 'BinaryExpression', operator, left, right, loc: this.makeBinaryLoc(left, right) } as BinaryExpression;
        }
        return left;
    }

    private parseLogicalOr(): Expression {
        let left = this.parseLogicalAnd();
        while (this.peek().type === TokenType.KeywordOr) {
            const operator = this.advance().value;
            const right = this.parseLogicalAnd();
            left = { type: 'BinaryExpression', operator, left, right, loc: this.makeBinaryLoc(left, right) } as BinaryExpression;
        }
        return left;
    }

    private parseLogicalAnd(): Expression {
        let left = this.parseLogicalNot();
        while (this.peek().type === TokenType.KeywordAnd) {
            const operator = this.advance().value;
            const right = this.parseLogicalNot();
            left = { type: 'BinaryExpression', operator, left, right, loc: this.makeBinaryLoc(left, right) } as BinaryExpression;
        }
        return left;
    }

    private parseLogicalNot(): Expression {
        if (this.peek().type === TokenType.KeywordNot) {
            const notTok = this.tokens[this.pos];
            this.advance();
            const argument = this.parseEquality();
            const node: UnaryExpression = { type: 'UnaryExpression', operator: 'Not', argument };
            node.loc = { start: { line: notTok.line, column: notTok.column }, end: argument.loc?.end ?? { line: notTok.line, column: notTok.column } };
            return node;
        }
        return this.parseEquality();
    }

    private makeBinaryLoc(left: Expression, right: Expression): SourceLocation | undefined {
        if (!left.loc || !right.loc) return undefined;
        return { start: left.loc.start, end: right.loc.end };
    }

    private parseEquality(): Expression {
        let left = this.parseRelational();
        while (
            this.peek().type === TokenType.OperatorEquals ||
            this.peek().type === TokenType.OperatorNotEquals ||
            this.peek().type === TokenType.KeywordIs ||
            this.peek().type === TokenType.KeywordLike
        ) {
            const operator = this.advance().value;
            const right = this.parseRelational();
            left = { type: 'BinaryExpression', operator, left, right, loc: this.makeBinaryLoc(left, right) } as BinaryExpression;
        }
        return left;
    }

    private parseRelational(): Expression {
        let left = this.parseConcatenation();
        while (
            this.peek().type === TokenType.OperatorLessThan ||
            this.peek().type === TokenType.OperatorGreaterThan ||
            this.peek().type === TokenType.OperatorLessThanOrEqual ||
            this.peek().type === TokenType.OperatorGreaterThanOrEqual
        ) {
            const operator = this.advance().value;
            const right = this.parseConcatenation();
            left = { type: 'BinaryExpression', operator, left, right, loc: this.makeBinaryLoc(left, right) } as BinaryExpression;
        }
        return left;
    }

    private parseConcatenation(): Expression {
        let left = this.parseAdditive();
        while (this.peek().type === TokenType.OperatorAmpersand) {
            const operator = this.advance().value;
            const right = this.parseAdditive();
            left = { type: 'BinaryExpression', operator, left, right, loc: this.makeBinaryLoc(left, right) } as BinaryExpression;
        }
        return left;
    }

    private parseAdditive(): Expression {
        let left = this.parseModulo();
        while (
            this.peek().type === TokenType.OperatorPlus ||
            this.peek().type === TokenType.OperatorMinus
        ) {
            const operator = this.advance().value;
            const right = this.parseModulo();
            left = { type: 'BinaryExpression', operator, left, right, loc: this.makeBinaryLoc(left, right) } as BinaryExpression;
        }
        return left;
    }

    private parseModulo(): Expression {
        let left = this.parseIntDivision();
        while (this.peek().type === TokenType.KeywordMod) {
            const operator = this.advance().value;
            const right = this.parseIntDivision();
            left = { type: 'BinaryExpression', operator, left, right, loc: this.makeBinaryLoc(left, right) } as BinaryExpression;
        }
        return left;
    }

    private parseIntDivision(): Expression {
        let left = this.parseMultiplicative();
        while (this.peek().type === TokenType.OperatorIntDivide) {
            const operator = this.advance().value;
            const right = this.parseMultiplicative();
            left = { type: 'BinaryExpression', operator, left, right, loc: this.makeBinaryLoc(left, right) } as BinaryExpression;
        }
        return left;
    }

    private parseMultiplicative(): Expression {
        let left = this.parseUnary();
        while (
            this.peek().type === TokenType.OperatorMultiply ||
            this.peek().type === TokenType.OperatorDivide
        ) {
            const operator = this.advance().value;
            const right = this.parseUnary();
            left = { type: 'BinaryExpression', operator, left, right, loc: this.makeBinaryLoc(left, right) } as BinaryExpression;
        }
        return left;
    }

    private parseUnary(): Expression {
        if (this.peek().type === TokenType.OperatorMinus || this.peek().type === TokenType.OperatorPlus) {
            const opTok = this.tokens[this.pos];
            const operator = this.advance().value;
            const argument = this.parseUnary();
            const node: UnaryExpression = { type: 'UnaryExpression', operator, argument };
            node.loc = { start: { line: opTok.line, column: opTok.column }, end: argument.loc?.end ?? { line: opTok.line, column: opTok.column } };
            return node;
        }
        return this.parseExponentiation();
    }

    private parseExponentiation(): Expression {
        let left = this.parsePrimary();
        while (this.peek().type === TokenType.OperatorPower) {
            const operator = this.advance().value;
            const right = this.parsePrimary();
            left = { type: 'BinaryExpression', operator, left, right, loc: this.makeBinaryLoc(left, right) } as BinaryExpression;
        }
        return left;
    }

    private parsePrimary(stopBeforeSpacedLParen: boolean = false): Expression {
        const startTok = this.tokens[this.pos];
        const token = this.advance();
        let expr: Expression;
        if (token.type === TokenType.Number) {
            const m = token.value.match(/[%&@!#^]$/);
            const typeSuffix = m ? m[0] as NumberLiteral['typeSuffix'] : undefined;
            const cleanVal = token.value.replace(/[%&@!#^]$/, '');
            // 0x/0o は整数なので isFloat = false。それ以外で . or e/E を含む場合は Double リテラル
            const isFloat = !cleanVal.startsWith('0x') && !cleanVal.startsWith('0o')
                && /[.eE]/.test(cleanVal) ? true as const : undefined;
            // Use Number() to support 0x (Hex) and 0o (Octal) prefixes
            expr = { type: 'NumberLiteral', value: Number(cleanVal), typeSuffix, isFloat } as NumberLiteral;
        } else if (token.type === TokenType.String) {
            expr = { type: 'StringLiteral', value: token.value } as StringLiteral;
        } else if (token.type === TokenType.Date) {
            expr = { type: 'DateLiteral', value: token.value } as DateLiteral;
        } else if (token.type === TokenType.ForeignName) {
            expr = { type: 'Identifier', name: token.value, foreign: true } as Identifier;
        } else if (token.type === TokenType.Identifier ||
                   Parser.CONTEXTUAL_KW.has(token.type) ||
                   Parser.COMPAT_KW_EXPR.has(token.type)) {
            expr = { type: 'Identifier', name: token.value } as Identifier;
        } else if (token.type === TokenType.KeywordAddressOf) {
            const firstTok = this.advance();
            if (!this.isIdentifier(firstTok)) this.throwError(`Parse error at line ${firstTok.line}: Expected procedure name after 'AddressOf'`);
            let moduleName: string | undefined;
            let procTok = firstTok;
            if (this.peek().type === TokenType.OperatorDot) {
                this.advance(); // consume '.'
                moduleName = firstTok.value;
                procTok = this.advance();
                if (!this.isIdentifier(procTok)) this.throwError(`Parse error at line ${procTok.line}: Expected procedure name after 'AddressOf ${moduleName}.'`);
            }
            expr = { type: 'AddressOfExpression', procedureName: { type: 'Identifier', name: procTok.value }, moduleName } as AddressOfExpression;
        } else if (token.type === TokenType.KeywordEmpty) {
            expr = { type: 'Identifier', name: token.value } as Identifier;
        } else if (token.type === TokenType.KeywordNothing) {
            expr = { type: 'Identifier', name: 'Nothing' } as Identifier;
        } else if (token.type === TokenType.KeywordNull) {
            expr = { type: 'Identifier', name: 'Null' } as Identifier;
        } else if (token.type === TokenType.KeywordMe) {
            expr = { type: 'Identifier', name: 'Me' } as Identifier;
        } else if (token.type === TokenType.KeywordError) {
            expr = { type: 'Identifier', name: 'Error' } as Identifier;
        } else if (token.type === TokenType.KeywordNew) {
            const classNameToken = this.advance();
            if (!this.isNameToken(classNameToken)) {
                this.throwError(`Parse error: Expected class name after 'New' at line ${classNameToken.line}`);
            }
            let className = classNameToken.value;
            if (this.peek().type === TokenType.OperatorDot) {
                this.advance(); // consume '.'
                className += '.' + this.advance().value;
            }
            expr = { type: 'NewExpression', className } as NewExpression;
        } else if (token.type === TokenType.OperatorLParen) {
            const innerExpr = this.parseExpression();
            if (!this.match(TokenType.OperatorRParen)) {
                this.throwMissingRParen();
            }
            expr = { type: 'ParenthesizedExpression', expression: innerExpr } as any; // Type added implicitly or via cast
        } else if (token.type === TokenType.KeywordTypeOf) {
            expr = this.parseRelational(); // Stop before 'Is'
            if (!this.match(TokenType.KeywordIs)) {
                this.throwError(`Parse error: Expected 'Is' after 'TypeOf' at line ${this.peek().line}`);
            }
            const typeToken = this.advance();
            if (!this.isIdentifier(typeToken) && typeToken.type !== TokenType.KeywordCollection) {
                 this.throwError(`Parse error: Expected type name after 'Is' at line ${typeToken.line}`);
            }
            expr = { type: 'TypeOfIsExpression', expression: expr, typeName: typeToken.value } as TypeOfIsExpression;
        } else if (token.type === TokenType.OperatorDot) {
            const propToken = this.advance();
            if (!this.isNameToken(propToken)) {
                this.throwError(`Parse error: Expected identifier after '.' at line ${propToken.line}`);
            }
            const property = { type: 'Identifier', name: propToken.value } as Identifier;
            expr = { type: 'ImplicitWithObjectExpression', property } as ImplicitWithObjectExpression;
        } else if (token.type === TokenType.Newline) {
            // Check if the token before the newline suggests a missing line continuation
            const prevToken = this.tokens[Math.max(0, this.pos - 2)];
            if (prevToken && this.isContinuationEndToken(prevToken.type)) {
                this.throwError(
                    `行継続文字 '_' が必要です（'${prevToken.value}' の後で改行されています）`,
                    token
                );
            } else {
                this.throwError(`Parse error: Unexpected token in expression '${this.tokenDisplay(token.value)}' at line ${token.line} `, token);
            }
        } else {
            this.throwError(`Parse error: Unexpected token in expression '${this.tokenDisplay(token.value)}' at line ${token.line} `, token);
        }

        expr.loc = this.exprLoc(startTok, this.tokens[this.pos - 1]);

        while (true) {
            if (this.match(TokenType.OperatorDot)) {
                if (this.peek().type === TokenType.EOF) this.throwError("Expected property name after '.'");
                const propToken = this.advance();
                const property = { type: 'Identifier', name: propToken.value } as Identifier;
                expr = { type: 'MemberExpression', object: expr, property } as MemberExpression;
                expr.loc = this.exprLoc(startTok, this.tokens[this.pos - 1]);
            } else if (this.match(TokenType.OperatorExclamation)) {
                if (this.peek().type === TokenType.EOF) this.throwError("Expected identifier after '!'");
                const propToken = this.advance();
                const property = { type: 'Identifier', name: propToken.value } as Identifier;
                expr = { type: 'DictionaryAccessExpression', object: expr, property } as DictionaryAccessExpression;
                expr.loc = this.exprLoc(startTok, this.tokens[this.pos - 1]);
            } else if (this.peek().type === TokenType.OperatorLParen) {
                if (stopBeforeSpacedLParen && this.hasSpaceBeforeCurrentToken()) {
                    break;
                }
                this.advance(); // consume '('
                const args: Expression[] = [];
                if (this.peek().type !== TokenType.OperatorRParen) {
                    args.push(this.parseCallArgument());
                    while (this.match(TokenType.OperatorComma)) {
                        args.push(this.parseCallArgument());
                    }
                }
                if (!this.match(TokenType.OperatorRParen)) {
                    this.throwMissingRParen();
                }
                expr = { type: 'CallExpression', callee: expr, args } as CallExpression;
                expr.loc = this.exprLoc(startTok, this.tokens[this.pos - 1]);
            } else {
                break;
            }
        }
        return expr;
    }

    // Parse a call argument, handling named arguments (e.g., shift:=xlUp)
    private parseCallArgument(): Expression {
        const next = this.peek().type;
        if (next === TokenType.OperatorComma || next === TokenType.OperatorRParen || this.isAtTerminator()) {
            return { type: 'MissingArgument' } as any;
        }

        // Check for named argument: name := Expression
        // MS-VBAL §5.6.13.1: named-argument ::= unrestricted-name ':=' expression
        // unrestricted-name includes keywords. User-defined procedures cannot declare
        // keyword-named parameters, but COM/built-in methods (e.g. Validation.Add) can have
        // parameters named Type, Date, Name, etc.
        // Restrict to tokens whose value is an identifier-like string to avoid treating
        // number literals or other tokens as names (e.g. "1:=x" must not match).
        if (this.pos + 1 < this.tokens.length &&
            this.tokens[this.pos + 1].type === TokenType.OperatorColonEquals &&
            typeof this.peek().value === 'string' &&
            /^[A-Za-z_]\w*$/.test(this.peek().value as string)) {
            const nameToken = this.advance(); // consume the name token (identifier or keyword)
            this.advance(); // consume ':='
            const value = this.parseExpression();
            return { type: 'NamedArgument', name: nameToken.value, value } as NamedArgument;
        }
        return this.parseExpression();
    }

    private parseOnGoToSubStatement(): OnGoToSubStatement {
        this.advance(); // 'On'
        const expression = this.parseExpression();
        
        let isGoSub = false;
        if (this.match(TokenType.KeywordGoTo)) {
            isGoSub = false;
        } else if (this.match(TokenType.KeywordGoSub)) {
            isGoSub = true;
        } else {
            this.throwError(`Parse error: Expected 'GoTo' or 'GoSub' after 'On' expression at line ${this.peek().line}`);
        }

        const labels: string[] = [];
        while (true) {
            const labelToken = this.advance();
            if (!this.isIdentifier(labelToken) && labelToken.type !== TokenType.Number) {
                this.throwError(`Parse error: Expected label (identifier or number) in On...GoTo/GoSub at line ${labelToken.line}`);
            }
            labels.push(labelToken.value);

            if (!this.match(TokenType.OperatorComma)) {
                break;
            }
        }

        return { type: 'OnGoToSubStatement', expression, isGoSub, labels };
    }

    private parseGoSubStatement(): GoSubStatement {
        this.advance(); // 'GoSub'
        const labelToken = this.advance();
        if (!this.isIdentifier(labelToken) && labelToken.type !== TokenType.Number) {
            this.throwError(`Parse error: Expected label after GoSub at line ${labelToken.line}`);
        }
        return { type: 'GoSubStatement', label: labelToken.value };
    }

    private parseLSetStatement(): LSetStatement {
        this.advance(); // 'LSet'
        const left = this.parsePrimary();
        if (!this.match(TokenType.OperatorEquals)) {
            this.throwError(`Parse error: Expected '=' in LSet statement at line ${this.peek().line}`);
        }
        const right = this.parseExpression();
        return { type: 'LSetStatement', left, right };
    }

    private parseRSetStatement(): RSetStatement {
        this.advance(); // 'RSet'
        const left = this.parsePrimary();
        if (!this.match(TokenType.OperatorEquals)) {
            this.throwError(`Parse error: Expected '=' in RSet statement at line ${this.peek().line}`);
        }
        const right = this.parseExpression();
        return { type: 'RSetStatement', left, right };
    }

    private hasMidAssignmentAhead(): boolean {
        // Mid ( ... ) =
        if (this.peek(1).type !== TokenType.OperatorLParen) return false;
        let depth = 0;
        let i = 1;
        while (this.pos + i < this.tokens.length) {
            const t = this.tokens[this.pos + i];
            if (t.type === TokenType.OperatorLParen) depth++;
            else if (t.type === TokenType.OperatorRParen) {
                depth--;
                if (depth === 0) {
                    return this.tokens[this.pos + i + 1]?.type === TokenType.OperatorEquals;
                }
            } else if (t.type === TokenType.Newline || t.type === TokenType.EOF) break;
            i++;
        }
        return false;
    }

    private parseMidStatement(): MidStatement {
        const midToken = this.advance(); // 'Mid' or 'Mid$' or 'MidB' or 'MidB$'
        const isByte = midToken.value.toLowerCase().startsWith('midb');
        this.consume(TokenType.OperatorLParen, "Expected '(' after Mid");
        const target = this.parseExpression();
        this.consume(TokenType.OperatorComma, "Expected ',' after Mid target");
        const start = this.parseExpression();
        let length: Expression | null = null;
        if (this.match(TokenType.OperatorComma)) {
            length = this.parseExpression();
        }
        this.consume(TokenType.OperatorRParen, "Expected ')' after Mid arguments");
        this.consume(TokenType.OperatorEquals, "Expected '=' in Mid statement");
        const value = this.parseExpression();
        return { type: 'MidStatement', target, start, length, value, isByte };
    }

    private parseErrorStatement(): ErrorStatement {
        this.advance(); // 'Error'
        const errorNumber = this.parseExpression();
        return { type: 'ErrorStatement', errorNumber };
    }

    private parseEventDeclaration(scope?: 'public' | 'private' | 'friend'): EventDeclaration {
        this.advance(); // 'Event'
        const idToken = this.advance();
        if (!this.isIdentifier(idToken)) this.throwError(`Expected identifier after 'Event' at line ${idToken.line}`);
        const name: Identifier = { type: 'Identifier', name: idToken.value };
        const parameters: Parameter[] = [];
        
        if (this.match(TokenType.OperatorLParen)) {
            if (this.peek().type !== TokenType.OperatorRParen) {
                parameters.push(this.parseParameter());
                while (this.match(TokenType.OperatorComma)) {
                    parameters.push(this.parseParameter());
                }
            }
            this.consume(TokenType.OperatorRParen, "Expected ')' after event parameters");
            this.validateParameterOrder(parameters);
        }

        return { type: 'EventDeclaration', name, parameters, scope };
    }

    private parseRaiseEventStatement(): RaiseEventStatement {
        this.advance(); // 'RaiseEvent'
        const idToken = this.advance();
        if (!this.isIdentifier(idToken)) this.throwError(`Expected identifier after 'RaiseEvent' at line ${idToken.line}`);
        const eventName: Identifier = { type: 'Identifier', name: idToken.value };
        const args: Expression[] = [];
        
        if (this.match(TokenType.OperatorLParen)) {
            if (this.peek().type !== TokenType.OperatorRParen) {
                args.push(this.parseExpression());
                while (this.match(TokenType.OperatorComma)) {
                    args.push(this.parseExpression());
                }
            }
            this.consume(TokenType.OperatorRParen, "Expected ')' after RaiseEvent arguments");
        }
        
        return { type: 'RaiseEventStatement', eventName, args };
    }

    private parseImplementsDirective(): ImplementsDirective {
        this.advance(); // 'Implements'
        const idToken = this.advance();
        if (!this.isIdentifier(idToken)) this.throwError(`Parse error at line ${idToken.line}: Expected interface name after 'Implements'`);
        return { type: 'ImplementsDirective', interfaceName: idToken.value };
    }

    private parseAppActivateStatement(): AppActivateStatement {
        this.advance(); // 'AppActivate'
        const title = this.parseExpression();
        let wait: Expression | undefined;
        if (this.match(TokenType.OperatorComma)) {
            wait = this.parseExpression();
        }
        return { type: 'AppActivateStatement', title, wait };
    }

    private parseSendKeysStatement(): SendKeysStatement {
        this.advance(); // 'SendKeys'
        const keys = this.parseExpression();
        let wait: Expression | undefined;
        if (this.match(TokenType.OperatorComma)) {
            wait = this.parseExpression();
        }
        return { type: 'SendKeysStatement', keys, wait };
    }

    private parseLockStatement(): LockStatement {
        this.advance(); // 'Lock'
        this.match(TokenType.OperatorHash);
        const fileNumber = this.parseExpression();
        let recordRange: any;
        if (this.match(TokenType.OperatorComma)) {
            const start = this.parseExpression();
            let end: any;
            if (this.match(TokenType.KeywordTo)) {
                end = this.parseExpression();
            }
            recordRange = { start, end };
        }
        return { type: 'LockStatement', fileNumber, recordRange };
    }

    private parseUnlockStatement(): UnlockStatement {
        this.advance(); // 'Unlock'
        this.match(TokenType.OperatorHash);
        const fileNumber = this.parseExpression();
        let recordRange: any;
        if (this.match(TokenType.OperatorComma)) {
            const start = this.parseExpression();
            let end: any;
            if (this.match(TokenType.KeywordTo)) {
                end = this.parseExpression();
            }
            recordRange = { start, end };
        }
        return { type: 'UnlockStatement', fileNumber, recordRange };
    }

    private parseWidthStatement(): WidthStatement {
        this.advance(); // 'Width'
        this.consume(TokenType.OperatorHash, "Expected '#' after Width");
        const fileNumber = this.parseExpression();
        this.consume(TokenType.OperatorComma, "Expected ',' after file number");
        const width = this.parseExpression();
        return { type: 'WidthStatement', fileNumber, width };
    }
}
