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
    body: Statement[];
    scope?: 'public' | 'private' | 'friend';
    isStatic?: boolean;
    moduleName?: string;
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
}

export interface ImplicitWithObjectExpression extends Expression {
    type: 'ImplicitWithObjectExpression';
    property: Identifier;
}

export interface NumberLiteral extends Expression {
    type: 'NumberLiteral';
    value: number;
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

    /** Union of all contextual keyword groups above.
     *  Tokens in this Set are valid IDENTIFIERs in Dim declarations,
     *  expression context, and assignment statements. */
    private static readonly CONTEXTUAL_KW = new Set<TokenType>([
        ...Parser.CONTEXTUAL_KW_FILE_MODE,
        ...Parser.CONTEXTUAL_KW_FILE_ACCESS,
        ...Parser.CONTEXTUAL_KW_OPTION,
        ...Parser.CONTEXTUAL_KW_DECLARE,
        ...Parser.CONTEXTUAL_KW_STMT_ABSENT,
    ]);

    /** <statement-keyword> tokens additionally permitted as IDENTIFIERs in
     *  expression context (parsePrimary) for practical VBA compatibility —
     *  e.g. as method names (obj.Print, ws.Get) or file I/O targets.
     *  These remain <reserved-identifier> per §3.3.5.2. */
    private static readonly COMPAT_KW_EXPR = new Set<TokenType>([
        TokenType.KeywordSeek,    // <statement-keyword>
        TokenType.KeywordInput,   // <statement-keyword> / <special-form>
        TokenType.KeywordPrint,   // <statement-keyword>
        TokenType.KeywordPut,     // <statement-keyword>
        TokenType.KeywordGet,     // <statement-keyword>
        TokenType.KeywordLock,    // <statement-keyword>
        TokenType.KeywordUnlock,  // <statement-keyword>
    ]);

    private readonly errorRecovery: boolean;

    constructor(tokens: Token[], options: { parseAsClass?: string; errorRecovery?: boolean } = {}) {
        this.tokens = tokens;
        this.parseAsClass = options.parseAsClass;
        this.errorRecovery = options.errorRecovery ?? false;
    }

    // Keywords can appear as property/class names in VBA (e.g. obj.Property, New Collection)
    private isNameToken(token: Token): boolean {
        return token.type === TokenType.Identifier
            || (token.type >= TokenType.KeywordFor && token.type <= TokenType.KeywordAddressOf);
    }

    private recordError(message: string, token: Token): void {
        const pos: Position = { line: token.line, column: token.column };
        this._diagnostics.push({ message, loc: { start: pos, end: pos }, severity: 'error' });
    }

    private syncToNextTopLevelStatement(): void {
        // After a parse error, skip tokens until we find what looks like
        // the start of a new top-level statement. We always advance at least
        // one token to guarantee forward progress, then advance until we hit
        // a Newline. The Newline itself is left for parse()'s skipNewlines().
        // This is coarse but keeps the parser state consistent.
        while (
            this.peek().type !== TokenType.EOF &&
            this.peek().type !== TokenType.Newline
        ) {
            this.advance();
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
            throw new Error(`Parse error: Expected 'Binary' or 'Text' after 'Option Compare' at line ${modeToken.line}`);
        }
        return { type: 'OptionCompareStatement', mode };
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

        if (this.peek().type === TokenType.KeywordByVal || this.peek().type === TokenType.KeywordByRef) {
            isByVal = this.advance().type === TokenType.KeywordByVal;
        }

        const token = this.peek();
        if (token.type !== TokenType.Identifier && (token.type < TokenType.KeywordBase || token.type > TokenType.KeywordAddressOf)) {
            throw new Error(`Parse error at line ${token.line}: Expected parameter name (Found ${token.value})`);
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

        return { type: 'Parameter', name: nameToken.value, isByVal, isOptional, isParamArray, isArray, paramType, defaultValue };
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
            default: throw new Error(`Parse error: Invalid Open mode '${modeToken.value}' at line ${modeToken.line}`);
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
            throw new Error(`Parse error: Expected variable name in Line Input at line ${this.peek().line}`);
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
            throw new Error(`Parser error: Expected Sub or Function after Declare at line ${this.peek().line}`);
        }

        const name = this.advance().value;

        if (this.peek().type !== TokenType.KeywordLib) {
            throw new Error(`Parser error: Expected Lib after Declare name at line ${this.peek().line}`);
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

    private isAtTerminator(): boolean {
        const type = this.peek().type;
        return type === TokenType.Newline || type === TokenType.EOF || type === TokenType.OperatorColon;
    }

    private consume(expectedType: TokenType, message: string): Token {
        if (this.peek().type === expectedType) {
            return this.advance();
        }
        throw new Error(`Parse error at line ${this.peek().line}: ${message}`);
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
    private isWordToken(t: Token): boolean {
        return /^[A-Za-z_][A-Za-z0-9_]*$/.test(t.value);
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
                this.recordError(msg, startToken);
                this.syncToNextTopLevelStatement();
            }
            // Defensive: ensure forward progress even if a parser returns without advancing
            if (this.pos === startPos && this.peek().type !== TokenType.EOF) {
                this.advance();
            }
            this.skipNewlines();
        }

        return program;
    }

    public parseExpressionPublic(): Expression {
        this.skipNewlines();
        return this.parseExpression();
    }

    private parseStatement(): Statement | null {
        this.skipNewlines();
        const startToken = this.peek();
        const stmt = this.parseStatementInner();
        if (stmt !== null) {
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

    private parseStatementInner(): Statement | null {
        const token = this.peek();

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
        } else if (token.type === TokenType.KeywordSub || token.type === TokenType.KeywordFunction || token.type === TokenType.KeywordProperty) {
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
            return { type: 'StopStatement' } as StopStatement;
        } else if (token.type === TokenType.KeywordEnd) {
            if (this.isAtEndTerminator()) {
                return null;
            }
            this.advance(); // consume standalone 'End'
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
        } else if (token.type === TokenType.KeywordError) {
            return this.parseErrorStatement();
        } else if (token.type === TokenType.KeywordGoSub) {
            return this.parseGoSubStatement();
        } else if (token.type === TokenType.KeywordReturn) {
            this.advance(); // consume 'Return'
            return { type: 'ReturnStatement' } as ReturnStatement;
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
                throw new Error(`Parse error: Option Base must be 0 or 1 at line ${baseToken.line}`);
            }
            if (this.match(TokenType.KeywordPrivate)) {
                this.consume(TokenType.KeywordModule, "Expected 'Module' after 'Option Private'");
                return { type: 'OptionPrivateModuleStatement' } as OptionPrivateModuleStatement;
            }
            return null;
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
            return this.parseOpenStatement();
        } else if (token.type === TokenType.KeywordClose) {
            return this.parseCloseStatement();
        } else if (token.type === TokenType.KeywordLine && this.peek(1).type !== TokenType.OperatorEquals) {
            return this.parseLineInputStatement();
        } else if (token.type === TokenType.KeywordPrint) {
            return this.parsePrintStatement();
        } else if (token.type === TokenType.KeywordPut) {
            return this.parsePutStatement();
        } else if (token.type === TokenType.KeywordGet) {
            return this.parseGetStatement();
        } else if (token.type === TokenType.KeywordInput) {
            return this.parseInputStatement();
        } else if (token.type === TokenType.KeywordWrite) {
            return this.parseWriteStatement();
        } else if (token.type === TokenType.KeywordSeek) {
            return this.parseSeekStatement();
        } else if (token.type === TokenType.KeywordReset && this.peek(1).type !== TokenType.OperatorEquals) {
            return this.parseResetStatement();
        } else if (token.type === TokenType.KeywordKill && this.peek(1).type !== TokenType.OperatorEquals) {
            return this.parseKillStatement();
        } else if (token.type === TokenType.KeywordEvent) {
            return this.parseEventDeclaration();
        } else if (token.type === TokenType.KeywordRaiseEvent) {
            return this.parseRaiseEventStatement();
        } else if (token.type === TokenType.KeywordLock) {
            return this.parseLockStatement();
        } else if (token.type === TokenType.KeywordUnlock) {
            return this.parseUnlockStatement();
        } else if (token.type === TokenType.KeywordWidth && this.peek(1).type !== TokenType.OperatorEquals) {
            return this.parseWidthStatement();
        } else if (token.type === TokenType.KeywordClass) {
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
            throw new Error(`Parse error: Expected procedure call after 'Call'`);
        } else if (token.type === TokenType.Identifier || token.type === TokenType.OperatorDot ||
                   token.type === TokenType.Number || Parser.CONTEXTUAL_KW.has(token.type)) {
            // Check if it's a label "Identifier:" or "Number" (line number)
            if (token.type === TokenType.Identifier && this.pos + 1 < this.tokens.length && this.tokens[this.pos + 1].type === TokenType.OperatorColon) {
                const labelName = token.value;
                this.advance(); // consume Identifier
                this.advance(); // consume ':'
                return { type: 'LabelStatement', label: labelName } as any;
            } else if (token.type === TokenType.Number) {
                // Line number label.
                const labelName = token.value;
                this.advance(); // consume Number
                // Optional colon after line number
                this.match(TokenType.OperatorColon);
                return { type: 'LabelStatement', label: labelName } as any;
            }

            // Unify assignment, array access, method call
            const expr = this.parsePrimary(); // will parse `foo`, `foo()`, `foo.bar`, `arr(0)` etc

            if (this.match(TokenType.OperatorEquals)) {
                return {
                    type: 'AssignmentStatement',
                    left: expr,
                    right: this.parseExpression()
                } as AssignmentStatement;
            } else {
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
                    // Call matched via parens e.g. `MainLoop()`
                    return { type: 'CallStatement', expression: expr } as CallStatement;
                } else {
                    // Call matched without parens e.g. `MainLoop`
                    return { type: 'CallStatement', expression: { type: 'CallExpression', callee: expr, args: [] } } as CallStatement;
                }
            }
        } else if (token.type === TokenType.Unknown) {
            throw new Error(`Parse error: Unknown token '${token.value}' at line ${token.line}`);
        } else {
            // Unknown or unexpected top-level token
            this.advance();
        }
        return null;
    }

    private parseProcedureDeclaration(scope?: 'public' | 'private' | 'friend', isStatic?: boolean): ProcedureDeclaration {
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
                throw new Error(`Parse error: Expected 'Get', 'Let', or 'Set' after 'Property' at line ${typeToken.line}`);
            }
        }

        const idToken = this.advance();
        if (idToken.type !== TokenType.Identifier && (idToken.type < TokenType.KeywordBase || idToken.type > TokenType.KeywordAddressOf)) {
            throw new Error(`Parse error at line ${idToken.line}: Expected procedure name (Found ${idToken.value})`);
        }
        const name: Identifier = this.makeIdentifier(idToken);
        const parameters: Parameter[] = [];

        if (this.match(TokenType.OperatorLParen)) {
            if (this.peek().type !== TokenType.OperatorRParen) {
                parameters.push(this.parseParameter());

                while (this.match(TokenType.OperatorComma)) {
                    parameters.push(this.parseParameter());
                }
            }
            this.consume(TokenType.OperatorRParen, "Expected ')' after procedure parameters");
        }

        // Optional Function return type (e.g. 'As Long', 'As Scripting.Dictionary')
        let returnType: string | undefined;
        if (this.match(TokenType.KeywordAs)) {
            returnType = this.advance().value;
            if (this.peek().type === TokenType.OperatorDot) {
                this.advance(); // consume '.'
                returnType += '.' + this.advance().value;
            }
        }

        // Trailing Static: Sub Foo() Static
        if (this.peek().type === TokenType.KeywordStatic) {
            this.advance(); // consume trailing 'Static'
            isStatic = true;
        }

        this.skipNewlines();
        const body: Statement[] = [];
        while (!this.isAtEndTerminator() && this.peek().type !== TokenType.EOF) {
            const stmt = this.parseStatement();
            if (stmt) body.push(stmt);
            this.skipNewlines();
        }

        if (this.peek().type === TokenType.KeywordEnd) {
            this.advance(); // consume 'End'
            let expectedEndStr = isFunction ? 'Function' : 'Sub';
            if (isProperty) expectedEndStr = 'Property';
            const endToken = this.advance();
            if (endToken.value.toLowerCase() !== expectedEndStr.toLowerCase()) {
                throw new Error(`Parse error: Expected '${expectedEndStr}' after 'End' at line ${endToken.line}`);
            }
        }

        return { type: 'ProcedureDeclaration', isFunction, isProperty, propertyType, name, parameters, returnType, body, scope: scope || 'public', isStatic };
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
            // CONTEXTUAL_KW covers keywords below KeywordBase that are valid IDENTIFIERs
            // per §3.3.5.2; the range check catches the rest (KeywordBase..KeywordAddressOf).
            if (idToken.type !== TokenType.Identifier &&
                !Parser.CONTEXTUAL_KW.has(idToken.type) &&
                (idToken.type < TokenType.KeywordBase || idToken.type > TokenType.KeywordAddressOf)) {
                throw new Error(`Parse error at line ${idToken.line}: Expected variable name (Found ${idToken.value})`);
            }
            const name: Identifier = this.makeIdentifier(idToken);

            let isArray = false;
            let arrayBounds: ArrayBound[] | undefined;
            let isNew = false;
            let objectType: string | undefined;

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
                this.match(TokenType.OperatorRParen);
            }

            if (this.match(TokenType.KeywordAs)) {
                if (this.match(TokenType.KeywordNew)) {
                    isNew = true;
                }
                const typeToken = this.peek();
                if (this.isNameToken(typeToken)) {
                    objectType = this.advance().value;
                    if (this.peek().type === TokenType.OperatorDot) {
                        this.advance(); // consume '.'
                        objectType += '.' + this.advance().value;
                    }
                }
            }

            declarations.push({ name, isArray, arrayBounds, isNew, isWithEvents, objectType });

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
        if (idToken.type !== TokenType.Identifier) throw new Error(`Parse error: Expected identifier after Const at line ${idToken.line}`);
        const name = this.makeIdentifier(idToken);

        // Optional 'As Type'
        if (this.match(TokenType.KeywordAs)) {
            this.advance(); // Ignore Type for now
        }

        if (!this.match(TokenType.OperatorEquals)) throw new Error(`Parse error: Expected '=' in Const at line ${this.peek().line}`);
        const value = this.parseExpression();

        return { type: 'ConstDeclaration', name, value };
    }

    private parseSetStatement(): SetStatement {
        this.advance(); // 'Set'
        const left = this.parsePrimary(); // parse identifier or member access
        if (!this.match(TokenType.OperatorEquals)) throw new Error(`Parse error: Expected '=' in Set statement at line ${this.peek().line}`);
        const right = this.parseExpression();
        return { type: 'SetStatement', left, right };
    }

    private parseOnErrorStatement(): OnErrorStatement {
        this.advance(); // 'On'
        if (!this.match(TokenType.KeywordError)) throw new Error(`Parse error: Expected 'Error' after 'On' at line ${this.peek().line}`);

        let label = '';
        if (this.match(TokenType.KeywordGoTo)) {
            const labelToken = this.advance(); // Identifier or 0
            label = labelToken.value;
        } else {
            // "Resume Next" fallback
            while (this.peek().type !== TokenType.Newline && this.peek().type !== TokenType.EOF) {
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
            throw new Error(`Parse error: Expected 'For', 'Do', 'Sub', 'Function', or 'Property' after 'Exit' at line ${typeToken.line}`);
        }
        return { type: 'ExitStatement', exitType };
    }

    private parseGoToStatement(): GoToStatement {
        this.advance(); // consume 'GoTo'
        const labelToken = this.advance();
        if (labelToken.type !== TokenType.Identifier && labelToken.type !== TokenType.Number) {
            throw new Error(`Parse error: Expected identifier or number after 'GoTo' at line ${labelToken.line}`);
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
        if (nameToken.type !== TokenType.Identifier) {
            throw new Error(`Parse error: Expected identifier after 'Type' at line ${nameToken.line}`);
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
                throw new Error(`Parse error: Expected member name in Type at line ${memberNameToken.line}`);
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
                throw new Error(`Parse error: Expected 'As' in Type member declaration at line ${this.peek().line}`);
            }

            const memberTypeToken = this.advance();
            members.push({ name: memberNameToken.value, memberType: memberTypeToken.value });

            this.skipNewlines();
        }

        // Consume 'End Type'
        if (this.peek().type === TokenType.KeywordEnd) {
            this.advance(); // consume 'End'
            if (!this.match(TokenType.KeywordType)) {
                throw new Error(`Parse error: Expected 'Type' after 'End' at line ${this.peek().line}`);
            }
        }

        return { type: 'TypeDeclaration', name: typeName, members } as TypeDeclaration;
    }

    private parseEnumDeclaration(): EnumDeclaration {
        this.advance(); // consume 'Enum'
        const nameToken = this.advance();
        if (nameToken.type !== TokenType.Identifier) {
            throw new Error(`Parse error: Expected identifier after 'Enum' at line ${nameToken.line}`);
        }
        const name: Identifier = this.makeIdentifier(nameToken);
        const members: EnumMember[] = [];

        this.skipNewlines();

        while (this.peek().type !== TokenType.KeywordEnd && this.peek().type !== TokenType.EOF) {
            // VBA §5.2.3.3: reserved words are valid member names (reserved-name-member-dcl)
            const memberNameToken = this.advance();
            if (!this.isWordToken(memberNameToken)) {
                throw new Error(`Parse error: Expected member name in Enum at line ${memberNameToken.line}`);
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
                throw new Error(`Parse error: Expected 'Enum' after 'End' at line ${this.peek().line}`);
            }
        }

        return { type: 'EnumDeclaration', name, members } as EnumDeclaration;
    }

    private parseClassDeclaration(): ClassDeclaration {
        this.advance(); // consume 'Class'
        const nameToken = this.advance();
        if (nameToken.type !== TokenType.Identifier) {
            throw new Error(`Parse error: Expected class name after 'Class' at line ${nameToken.line}`);
        }
        return this.parseClassBody(nameToken.value, true);
    }

    private parseClassBody(className: string, untilEndClass: boolean): ClassDeclaration {
        const fields: VariableDeclaration[] = [];
        const procedures: ProcedureDeclaration[] = [];
        const body: Statement[] = [];

        this.skipNewlines();

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

            // Scope modifiers before fields/procedures
            let scope: 'public' | 'private' | 'friend' | undefined;
            if (tok.type === TokenType.KeywordPublic || tok.type === TokenType.KeywordPrivate || tok.type === TokenType.KeywordFriend) {
                scope = tok.value.toLowerCase() as 'public' | 'private' | 'friend';
                this.advance(); // consume scope keyword
            }

            const inner = this.peek();
            if (inner.type === TokenType.KeywordSub || inner.type === TokenType.KeywordFunction || inner.type === TokenType.KeywordProperty) {
                const proc = this.parseProcedureDeclaration(scope);
                proc.moduleName = className;
                procedures.push(proc);
                body.push(proc);
            } else if (inner.type === TokenType.KeywordDim) {
                this.advance(); // consume 'Dim'
                const field = this.parseDimStatement(false, true);
                field.scope = scope ?? 'public';
                fields.push(field);
                body.push(field);
            } else if (inner.type === TokenType.KeywordImplements) {
                const impl = this.parseImplementsDirective();
                body.push(impl);
            } else if (inner.type === TokenType.KeywordEvent) {
                const event = this.parseEventDeclaration(scope);
                body.push(event);
            } else if (inner.type === TokenType.KeywordStatic) {
                this.advance(); // consume 'Static'
                const field = this.parseDimStatement(true, true);
                field.scope = scope ?? 'public';
                fields.push(field);
                body.push(field);
            } else if (scope !== undefined && inner.type === TokenType.Identifier) {
                // Public/Private Name As Type (no Dim keyword)
                const field = this.parseDimStatement(false, true);
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
                throw new Error(`Parse error: Expected 'Class' after 'End' at line ${this.peek().line}`);
            }
        }

        return { type: 'ClassDeclaration', name: className, fields, procedures, body } as ClassDeclaration;
    }

    private parseForStatement(): ForStatement | ForEachStatement {
        this.advance(); // consume 'For'

        if (this.peek().type === TokenType.KeywordEach) {
            return this.parseForEachStatementBody();
        }

        const idToken = this.advance();
        if (idToken.type !== TokenType.Identifier) {
            throw new Error(`Parse error: Expected identifier after 'For' at line ${idToken.line} `);
        }
        const identifier: Identifier = this.makeIdentifier(idToken);

        if (!this.match(TokenType.OperatorEquals)) {
            throw new Error(`Parse error: Expected '=' in For statement at line ${this.peek().line} `);
        }

        const startExpr = this.parseExpression();

        if (!this.match(TokenType.KeywordTo)) {
            throw new Error(`Parse error: Expected 'To' in For statement at line ${this.peek().line} `);
        }

        const endExpr = this.parseExpression();

        let stepExpr: Expression | undefined;
        if (this.match(TokenType.KeywordStep)) {
            stepExpr = this.parseExpression();
        }

        this.skipNewlines();

        const body: Statement[] = [];
        while (this.peek().type !== TokenType.KeywordNext && this.peek().type !== TokenType.EOF) {
            const stmt = this.parseStatement();
            if (stmt) body.push(stmt);
            this.skipNewlines();
        }

        if (!this.match(TokenType.KeywordNext)) {
            throw new Error(`Parse error: Expected 'Next' at line ${this.peek().line} `);
        }

        let nextIdentifier: Identifier | undefined;
        if (this.peek().type === TokenType.Identifier) {
            const nextIdToken = this.advance();
            nextIdentifier = this.makeIdentifier(nextIdToken);
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
        if (varToken.type !== TokenType.Identifier) {
            throw new Error(`Parse error: Expected identifier after 'For Each' at line ${varToken.line}`);
        }
        const variable: Identifier = this.makeIdentifier(varToken);

        if (!this.match(TokenType.KeywordIn)) {
            throw new Error(`Parse error: Expected 'In' in For Each statement at line ${this.peek().line}`);
        }

        const collection = this.parseExpression();

        this.skipNewlines();

        const body: Statement[] = [];
        while (this.peek().type !== TokenType.KeywordNext && this.peek().type !== TokenType.EOF) {
            const stmt = this.parseStatement();
            if (stmt) body.push(stmt);
            this.skipNewlines();
        }

        if (!this.match(TokenType.KeywordNext)) {
            throw new Error(`Parse error: Expected 'Next' at line ${this.peek().line}`);
        }

        let nextIdentifier: Identifier | undefined;
        if (this.peek().type === TokenType.Identifier) {
            const nextIdToken = this.advance();
            nextIdentifier = this.makeIdentifier(nextIdToken);
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
            throw new Error(`Parse error: Expected 'Then' after condition at line ${this.peek().line}`);
        }

        const isMultiLine = this.peek().type === TokenType.Newline;

        const consequent: Statement[] = [];
        let alternate: Statement[] | IfStatement | null = null;

        if (!isMultiLine) {
            // Single-line If: read consequent until Else/Newline/EOF
            while (
                this.peek().type !== TokenType.Newline &&
                this.peek().type !== TokenType.EOF &&
                this.peek().type !== TokenType.KeywordElse
            ) {
                const stmt = this.parseStatement();
                if (stmt) consequent.push(stmt);
            }
            // Handle optional inline Else
            if (this.peek().type === TokenType.KeywordElse) {
                this.advance(); // consume 'Else'
                alternate = [];
                while (this.peek().type !== TokenType.Newline && this.peek().type !== TokenType.EOF) {
                    const stmt = this.parseStatement();
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
                throw new Error(`Parse error: Expected 'If' after 'End' at line ${this.peek().line}`);
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
        while (this.peek().type !== TokenType.KeywordLoop && this.peek().type !== TokenType.EOF) {
            const stmt = this.parseStatement();
            if (stmt) body.push(stmt);
            this.skipNewlines();
        }

        if (!this.match(TokenType.KeywordLoop)) {
            throw new Error(`Parse error: Expected 'Loop' at line ${this.peek().line} `);
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
        while (this.peek().type !== TokenType.KeywordWend && this.peek().type !== TokenType.EOF) {
            const stmt = this.parseStatement();
            if (stmt) body.push(stmt);
            this.skipNewlines();
        }

        if (!this.match(TokenType.KeywordWend)) {
            throw new Error(`Parse error: Expected 'Wend' at line ${this.peek().line}`);
        }

        return { type: 'WhileStatement', condition, body };
    }

    private parseSelectCaseStatement(): SelectCaseStatement {
        this.advance(); // consume 'Select'
        if (!this.match(TokenType.KeywordCase)) {
            throw new Error(`Parse error: Expected 'Case' after 'Select' at line ${this.peek().line}`);
        }
        const expression = this.parseExpression();
        this.skipNewlines();

        const cases: CaseClause[] = [];
        let elseBody: Statement[] | null = null;

        while (!this.isAtEndTerminator() && this.peek().type !== TokenType.EOF) {
            if (this.peek().type !== TokenType.KeywordCase) {
                throw new Error(`Parse error: Expected 'Case' in Select Case at line ${this.peek().line}`);
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
                throw new Error(`Parse error: Expected 'Select' after 'End' at line ${this.peek().line}`);
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
                throw new Error(`Parse error: Expected 'With' after 'End' at line ${this.peek().line}`);
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
            throw new Error(`Parse error: Expected comparison operator after 'Is' at line ${this.peek().line}`);
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

    private parsePrimary(): Expression {
        const startTok = this.tokens[this.pos];
        const token = this.advance();
        let expr: Expression;
        if (token.type === TokenType.Number) {
            // Remove VBA type suffix (%, &, @, !, #, ^) before parsing into a float/int
            const cleanVal = token.value.replace(/[%&@!#^]$/, '');
            // Use Number() to support 0x (Hex) and 0o (Octal) prefixes
            expr = { type: 'NumberLiteral', value: Number(cleanVal) } as NumberLiteral;
        } else if (token.type === TokenType.String) {
            expr = { type: 'StringLiteral', value: token.value } as StringLiteral;
        } else if (token.type === TokenType.Date) {
            expr = { type: 'DateLiteral', value: token.value } as DateLiteral;
        } else if (token.type === TokenType.Identifier ||
                   Parser.CONTEXTUAL_KW.has(token.type) ||
                   Parser.COMPAT_KW_EXPR.has(token.type)) {
            expr = { type: 'Identifier', name: token.value } as Identifier;
        } else if (token.type === TokenType.KeywordAddressOf) {
            const procName = this.consume(TokenType.Identifier, "Expected procedure name after 'AddressOf'");
            expr = { type: 'AddressOfExpression', procedureName: { type: 'Identifier', name: procName.value } } as AddressOfExpression;
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
                throw new Error(`Parse error: Expected class name after 'New' at line ${classNameToken.line}`);
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
                throw new Error(`Parse error: Expected ')' at line ${this.peek().line} `);
            }
            expr = { type: 'ParenthesizedExpression', expression: innerExpr } as any; // Type added implicitly or via cast
        } else if (token.type === TokenType.KeywordTypeOf) {
            expr = this.parseRelational(); // Stop before 'Is'
            if (!this.match(TokenType.KeywordIs)) {
                throw new Error(`Parse error: Expected 'Is' after 'TypeOf' at line ${this.peek().line}`);
            }
            const typeToken = this.advance();
            if (typeToken.type !== TokenType.Identifier && typeToken.type !== TokenType.KeywordCollection) {
                 throw new Error(`Parse error: Expected type name after 'Is' at line ${typeToken.line}`);
            }
            expr = { type: 'TypeOfIsExpression', expression: expr, typeName: typeToken.value } as TypeOfIsExpression;
        } else if (token.type === TokenType.OperatorDot) {
            const propToken = this.advance();
            if (!this.isNameToken(propToken)) {
                throw new Error(`Parse error: Expected identifier after '.' at line ${propToken.line}`);
            }
            const property = { type: 'Identifier', name: propToken.value } as Identifier;
            expr = { type: 'ImplicitWithObjectExpression', property } as ImplicitWithObjectExpression;
        } else {
            throw new Error(`Parse error: Unexpected token in expression '${token.value}' at line ${token.line} `);
        }

        expr.loc = this.exprLoc(startTok, this.tokens[this.pos - 1]);

        while (true) {
            if (this.match(TokenType.OperatorDot)) {
                if (this.peek().type === TokenType.EOF) throw new Error("Expected property name after '.'");
                const propToken = this.advance();
                const property = { type: 'Identifier', name: propToken.value } as Identifier;
                expr = { type: 'MemberExpression', object: expr, property } as MemberExpression;
                expr.loc = this.exprLoc(startTok, this.tokens[this.pos - 1]);
            } else if (this.match(TokenType.OperatorExclamation)) {
                if (this.peek().type === TokenType.EOF) throw new Error("Expected identifier after '!'");
                const propToken = this.advance();
                const property = { type: 'Identifier', name: propToken.value } as Identifier;
                expr = { type: 'DictionaryAccessExpression', object: expr, property } as DictionaryAccessExpression;
                expr.loc = this.exprLoc(startTok, this.tokens[this.pos - 1]);
            } else if (this.match(TokenType.OperatorLParen)) {
                const args: Expression[] = [];
                if (this.peek().type !== TokenType.OperatorRParen) {
                    args.push(this.parseCallArgument());
                    while (this.match(TokenType.OperatorComma)) {
                        args.push(this.parseCallArgument());
                    }
                }
                if (!this.match(TokenType.OperatorRParen)) {
                    throw new Error(`Parse error: Expected ')' at line ${this.peek().line} `);
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
            throw new Error(`Parse error: Expected 'GoTo' or 'GoSub' after 'On' expression at line ${this.peek().line}`);
        }

        const labels: string[] = [];
        while (true) {
            const labelToken = this.advance();
            if (labelToken.type !== TokenType.Identifier && labelToken.type !== TokenType.Number) {
                throw new Error(`Parse error: Expected label (identifier or number) in On...GoTo/GoSub at line ${labelToken.line}`);
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
        if (labelToken.type !== TokenType.Identifier && labelToken.type !== TokenType.Number) {
            throw new Error(`Parse error: Expected label after GoSub at line ${labelToken.line}`);
        }
        return { type: 'GoSubStatement', label: labelToken.value };
    }

    private parseLSetStatement(): LSetStatement {
        this.advance(); // 'LSet'
        const left = this.parsePrimary();
        if (!this.match(TokenType.OperatorEquals)) {
            throw new Error(`Parse error: Expected '=' in LSet statement at line ${this.peek().line}`);
        }
        const right = this.parseExpression();
        return { type: 'LSetStatement', left, right };
    }

    private parseRSetStatement(): RSetStatement {
        this.advance(); // 'RSet'
        const left = this.parsePrimary();
        if (!this.match(TokenType.OperatorEquals)) {
            throw new Error(`Parse error: Expected '=' in RSet statement at line ${this.peek().line}`);
        }
        const right = this.parseExpression();
        return { type: 'RSetStatement', left, right };
    }

    private parseErrorStatement(): ErrorStatement {
        this.advance(); // 'Error'
        const errorNumber = this.parseExpression();
        return { type: 'ErrorStatement', errorNumber };
    }

    private parseEventDeclaration(scope?: 'public' | 'private' | 'friend'): EventDeclaration {
        this.advance(); // 'Event'
        const idToken = this.consume(TokenType.Identifier, "Expected identifier after 'Event'");
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
        }
        
        return { type: 'EventDeclaration', name, parameters, scope };
    }

    private parseRaiseEventStatement(): RaiseEventStatement {
        this.advance(); // 'RaiseEvent'
        const idToken = this.consume(TokenType.Identifier, "Expected identifier after 'RaiseEvent'");
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
        const idToken = this.consume(TokenType.Identifier, "Expected interface name after 'Implements'");
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
