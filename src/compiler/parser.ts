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

export interface ASTNode {
    type: string;
}

export interface Program extends ASTNode {
    type: 'Program';
    body: Statement[];
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
}

export interface ProcedureDeclaration extends Statement {
    type: 'ProcedureDeclaration';
    isFunction: boolean;
    isProperty: boolean;
    propertyType?: 'get' | 'let' | 'set';
    name: Identifier;
    parameters: Parameter[];
    body: Statement[];
    scope?: 'public' | 'private' | 'friend';
    isStatic?: boolean;
    moduleName?: string;
}

export interface VariableDeclarator {
    name: Identifier;
    isArray: boolean;
    arraySize?: Expression; // TODO: Array size should ideally support multiple dimensions
    isNew: boolean;
    objectType?: string;
}

export interface VariableDeclaration extends Statement {
    type: 'VariableDeclaration';
    declarations: VariableDeclarator[];
    isStatic?: boolean;
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

export interface OptionCompareStatement extends Statement {
    type: 'OptionCompareStatement';
    mode: 'Binary' | 'Text';
}

export interface AttributeStatement extends Statement {
    type: 'AttributeStatement';
    name: string;
    value: Expression;
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
    bounds: Expression[]; // Multi-dimensional bounds (e.g. 1 To numDays)
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

export interface CallExpression extends Expression {
    type: 'CallExpression';
    callee: Expression;
    args: Expression[];
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

    constructor(tokens: Token[]) {
        this.tokens = tokens;
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

        if (this.peek().type === TokenType.KeywordByVal || this.peek().type === TokenType.KeywordByRef) {
            isByVal = this.advance().type === TokenType.KeywordByVal;
        }

        const nameToken = this.consume(TokenType.Identifier, "Expected parameter name");

        if (this.match(TokenType.KeywordAs)) {
            this.advance(); // consume type name
        }

        if (this.match(TokenType.OperatorEquals)) {
            this.parseExpression(); // skip default value
        }

        return { type: 'Parameter', name: nameToken.value, isByVal, isOptional };
    }

    private parseAttributeStatement(): AttributeStatement {
        this.advance(); // 'Attribute'
        const name = this.advance().value;
        this.consume(TokenType.OperatorEquals, "Expected '=' after Attribute name");
        const value = this.parseExpression();
        return { type: 'AttributeStatement', name, value };
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

    private consume(expectedType: TokenType, message: string): Token {
        if (this.peek().type === expectedType) {
            return this.advance();
        }
        throw new Error(`Parse error at line ${this.peek().line}: ${message}`);
    }

    private skipNewlines() {
        while (this.match(TokenType.Newline)) { /* skip */ }
    }

    public parse(): Program {
        const program: Program = {
            type: 'Program',
            body: []
        };

        this.skipNewlines();
        while (this.peek().type !== TokenType.EOF) {
            const stmt = this.parseStatement();
            if (stmt) {
                program.body.push(stmt);
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
        const token = this.peek();

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
        } else if (
            token.type === TokenType.KeywordPublic ||
            token.type === TokenType.KeywordPrivate ||
            token.type === TokenType.KeywordFriend
        ) {
            const scopeToken = this.advance();
            const scope = scopeToken.value.toLowerCase() as 'public' | 'private' | 'friend';
            const next = this.peek();
            if (next.type === TokenType.KeywordSub || next.type === TokenType.KeywordFunction || next.type === TokenType.KeywordProperty) {
                return this.parseProcedureDeclaration(scope);
            }
            if (next.type === TokenType.KeywordStatic) {
                this.advance(); // consume 'Static'
                return this.parseProcedureDeclaration(scope, true);
            }
            // Public/Private on Dim/Const — consume scope keyword and parse normally
            return this.parseStatement();
        } else if (token.type === TokenType.KeywordStatic) {
            this.advance(); // consume 'Static'
            const next = this.peek();
            if (next.type === TokenType.KeywordSub || next.type === TokenType.KeywordFunction || next.type === TokenType.KeywordProperty) {
                return this.parseProcedureDeclaration(undefined, true);
            }
            // Static variable declaration inside a procedure
            return this.parseDimStatement(true);
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
        } else if (token.type === TokenType.KeywordOption) {
            this.advance(); // 'Option'
            if (this.match(TokenType.KeywordCompare)) {
                return this.parseOptionCompareStatement();
            }
            if (this.match(TokenType.KeywordExplicit)) {
                // Ignore for now
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
        } else if (token.type === TokenType.KeywordClass) {
            return this.parseClassDeclaration();
        } else if (token.type === TokenType.KeywordCall) {
            this.advance(); // consume 'Call'
            const expr = this.parsePrimary();
            if (expr.type === 'CallExpression') {
                return { type: 'CallStatement', expression: expr } as CallStatement;
            } else if (expr.type === 'Identifier') {
                return { type: 'CallStatement', expression: { type: 'CallExpression', callee: expr, args: [] } } as CallStatement;
            }
            throw new Error(`Parse error: Expected procedure call after 'Call'`);
        } else if (token.type === TokenType.Identifier || token.type === TokenType.OperatorDot || token.type === TokenType.Number) {
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
        if (idToken.type !== TokenType.Identifier) {
            throw new Error(`Parse error: Expected identifier after Sub/Function at line ${idToken.line}`);
        }
        const name: Identifier = { type: 'Identifier', name: idToken.value };
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

        // Optional Function return type (e.g. 'As Long')
        if (this.match(TokenType.KeywordAs)) {
            this.advance(); // consume Type name
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

        return { type: 'ProcedureDeclaration', isFunction, isProperty, propertyType, name, parameters, body, scope, isStatic };
    }

    private parseDimStatement(isStatic?: boolean): VariableDeclaration {
        if (!isStatic) this.advance(); // 'Dim' (already consumed for Static)
        const declarations: VariableDeclarator[] = [];

        while (true) {
            const idToken = this.advance();
            const name: Identifier = { type: 'Identifier', name: idToken.value };

            let isArray = false;
            let arraySize: Expression | undefined;
            let isNew = false;
            let objectType: string | undefined;

            if (this.match(TokenType.OperatorLParen)) {
                isArray = true;
                if (this.peek().type !== TokenType.OperatorRParen) {
                    arraySize = this.parseExpression();
                }
                this.match(TokenType.OperatorRParen);
            }

            if (this.match(TokenType.KeywordAs)) {
                if (this.match(TokenType.KeywordNew)) {
                    isNew = true;
                }
                const typeToken = this.peek();
                if (typeToken.type === TokenType.KeywordCollection || typeToken.type === TokenType.Identifier) {
                    objectType = this.advance().value;
                }
            }

            declarations.push({ name, isArray, arraySize, isNew, objectType });

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
        const name = { type: 'Identifier', name: idToken.value } as Identifier;

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
        const name = { type: 'Identifier', name: idToken.value } as Identifier;
        return { type: 'EraseStatement', name };
    }

    private parseReDimStatement(): ReDimStatement {
        this.advance(); // 'ReDim'

        // Optional 'Preserve' keyword
        if (this.peek().type === TokenType.Identifier && this.peek().value.toLowerCase() === 'preserve') {
            this.advance();
        }

        const idToken = this.advance();
        const name = { type: 'Identifier', name: idToken.value } as Identifier;
        const bounds: Expression[] = [];

        if (this.match(TokenType.OperatorLParen)) {
            if (this.peek().type !== TokenType.OperatorRParen) {
                bounds.push(this.parseExpression());
                while (this.match(TokenType.KeywordTo) || this.match(TokenType.OperatorComma)) {
                    bounds.push(this.parseExpression());
                }
            }
            this.match(TokenType.OperatorRParen);
        }

        if (this.match(TokenType.KeywordAs)) {
            this.advance();
        }

        return { type: 'ReDimStatement', name, bounds };
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
            // Each member line: memberName As memberType
            const memberNameToken = this.advance();
            if (memberNameToken.type !== TokenType.Identifier) {
                throw new Error(`Parse error: Expected member name in Type at line ${memberNameToken.line}`);
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
        const name: Identifier = { type: 'Identifier', name: nameToken.value };
        const members: EnumMember[] = [];

        this.skipNewlines();

        while (this.peek().type !== TokenType.KeywordEnd && this.peek().type !== TokenType.EOF) {
            const memberNameToken = this.advance();
            if (memberNameToken.type !== TokenType.Identifier) {
                throw new Error(`Parse error: Expected member name in Enum at line ${memberNameToken.line}`);
            }
            const memberName: Identifier = { type: 'Identifier', name: memberNameToken.value };
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
        const className = nameToken.value;
        const fields: VariableDeclaration[] = [];
        const procedures: ProcedureDeclaration[] = [];

        this.skipNewlines();

        while (this.peek().type !== TokenType.EOF) {
            if (this.peek().type === TokenType.KeywordEnd && this.peek(1).type === TokenType.KeywordClass) {
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
            } else if (inner.type === TokenType.KeywordDim) {
                this.advance(); // consume 'Dim'
                const field = this.parseDimStatement(true) as VariableDeclaration; // true = keyword already consumed
                field.scope = scope ?? 'public';
                fields.push(field);
            } else if (inner.type === TokenType.KeywordStatic) {
                this.advance(); // consume 'Static'
                const field = this.parseDimStatement(true) as VariableDeclaration;
                field.isStatic = true;
                field.scope = scope ?? 'public';
                fields.push(field);
            } else if (scope !== undefined && inner.type === TokenType.Identifier) {
                // Public/Private Name As Type (no Dim keyword)
                const field = this.parseDimStatement(true) as VariableDeclaration;
                field.scope = scope;
                fields.push(field);
            } else {
                // Skip unknown tokens gracefully
                this.advance();
            }
            this.skipNewlines();
        }

        // Consume 'End Class'
        if (this.peek().type === TokenType.KeywordEnd) {
            this.advance();
            if (!this.match(TokenType.KeywordClass)) {
                throw new Error(`Parse error: Expected 'Class' after 'End' at line ${this.peek().line}`);
            }
        }

        return { type: 'ClassDeclaration', name: className, fields, procedures } as ClassDeclaration;
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
        const identifier: Identifier = { type: 'Identifier', name: idToken.value };

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
            nextIdentifier = { type: 'Identifier', name: nextIdToken.value } as Identifier;
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
        const variable: Identifier = { type: 'Identifier', name: varToken.value };

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
            nextIdentifier = { type: 'Identifier', name: nextIdToken.value } as Identifier;
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
            left = { type: 'BinaryExpression', operator, left, right } as BinaryExpression;
        }
        return left;
    }

    private parseLogicalEqv(): Expression {
        let left = this.parseLogicalXor();
        while (this.peek().type === TokenType.KeywordEqv) {
            const operator = this.advance().value;
            const right = this.parseLogicalXor();
            left = { type: 'BinaryExpression', operator, left, right } as BinaryExpression;
        }
        return left;
    }

    private parseLogicalXor(): Expression {
        let left = this.parseLogicalOr();
        while (this.peek().type === TokenType.KeywordXor) {
            const operator = this.advance().value;
            const right = this.parseLogicalOr();
            left = { type: 'BinaryExpression', operator, left, right } as BinaryExpression;
        }
        return left;
    }

    private parseLogicalOr(): Expression {
        let left = this.parseLogicalAnd();
        while (this.peek().type === TokenType.KeywordOr) {
            const operator = this.advance().value;
            const right = this.parseLogicalAnd();
            left = { type: 'BinaryExpression', operator, left, right } as BinaryExpression;
        }
        return left;
    }

    private parseLogicalAnd(): Expression {
        let left = this.parseLogicalNot();
        while (this.peek().type === TokenType.KeywordAnd) {
            const operator = this.advance().value;
            const right = this.parseLogicalNot();
            left = { type: 'BinaryExpression', operator, left, right } as BinaryExpression;
        }
        return left;
    }

    private parseLogicalNot(): Expression {
        if (this.match(TokenType.KeywordNot)) {
            const argument = this.parseEquality();
            return { type: 'UnaryExpression', operator: 'Not', argument } as UnaryExpression;
        }
        return this.parseEquality();
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
            left = { type: 'BinaryExpression', operator, left, right } as BinaryExpression;
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
            left = { type: 'BinaryExpression', operator, left, right } as BinaryExpression;
        }
        return left;
    }

    private parseConcatenation(): Expression {
        let left = this.parseAdditive();
        while (this.peek().type === TokenType.OperatorAmpersand) {
            const operator = this.advance().value;
            const right = this.parseAdditive();
            left = { type: 'BinaryExpression', operator, left, right } as BinaryExpression;
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
            left = { type: 'BinaryExpression', operator, left, right } as BinaryExpression;
        }
        return left;
    }

    private parseModulo(): Expression {
        let left = this.parseIntDivision();
        while (this.peek().type === TokenType.KeywordMod) {
            const operator = this.advance().value;
            const right = this.parseIntDivision();
            left = { type: 'BinaryExpression', operator, left, right } as BinaryExpression;
        }
        return left;
    }

    private parseIntDivision(): Expression {
        let left = this.parseMultiplicative();
        while (this.peek().type === TokenType.OperatorIntDivide) {
            const operator = this.advance().value;
            const right = this.parseMultiplicative();
            left = { type: 'BinaryExpression', operator, left, right } as BinaryExpression;
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
            left = { type: 'BinaryExpression', operator, left, right } as BinaryExpression;
        }
        return left;
    }

    private parseUnary(): Expression {
        if (this.peek().type === TokenType.OperatorMinus || this.peek().type === TokenType.OperatorPlus) {
            const operator = this.advance().value;
            const argument = this.parseUnary();
            return { type: 'UnaryExpression', operator, argument } as UnaryExpression;
        }
        return this.parseExponentiation();
    }

    private parseExponentiation(): Expression {
        let left = this.parsePrimary();
        while (this.peek().type === TokenType.OperatorPower) {
            const operator = this.advance().value;
            const right = this.parsePrimary();
            left = { type: 'BinaryExpression', operator, left, right } as BinaryExpression;
        }
        return left;
    }

    private parsePrimary(): Expression {
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
        } else if (token.type === TokenType.Identifier) {
            expr = { type: 'Identifier', name: token.value } as Identifier;
        } else if (token.type === TokenType.KeywordEmpty) {
            expr = { type: 'Identifier', name: token.value } as Identifier;
        } else if (token.type === TokenType.KeywordNothing) {
            expr = { type: 'Identifier', name: 'Nothing' } as Identifier;
        } else if (token.type === TokenType.KeywordNull) {
            expr = { type: 'Identifier', name: 'Null' } as Identifier;
        } else if (token.type === TokenType.KeywordMe) {
            expr = { type: 'Identifier', name: 'Me' } as Identifier;
        } else if (token.type === TokenType.KeywordNew) {
            const classNameToken = this.advance();
            if (classNameToken.type !== TokenType.Identifier) {
                throw new Error(`Parse error: Expected class name after 'New' at line ${classNameToken.line}`);
            }
            expr = { type: 'NewExpression', className: classNameToken.value } as NewExpression;
        } else if (token.type === TokenType.OperatorLParen) {
            expr = this.parseExpression();
            if (!this.match(TokenType.OperatorRParen)) {
                throw new Error(`Parse error: Expected ')' at line ${this.peek().line} `);
            }
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
            if (propToken.type !== TokenType.Identifier) {
                throw new Error(`Parse error: Expected identifier after '.' at line ${propToken.line}`);
            }
            const property = { type: 'Identifier', name: propToken.value } as Identifier;
            expr = { type: 'ImplicitWithObjectExpression', property } as ImplicitWithObjectExpression;
        } else {
            throw new Error(`Parse error: Unexpected token in expression '${token.value}' at line ${token.line} `);
        }

        while (true) {
            if (this.match(TokenType.OperatorDot)) {
                const propToken = this.advance();
                const property = { type: 'Identifier', name: propToken.value } as Identifier;
                expr = { type: 'MemberExpression', object: expr, property } as MemberExpression;
            } else if (this.match(TokenType.OperatorExclamation)) {
                const propToken = this.advance();
                const property = { type: 'Identifier', name: propToken.value } as Identifier;
                expr = { type: 'DictionaryAccessExpression', object: expr, property } as DictionaryAccessExpression;
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
            } else {
                break;
            }
        }
        return expr;
    }

    // Parse a call argument, handling named arguments (e.g., shift:=xlUp)
    private parseCallArgument(): Expression {
        // Check for named argument: Identifier := Expression
        if (this.peek().type === TokenType.Identifier &&
            this.pos + 1 < this.tokens.length &&
            this.tokens[this.pos + 1].type === TokenType.OperatorColonEquals) {
            this.advance(); // consume identifier (the name)
            this.advance(); // consume ':='
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
}
