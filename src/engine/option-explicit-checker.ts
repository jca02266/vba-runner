import {
    Program,
    Statement,
    Expression,
    ProcedureDeclaration,
    VariableDeclaration,
    ConstDeclaration,
    ForStatement,
    ForEachStatement,
    IfStatement,
    DoWhileStatement,
    WhileStatement,
    WithStatement,
    SelectCaseStatement,
    AssignmentStatement,
    SetStatement,
    CallStatement,
    CallExpression,
    MemberExpression,
    BinaryExpression,
    UnaryExpression,
    ParenthesizedExpression,
    Identifier,
    ReDimStatement,
    LSetStatement,
    RSetStatement,
    NamedArgument,
    ClassDeclaration,
    EnumDeclaration,
    ParseDiagnostic,
} from './parser';

export interface OptionExplicitResult {
    /** Set of procedure names (lower-cased) that have undeclared variable violations */
    violatedProcedures: Set<string>;
}

// VBA built-in names that are always available without a Dim declaration.
// Includes language keywords used as values, built-in constants, and built-in functions.
const VBA_BUILTINS: ReadonlySet<string> = new Set([
    // Value keywords
    'empty', 'nothing', 'null', 'true', 'false',
    // String constants
    'vbcrlf', 'vblf', 'vbcr', 'vbtab', 'vbnewline', 'vbnullchar', 'vbnullstring',
    'vbback', 'vbformfeed', 'vbverticaltab',
    // MsgBox button constants
    'vbokonly', 'vbokcancel', 'vbabortretryignore', 'vbyesnocancel', 'vbyesno', 'vbretrycancel',
    // MsgBox return values
    'vbok', 'vbcancel', 'vbabort', 'vbretry', 'vbignore', 'vbyes', 'vbno',
    // MsgBox icon/modal constants
    'vbcritical', 'vbquestion', 'vbexclamation', 'vbinformation',
    'vbapplicationmodal', 'vbsystemmodal',
    'vbdefaultbutton1', 'vbdefaultbutton2', 'vbdefaultbutton3',
    // VarType constants
    'vbboolean', 'vbbyte', 'vbcurrency', 'vbdate', 'vbdecimal', 'vbdouble',
    'vbempty', 'vberror', 'vbinteger', 'vblong', 'vbnull', 'vbobject',
    'vbsingle', 'vbstring', 'vbuserdefinedtype', 'vbvariant', 'vbarray',
    // StrConv / comparison constants
    'vbunicode', 'vbfromunicode', 'vbuppercase', 'vblowercase', 'vbpropercase',
    'vbtextcompare', 'vbbinarycompare', 'vbdatabasecompare',
    // Date / week constants
    'vbsunday', 'vbmonday', 'vbtuesday', 'vbwednesday', 'vbthursday', 'vbfriday', 'vbsaturday',
    'vbfirstjan1', 'vbfirstfourdays', 'vbfirstfullweek', 'vbusedefault', 'vbusecompatible',
    // File attribute constants
    'vbnormal', 'vbreadonly', 'vbhidden', 'vbsystem', 'vbvolume', 'vbdirectory', 'vbarchive',
    // Tristate / comparison
    'vbtristatefalse', 'vbtristatetrue', 'vbtristateusefault',
    // Built-in functions and objects
    'abs', 'array', 'asc', 'ascb', 'ascw', 'atn',
    'cbool', 'cbyte', 'ccur', 'cdate', 'cdbl', 'cdec', 'chr', 'chrb', 'chrw',
    'cint', 'clng', 'clnglng', 'clngptr', 'cobj', 'cos', 'createobject', 'csng', 'cstr', 'curdir',
    'cvar', 'cverr',
    'date', 'dateadd', 'datediff', 'datepart', 'dateserial', 'datevalue',
    'day', 'dir', 'doevents',
    'environ', 'eof', 'erl', 'error', 'exp',
    'fileattr', 'filedatetime', 'filelen', 'fix', 'format', 'freefile',
    'getattr', 'getobject', 'hex', 'hour',
    'iif', 'inputb', 'instr', 'instrrev', 'int', 'isarray', 'isdate',
    'isempty', 'iserror', 'ismissing', 'isnull', 'isnumeric', 'isobject',
    'join', 'lbound', 'lcase', 'left', 'leftb', 'len', 'lenb', 'loc', 'lof', 'log', 'ltrim',
    'mid', 'midb', 'minute', 'month', 'monthname', 'msgbox',
    'now', 'oct',
    'partition', 'replace', 'rgb', 'right', 'rightb', 'rnd', 'round', 'rtrim',
    'second', 'seek', 'sgn', 'shell', 'sin', 'space', 'split', 'sqr',
    'str', 'strcomp', 'strconv', 'strdup', 'strreverse', 'string', 'switch',
    'tab', 'tan', 'time', 'timer', 'timeserial', 'timevalue', 'trim',
    'typename', 'ubound', 'ucase', 'val', 'vartype',
    'weekday', 'weekdayname', 'year',
    // Special built-in objects (used as identifiers, not called)
    'debug', 'err',
]);

/**
 * Checks Option Explicit rules across the entire program AST.
 * Violations are appended to program.diagnostics.
 * Returns the set of procedure names that have violations.
 */
export function checkOptionExplicit(program: Program): OptionExplicitResult {
    const violatedProcedures = new Set<string>();

    // Determine if Option Explicit is active
    const hasOptionExplicit = program.body.some(s => s.type === 'OptionExplicitStatement');
    if (!hasOptionExplicit) {
        return { violatedProcedures };
    }

    // --- Pass 1: collect module-level declared names ---
    const moduleLevelNames = new Set<string>();

    for (const stmt of program.body) {
        collectModuleLevelDeclaredNames(stmt, moduleLevelNames);
    }

    // --- Pass 2: check each procedure ---
    for (const stmt of program.body) {
        if (stmt.type === 'ProcedureDeclaration') {
            const proc = stmt as ProcedureDeclaration;
            const violated = checkProcedure(proc, moduleLevelNames, program.diagnostics);
            if (violated) {
                violatedProcedures.add(proc.name.name.toLowerCase());
            }
        } else if (stmt.type === 'ClassDeclaration') {
            const cls = stmt as ClassDeclaration;
            const classModuleNames = new Set<string>(moduleLevelNames);
            // Add class fields to the class-level name set
            for (const field of cls.fields) {
                for (const decl of field.declarations) {
                    classModuleNames.add(decl.name.name.toLowerCase());
                }
            }
            for (const proc of cls.procedures) {
                const violated = checkProcedure(proc, classModuleNames, program.diagnostics);
                if (violated) {
                    violatedProcedures.add(proc.name.name.toLowerCase());
                }
            }
        }
    }

    return { violatedProcedures };
}

function collectModuleLevelDeclaredNames(stmt: Statement, names: Set<string>): void {
    switch (stmt.type) {
        case 'VariableDeclaration': {
            const v = stmt as VariableDeclaration;
            for (const decl of v.declarations) {
                names.add(decl.name.name.toLowerCase());
            }
            break;
        }
        case 'ConstDeclaration': {
            const c = stmt as ConstDeclaration;
            names.add(c.name.name.toLowerCase());
            break;
        }
        case 'ProcedureDeclaration': {
            const p = stmt as ProcedureDeclaration;
            names.add(p.name.name.toLowerCase());
            break;
        }
        case 'EnumDeclaration': {
            const e = stmt as EnumDeclaration;
            // Enum name and its members are accessible as module-level names
            names.add(e.name.name.toLowerCase());
            for (const member of e.members) {
                names.add(member.name.name.toLowerCase());
            }
            break;
        }
    }
}

/**
 * Returns true if the procedure has at least one undeclared-variable violation.
 */
function checkProcedure(
    proc: ProcedureDeclaration,
    moduleLevelNames: Set<string>,
    diagnostics: ParseDiagnostic[]
): boolean {
    // Build the declared-name set for this procedure:
    // module-level names + parameters + function name (for return value)
    const declared = new Set<string>(moduleLevelNames);

    declared.add(proc.name.name.toLowerCase());

    for (const param of proc.parameters) {
        declared.add(param.name.toLowerCase());
    }

    let violated = false;

    // Forward scan through the body
    const visitStmt = (stmt: Statement): void => {
        switch (stmt.type) {
            case 'VariableDeclaration': {
                const v = stmt as VariableDeclaration;
                for (const decl of v.declarations) {
                    declared.add(decl.name.name.toLowerCase());
                }
                break;
            }
            case 'ConstDeclaration': {
                const c = stmt as ConstDeclaration;
                declared.add(c.name.name.toLowerCase());
                break;
            }
            case 'ReDimStatement': {
                const r = stmt as ReDimStatement;
                // ReDim does NOT declare — it requires prior Dim
                checkExprIdents(r.name, declared, diagnostics, () => { violated = true; });
                for (const bound of r.bounds) {
                    if (bound.lower) checkExpr(bound.lower, declared, diagnostics, () => { violated = true; });
                    checkExpr(bound.upper, declared, diagnostics, () => { violated = true; });
                }
                break;
            }
            case 'AssignmentStatement': {
                const a = stmt as AssignmentStatement;
                checkExpr(a.left, declared, diagnostics, () => { violated = true; });
                checkExpr(a.right, declared, diagnostics, () => { violated = true; });
                break;
            }
            case 'SetStatement': {
                const s = stmt as SetStatement;
                checkExpr(s.left, declared, diagnostics, () => { violated = true; });
                checkExpr(s.right, declared, diagnostics, () => { violated = true; });
                break;
            }
            case 'CallStatement': {
                const c = stmt as CallStatement;
                checkCallExpr(c.expression, declared, diagnostics, () => { violated = true; });
                break;
            }
            case 'ForStatement': {
                const f = stmt as ForStatement;
                // For loop variable must be explicitly declared
                checkExprIdents(f.identifier, declared, diagnostics, () => { violated = true; });
                checkExpr(f.start, declared, diagnostics, () => { violated = true; });
                checkExpr(f.end, declared, diagnostics, () => { violated = true; });
                if (f.step) checkExpr(f.step, declared, diagnostics, () => { violated = true; });
                for (const s of f.body) visitStmt(s);
                break;
            }
            case 'ForEachStatement': {
                const f = stmt as ForEachStatement;
                checkExprIdents(f.variable, declared, diagnostics, () => { violated = true; });
                checkExpr(f.collection, declared, diagnostics, () => { violated = true; });
                for (const s of f.body) visitStmt(s);
                break;
            }
            case 'IfStatement': {
                const i = stmt as IfStatement;
                checkExpr(i.condition, declared, diagnostics, () => { violated = true; });
                for (const s of i.consequent) visitStmt(s);
                if (i.alternate) {
                    if (Array.isArray(i.alternate)) {
                        for (const s of i.alternate) visitStmt(s);
                    } else {
                        visitStmt(i.alternate as Statement);
                    }
                }
                break;
            }
            case 'DoWhileStatement': {
                const d = stmt as DoWhileStatement;
                if (d.condition) checkExpr(d.condition, declared, diagnostics, () => { violated = true; });
                for (const s of d.body) visitStmt(s);
                break;
            }
            case 'WhileStatement': {
                const w = stmt as WhileStatement;
                checkExpr(w.condition, declared, diagnostics, () => { violated = true; });
                for (const w2 of w.body) visitStmt(w2);
                break;
            }
            case 'WithStatement': {
                const w = stmt as WithStatement;
                checkExpr(w.expression, declared, diagnostics, () => { violated = true; });
                for (const s of w.body) visitStmt(s);
                break;
            }
            case 'SelectCaseStatement': {
                const sc = stmt as SelectCaseStatement;
                checkExpr(sc.expression, declared, diagnostics, () => { violated = true; });
                for (const c of sc.cases) {
                    for (const range of c.ranges) {
                        if (range.kind === 'expression') {
                            checkExpr(range.value, declared, diagnostics, () => { violated = true; });
                        } else if (range.kind === 'to') {
                            checkExpr(range.start, declared, diagnostics, () => { violated = true; });
                            checkExpr(range.end, declared, diagnostics, () => { violated = true; });
                        } else {
                            checkExpr(range.value, declared, diagnostics, () => { violated = true; });
                        }
                    }
                    for (const s of c.body) visitStmt(s);
                }
                if (sc.elseBody) {
                    for (const s of sc.elseBody) visitStmt(s);
                }
                break;
            }
            case 'LSetStatement': {
                const l = stmt as LSetStatement;
                checkExpr(l.left, declared, diagnostics, () => { violated = true; });
                checkExpr(l.right, declared, diagnostics, () => { violated = true; });
                break;
            }
            case 'RSetStatement': {
                const r = stmt as RSetStatement;
                checkExpr(r.left, declared, diagnostics, () => { violated = true; });
                checkExpr(r.right, declared, diagnostics, () => { violated = true; });
                break;
            }
            // Statements with no variable references to check:
            // ExitStatement, LabelStatement, GoSubStatement, ReturnStatement,
            // OnErrorStatement, ResumeStatement, GoToStatement, StopStatement, EndStatement
            // PrintStatement (#), WriteStatement (#), etc. — these could have expressions
            // but are uncommon in the contexts we need to check and we'll handle via the
            // generic expression walker if they contain Identifiers in their expressions.
            // For now, handled by the default (no-op) case.
            default:
                // For any other statement types that have expressions, we do a best-effort
                // check by scanning their sub-expressions if exposed.
                break;
        }
    };

    for (const stmt of proc.body) {
        visitStmt(stmt);
    }

    return violated;
}

type OnViolation = () => void;

/** Check an expression that is known to be an Identifier (for loop var, etc.) */
function checkExprIdents(expr: Expression, declared: Set<string>, diagnostics: ParseDiagnostic[], onViolation: OnViolation): void {
    if (expr.type === 'Identifier') {
        const id = expr as Identifier;
        const lower = id.name.toLowerCase();
        if (!declared.has(lower) && !VBA_BUILTINS.has(lower)) {
            reportUndeclared(id.name, expr, diagnostics);
            onViolation();
        }
    }
}

/** Recursively check an expression for undeclared Identifiers */
function checkExpr(expr: Expression, declared: Set<string>, diagnostics: ParseDiagnostic[], onViolation: OnViolation): void {
    switch (expr.type) {
        case 'Identifier': {
            const id = expr as Identifier;
            const lower = id.name.toLowerCase();
            if (!declared.has(lower) && !VBA_BUILTINS.has(lower)) {
                reportUndeclared(id.name, expr, diagnostics);
                onViolation();
            }
            break;
        }
        case 'CallExpression': {
            checkCallExpr(expr as CallExpression, declared, diagnostics, onViolation);
            break;
        }
        case 'MemberExpression': {
            const m = expr as MemberExpression;
            // Only check the object (left side); property is always a name, not a variable reference
            checkExpr(m.object, declared, diagnostics, onViolation);
            break;
        }
        case 'DictionaryAccessExpression': {
            const d = expr as { type: string; object: Expression; property: Identifier };
            checkExpr(d.object, declared, diagnostics, onViolation);
            break;
        }
        case 'BinaryExpression': {
            const b = expr as BinaryExpression;
            checkExpr(b.left, declared, diagnostics, onViolation);
            checkExpr(b.right, declared, diagnostics, onViolation);
            break;
        }
        case 'UnaryExpression': {
            const u = expr as UnaryExpression;
            checkExpr(u.argument, declared, diagnostics, onViolation);
            break;
        }
        case 'ParenthesizedExpression': {
            const p = expr as ParenthesizedExpression;
            checkExpr(p.expression, declared, diagnostics, onViolation);
            break;
        }
        case 'NamedArgument': {
            const n = expr as NamedArgument;
            checkExpr(n.value, declared, diagnostics, onViolation);
            break;
        }
        // Literals, NewExpression, ImplicitWithObjectExpression, etc. — no variable refs
        default:
            break;
    }
}

function checkCallExpr(expr: CallExpression, declared: Set<string>, diagnostics: ParseDiagnostic[], onViolation: OnViolation): void {
    // callee: if it's a bare Identifier, it's a function/sub call — NOT a variable reference
    // if it's a MemberExpression, check the object part only
    const callee = expr.callee;
    if (callee.type === 'MemberExpression') {
        const m = callee as MemberExpression;
        checkExpr(m.object, declared, diagnostics, onViolation);
        // m.property is the method name — skip
    }
    // Bare Identifier callee = procedure name, skip variable check
    // Check all arguments
    for (const arg of expr.args) {
        checkExpr(arg, declared, diagnostics, onViolation);
    }
}

function reportUndeclared(name: string, expr: Expression, diagnostics: ParseDiagnostic[]): void {
    const loc = (expr as any).loc ?? { start: { line: 0, column: 0 }, end: { line: 0, column: 0 } };
    diagnostics.push({
        message: `Variable '${name}' not declared`,
        loc,
        severity: 'error',
    });
}
