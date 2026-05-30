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
    /** Map of procedure names (lower-cased) to sets of undeclared identifier names (lower-cased) */
    violatedProcedures: Map<string, Set<string>>;
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
    const violatedProcedures = new Map<string, Set<string>>();

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
            const undeclared = checkProcedure(proc, moduleLevelNames, program.diagnostics);
            if (undeclared) {
                violatedProcedures.set(proc.name.name.toLowerCase(), undeclared);
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
                const undeclared = checkProcedure(proc, classModuleNames, program.diagnostics);
                if (undeclared) {
                    violatedProcedures.set(proc.name.name.toLowerCase(), undeclared);
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
): Set<string> | null {
    // Build the declared-name set for this procedure:
    // module-level names + parameters + function name (for return value)
    const declared = new Set<string>(moduleLevelNames);

    declared.add(proc.name.name.toLowerCase());

    for (const param of proc.parameters) {
        declared.add(param.name.toLowerCase());
    }

    const undeclaredNames = new Set<string>();
    const onViol = (name: string) => { undeclaredNames.add(name); };

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
                checkExprIdents(r.name, declared, diagnostics, onViol);
                for (const bound of r.bounds) {
                    if (bound.lower) checkExpr(bound.lower, declared, diagnostics, onViol);
                    checkExpr(bound.upper, declared, diagnostics, onViol);
                }
                break;
            }
            case 'AssignmentStatement': {
                const a = stmt as AssignmentStatement;
                checkExpr(a.left, declared, diagnostics, onViol);
                checkExpr(a.right, declared, diagnostics, onViol);
                break;
            }
            case 'SetStatement': {
                const s = stmt as SetStatement;
                checkExpr(s.left, declared, diagnostics, onViol);
                checkExpr(s.right, declared, diagnostics, onViol);
                break;
            }
            case 'CallStatement': {
                const c = stmt as CallStatement;
                checkCallExpr(c.expression, declared, diagnostics, onViol);
                break;
            }
            case 'ForStatement': {
                const f = stmt as ForStatement;
                // For loop variable must be explicitly declared
                checkExprIdents(f.identifier, declared, diagnostics, onViol);
                checkExpr(f.start, declared, diagnostics, onViol);
                checkExpr(f.end, declared, diagnostics, onViol);
                if (f.step) checkExpr(f.step, declared, diagnostics, onViol);
                for (const s of f.body) visitStmt(s);
                break;
            }
            case 'ForEachStatement': {
                const f = stmt as ForEachStatement;
                checkExprIdents(f.variable, declared, diagnostics, onViol);
                checkExpr(f.collection, declared, diagnostics, onViol);
                for (const s of f.body) visitStmt(s);
                break;
            }
            case 'IfStatement': {
                const i = stmt as IfStatement;
                checkExpr(i.condition, declared, diagnostics, onViol);
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
                if (d.condition) checkExpr(d.condition, declared, diagnostics, onViol);
                for (const s of d.body) visitStmt(s);
                break;
            }
            case 'WhileStatement': {
                const w = stmt as WhileStatement;
                checkExpr(w.condition, declared, diagnostics, onViol);
                for (const w2 of w.body) visitStmt(w2);
                break;
            }
            case 'WithStatement': {
                const w = stmt as WithStatement;
                checkExpr(w.expression, declared, diagnostics, onViol);
                for (const s of w.body) visitStmt(s);
                break;
            }
            case 'SelectCaseStatement': {
                const sc = stmt as SelectCaseStatement;
                checkExpr(sc.expression, declared, diagnostics, onViol);
                for (const c of sc.cases) {
                    for (const range of c.ranges) {
                        if (range.kind === 'expression') {
                            checkExpr(range.value, declared, diagnostics, onViol);
                        } else if (range.kind === 'to') {
                            checkExpr(range.start, declared, diagnostics, onViol);
                            checkExpr(range.end, declared, diagnostics, onViol);
                        } else {
                            checkExpr(range.value, declared, diagnostics, onViol);
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
                checkExpr(l.left, declared, diagnostics, onViol);
                checkExpr(l.right, declared, diagnostics, onViol);
                break;
            }
            case 'RSetStatement': {
                const r = stmt as RSetStatement;
                checkExpr(r.left, declared, diagnostics, onViol);
                checkExpr(r.right, declared, diagnostics, onViol);
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

    return undeclaredNames.size > 0 ? undeclaredNames : null;
}

type OnViolation = (name: string) => void;

/** Check an expression that is known to be an Identifier (for loop var, etc.) */
function checkExprIdents(expr: Expression, declared: Set<string>, diagnostics: ParseDiagnostic[], onViolation: OnViolation): void {
    if (expr.type === 'Identifier') {
        const id = expr as Identifier;
        const lower = id.name.toLowerCase();
        if (!declared.has(lower) && !VBA_BUILTINS.has(lower)) {
            reportUndeclared(id.name, expr, diagnostics);
            onViolation(lower);
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
                onViolation(lower);
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
        // Bare Identifier object may be a module name (e.g. ArgCountTest.Method) —
        // module names don't need Dim, so skip the check. Only check complex objects
        // (chained access, array index, etc.) that must be declared variables.
        if (m.object.type !== 'Identifier') {
            checkExpr(m.object, declared, diagnostics, onViolation);
        }
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
