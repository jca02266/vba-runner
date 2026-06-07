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
    /** Map of procedure names (lower-cased) to maps of undeclared name (lower-cased) → first line in AST */
    violatedProcedures: Map<string, Map<string, number>>;
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
    // Built-in Sub statements parsed as CallStatement (not functions)
    'beep', 'chdir', 'mkdir', 'rmdir', 'randomize', 'sleep', 'wait',
]);

/**
 * Checks Option Explicit rules across the entire program AST.
 * Violations are appended to program.diagnostics.
 * Returns the set of procedure names that have violations.
 *
 * @param knownModuleNames - When provided (even as an empty set), bare Identifier objects in
 *   call expressions are checked precisely: only names in this set are treated as module
 *   references and skipped. When undefined (single-module first pass), all bare identifier
 *   objects in call-expression callees are skipped to avoid false positives from
 *   cross-module references that are not yet known.
 */
export function checkOptionExplicit(program: Program, knownModuleNames?: ReadonlySet<string>): OptionExplicitResult {
    const violatedProcedures = new Map<string, Map<string, number>>();

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
            const undeclared = checkProcedure(proc, moduleLevelNames, program.diagnostics, knownModuleNames);
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
                const undeclared = checkProcedure(proc, classModuleNames, program.diagnostics, knownModuleNames);
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
 * Returns the set of undeclared identifiers in this procedure, or null if no violations.
 */
function checkProcedure(
    proc: ProcedureDeclaration,
    moduleLevelNames: Set<string>,
    diagnostics: ParseDiagnostic[],
    knownModuleNames?: ReadonlySet<string>
): Map<string, number> | null {
    // Build the declared-name set for this procedure:
    // module-level names + parameters + function name (for return value)
    const declared = new Set<string>(moduleLevelNames);

    declared.add(proc.name.name.toLowerCase());

    for (const param of proc.parameters) {
        declared.add(param.name.toLowerCase());
    }

    const undeclaredNames = new Map<string, number>();
    const onViol = (name: string, line: number) => {
        if (!undeclaredNames.has(name)) undeclaredNames.set(name, line);
    };

    // Helpers that close over the mutable `declared` set and other context
    const chkExpr = (e: Expression) => checkExpr(e, declared, diagnostics, onViol, knownModuleNames);
    const chkCall = (e: CallExpression) => checkCallExpr(e, declared, diagnostics, onViol, knownModuleNames);
    const chkIdents = (e: Expression) => checkExprIdents(e, declared, diagnostics, onViol);

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
                chkIdents(r.name);
                for (const bound of r.bounds) {
                    if (bound.lower) chkExpr(bound.lower);
                    chkExpr(bound.upper);
                }
                break;
            }
            case 'AssignmentStatement': {
                const a = stmt as AssignmentStatement;
                chkExpr(a.left);
                chkExpr(a.right);
                break;
            }
            case 'SetStatement': {
                const s = stmt as SetStatement;
                chkExpr(s.left);
                chkExpr(s.right);
                break;
            }
            case 'CallStatement': {
                const c = stmt as CallStatement;
                chkCall(c.expression);
                break;
            }
            case 'ForStatement': {
                const f = stmt as ForStatement;
                // For loop variable must be explicitly declared
                chkIdents(f.identifier);
                chkExpr(f.start);
                chkExpr(f.end);
                if (f.step) chkExpr(f.step);
                for (const s of f.body) visitStmt(s);
                break;
            }
            case 'ForEachStatement': {
                const f = stmt as ForEachStatement;
                chkIdents(f.variable);
                chkExpr(f.collection);
                for (const s of f.body) visitStmt(s);
                break;
            }
            case 'IfStatement': {
                const i = stmt as IfStatement;
                chkExpr(i.condition);
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
                if (d.condition) chkExpr(d.condition);
                for (const s of d.body) visitStmt(s);
                break;
            }
            case 'WhileStatement': {
                const w = stmt as WhileStatement;
                chkExpr(w.condition);
                for (const w2 of w.body) visitStmt(w2);
                break;
            }
            case 'WithStatement': {
                const w = stmt as WithStatement;
                chkExpr(w.expression);
                for (const s of w.body) visitStmt(s);
                break;
            }
            case 'SelectCaseStatement': {
                const sc = stmt as SelectCaseStatement;
                chkExpr(sc.expression);
                for (const c of sc.cases) {
                    for (const range of c.ranges) {
                        if (range.kind === 'expression') {
                            chkExpr(range.value);
                        } else if (range.kind === 'to') {
                            chkExpr(range.start);
                            chkExpr(range.end);
                        } else {
                            chkExpr(range.value);
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
                chkExpr(l.left);
                chkExpr(l.right);
                break;
            }
            case 'RSetStatement': {
                const r = stmt as RSetStatement;
                chkExpr(r.left);
                chkExpr(r.right);
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
                break;
        }
    };

    for (const stmt of proc.body) {
        visitStmt(stmt);
    }

    return undeclaredNames.size > 0 ? undeclaredNames : null;
}

type OnViolation = (name: string, line: number) => void;

function exprLine(expr: Expression): number {
    return (expr as any).loc?.start?.line ?? 0;
}

/** Check an expression that is known to be an Identifier (for loop var, etc.) */
function checkExprIdents(expr: Expression, declared: Set<string>, diagnostics: ParseDiagnostic[], onViolation: OnViolation): void {
    if (expr.type === 'Identifier') {
        const id = expr as Identifier;
        if (id.foreign) return; // FOREIGN-NAME [identifier] は Option Explicit チェック対象外 (§3.3.5.2)
        const lower = id.name.toLowerCase();
        if (!declared.has(lower) && !VBA_BUILTINS.has(lower)) {
            reportUndeclared(id.name, expr, diagnostics);
            onViolation(lower, exprLine(expr));
        }
    }
}

/** Recursively check an expression for undeclared Identifiers */
function checkExpr(expr: Expression, declared: Set<string>, diagnostics: ParseDiagnostic[], onViolation: OnViolation, knownModuleNames?: ReadonlySet<string>): void {
    switch (expr.type) {
        case 'Identifier': {
            const id = expr as Identifier;
            if (id.foreign) break; // FOREIGN-NAME [identifier] は Option Explicit チェック対象外
            const lower = id.name.toLowerCase();
            if (!declared.has(lower) && !VBA_BUILTINS.has(lower)) {
                reportUndeclared(id.name, expr, diagnostics);
                onViolation(lower, exprLine(expr));
            }
            break;
        }
        case 'CallExpression': {
            checkCallExpr(expr as CallExpression, declared, diagnostics, onViolation, knownModuleNames);
            break;
        }
        case 'MemberExpression': {
            const m = expr as MemberExpression;
            // Only check the object (left side); property is always a name, not a variable reference
            checkExpr(m.object, declared, diagnostics, onViolation, knownModuleNames);
            break;
        }
        case 'DictionaryAccessExpression': {
            const d = expr as { type: string; object: Expression; property: Identifier };
            checkExpr(d.object, declared, diagnostics, onViolation, knownModuleNames);
            break;
        }
        case 'BinaryExpression': {
            const b = expr as BinaryExpression;
            checkExpr(b.left, declared, diagnostics, onViolation, knownModuleNames);
            checkExpr(b.right, declared, diagnostics, onViolation, knownModuleNames);
            break;
        }
        case 'UnaryExpression': {
            const u = expr as UnaryExpression;
            checkExpr(u.argument, declared, diagnostics, onViolation, knownModuleNames);
            break;
        }
        case 'ParenthesizedExpression': {
            const p = expr as ParenthesizedExpression;
            checkExpr(p.expression, declared, diagnostics, onViolation, knownModuleNames);
            break;
        }
        case 'NamedArgument': {
            const n = expr as NamedArgument;
            checkExpr(n.value, declared, diagnostics, onViolation, knownModuleNames);
            break;
        }
        // Literals, NewExpression, ImplicitWithObjectExpression, etc. — no variable refs
        default:
            break;
    }
}

function checkCallExpr(expr: CallExpression, declared: Set<string>, diagnostics: ParseDiagnostic[], onViolation: OnViolation, knownModuleNames?: ReadonlySet<string>): void {
    // callee: if it's a bare Identifier, it's a function/sub call — NOT a variable reference
    // if it's a MemberExpression, check the object part only
    const callee = expr.callee;
    if (callee.type === 'MemberExpression') {
        const m = callee as MemberExpression;
        if (m.object.type === 'Identifier') {
            const lower = (m.object as Identifier).name.toLowerCase();
            if (knownModuleNames !== undefined) {
                // 2nd pass: knownModuleNames is available — only skip actual module names.
                // Bare identifiers that are not a known module / declared var / built-in are
                // undeclared variables and should be flagged.
                if (!knownModuleNames.has(lower) && !declared.has(lower) && !VBA_BUILTINS.has(lower)) {
                    reportUndeclared((m.object as Identifier).name, m.object, diagnostics);
                    onViolation(lower, exprLine(m.object));
                }
            }
            // 1st pass (knownModuleNames === undefined): skip all bare identifier objects.
            // Cannot distinguish module references from undeclared variables without
            // knowing all loaded module names.
        } else {
            checkExpr(m.object, declared, diagnostics, onViolation, knownModuleNames);
        }
    }
    // Bare Identifier callee = procedure name, skip variable check
    // Check all arguments
    for (const arg of expr.args) {
        checkExpr(arg, declared, diagnostics, onViolation, knownModuleNames);
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

// -----------------------------------------------------------------------
// Undefined procedure call detection (prerun compile error)
// -----------------------------------------------------------------------

export interface UndefinedProcError {
    name: string;
    line: number;
}

/**
 * Walk all procedure bodies and collect bare-Identifier call-expression callees
 * that are not resolvable as a known procedure, built-in, or declared variable.
 * This models VBA's static "Sub or Function not defined" prerun compile error
 * for unqualified calls. Qualified calls (MemberExpression) are dynamic (runtime
 * error 424) and are intentionally excluded.
 *
 * @param program        - Parsed AST for one module.
 * @param knownProcNames - All procedure names from every loaded module (lower-cased).
 */
export function collectUndefinedProcCalls(
    program: Program,
    knownProcNames: ReadonlySet<string>,
): UndefinedProcError[] {
    const errors: UndefinedProcError[] = [];

    const moduleLevelNames = new Set<string>();
    for (const stmt of program.body) {
        collectModuleLevelDeclaredNames(stmt, moduleLevelNames);
    }

    for (const stmt of program.body) {
        if (stmt.type === 'ProcedureDeclaration') {
            walkProcForUndefinedCalls(
                stmt as ProcedureDeclaration,
                moduleLevelNames,
                knownProcNames,
                errors,
            );
        } else if (stmt.type === 'ClassDeclaration') {
            const cls = stmt as ClassDeclaration;
            const classLevel = new Set<string>(moduleLevelNames);
            for (const field of cls.fields) {
                for (const decl of field.declarations) {
                    classLevel.add(decl.name.name.toLowerCase());
                }
            }
            for (const proc of cls.procedures) {
                walkProcForUndefinedCalls(proc, classLevel, knownProcNames, errors);
            }
        }
    }

    return errors;
}

function walkProcForUndefinedCalls(
    proc: ProcedureDeclaration,
    moduleLevelNames: ReadonlySet<string>,
    knownProcNames: ReadonlySet<string>,
    errors: UndefinedProcError[],
): void {
    const declared = new Set<string>(moduleLevelNames);
    declared.add(proc.name.name.toLowerCase());
    for (const param of proc.parameters) {
        declared.add(param.name.toLowerCase());
    }

    const isKnown = (name: string): boolean =>
        declared.has(name) || VBA_BUILTINS.has(name) || knownProcNames.has(name);

    // On Error Resume Next が有効な間は未定義プロシージャ呼び出しを静的エラーにしない
    // (実行時エラーとして On Error で捕捉される意図の可能性があるため)
    let inResumeNext = false;

    const visitCall = (ce: CallExpression): void => {
        if (inResumeNext) return; // On Error Resume Next スコープ内はスキップ
        if (ce.callee.type === 'Identifier') {
            const id = ce.callee as Identifier;
            if (!id.foreign) { // FOREIGN-NAME は未定義プロシージャチェック対象外
                const lower = id.name.toLowerCase();
                if (!isKnown(lower)) {
                    const line = (ce.callee as any).loc?.start?.line ?? 0;
                    errors.push({ name: id.name, line });
                }
            }
        }
        for (const arg of ce.args) visitExpr(arg);
    };

    const visitExpr = (expr: Expression): void => {
        switch (expr.type) {
            case 'CallExpression':
                visitCall(expr as CallExpression);
                break;
            case 'MemberExpression':
                visitExpr((expr as MemberExpression).object);
                break;
            case 'BinaryExpression': {
                const b = expr as BinaryExpression;
                visitExpr(b.left);
                visitExpr(b.right);
                break;
            }
            case 'UnaryExpression':
                visitExpr((expr as UnaryExpression).argument);
                break;
            case 'ParenthesizedExpression':
                visitExpr((expr as ParenthesizedExpression).expression);
                break;
            case 'NamedArgument':
                visitExpr((expr as NamedArgument).value);
                break;
        }
    };

    const visitStmt = (stmt: Statement): void => {
        switch (stmt.type) {
            case 'VariableDeclaration':
                for (const decl of (stmt as VariableDeclaration).declarations) {
                    declared.add(decl.name.name.toLowerCase());
                }
                break;
            case 'ConstDeclaration':
                declared.add((stmt as ConstDeclaration).name.name.toLowerCase());
                break;
            case 'OnErrorStatement': {
                const oe = stmt as { type: string; label: string };
                inResumeNext = oe.label === 'Resume Next';
                break;
            }
            case 'CallStatement':
                visitCall((stmt as CallStatement).expression);
                break;
            case 'AssignmentStatement': {
                const a = stmt as AssignmentStatement;
                visitExpr(a.left);
                visitExpr(a.right);
                break;
            }
            case 'SetStatement': {
                const s = stmt as SetStatement;
                visitExpr(s.left);
                visitExpr(s.right);
                break;
            }
            case 'ForStatement': {
                const f = stmt as ForStatement;
                visitExpr(f.start);
                visitExpr(f.end);
                if (f.step) visitExpr(f.step);
                for (const s of f.body) visitStmt(s);
                break;
            }
            case 'ForEachStatement': {
                const f = stmt as ForEachStatement;
                visitExpr(f.collection);
                for (const s of f.body) visitStmt(s);
                break;
            }
            case 'IfStatement': {
                const i = stmt as IfStatement;
                visitExpr(i.condition);
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
                if (d.condition) visitExpr(d.condition);
                for (const s of d.body) visitStmt(s);
                break;
            }
            case 'WhileStatement': {
                const w = stmt as WhileStatement;
                visitExpr(w.condition);
                for (const s of w.body) visitStmt(s);
                break;
            }
            case 'WithStatement': {
                const w = stmt as WithStatement;
                visitExpr(w.expression);
                for (const s of w.body) visitStmt(s);
                break;
            }
            case 'SelectCaseStatement': {
                const sc = stmt as SelectCaseStatement;
                visitExpr(sc.expression);
                for (const c of sc.cases) {
                    for (const range of c.ranges) {
                        if (range.kind === 'to') {
                            visitExpr(range.start);
                            visitExpr(range.end);
                        } else {
                            visitExpr(range.value);
                        }
                    }
                    for (const s of c.body) visitStmt(s);
                }
                if (sc.elseBody) {
                    for (const s of sc.elseBody) visitStmt(s);
                }
                break;
            }
            case 'ReDimStatement': {
                const r = stmt as ReDimStatement;
                for (const bound of r.bounds) {
                    if (bound.lower) visitExpr(bound.lower);
                    visitExpr(bound.upper);
                }
                break;
            }
            case 'LSetStatement': {
                const l = stmt as LSetStatement;
                visitExpr(l.left);
                visitExpr(l.right);
                break;
            }
            case 'RSetStatement': {
                const r = stmt as RSetStatement;
                visitExpr(r.left);
                visitExpr(r.right);
                break;
            }
        }
    };

    for (const stmt of proc.body) visitStmt(stmt);
}
