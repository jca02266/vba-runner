/**
 * VBA 固有の Diagnostics 警告
 *
 * VBA 初心者・移行者が踏みやすい言語仕様の罠を静的解析で検出する。
 *
 * ルール一覧:
 *   VBA001 - Dim 複数宣言で途中の変数に型指定なし → Variant になる罠
 *   VBA002 - Integer 型の使用 → Long を推奨
 *   VBA003 - パラメーターに ByVal/ByRef なし → デフォルト ByRef の罠
 *   VBA004 - While...Wend → Do While...Loop を推奨
 *   VBA005 - Select Case に Case Else なし → 想定外の値をサイレントスルー
 *   VBA006 - Sheets(n) / Worksheets(n) 数値インデックス → シート順変更で壊れる
 *   VBA007 - ActiveSheet / ActiveWorkbook 直接参照 → 何が選択されているか依存
 *   VBA008 - GoTo（エラーハンドラー以外） → スパゲッティ化の原因
 */

import {
    Program,
    Statement,
    Expression,
    ProcedureDeclaration,
    VariableDeclaration,
    Parameter,
    WhileStatement,
    SelectCaseStatement,
    GoToStatement,
    ClassDeclaration,
    CallExpression,
    Identifier,
    NumberLiteral,
    MemberExpression,
} from './parser';

// ─── 公開型 ──────────────────────────────────────────────────────────────────

export interface LintDiagnostic {
    /** ルール ID */
    code: string;
    /** LSP severity: 1=Error 2=Warning 3=Information 4=Hint */
    severity: 1 | 2 | 3 | 4;
    message: string;
    /** 0-based */
    line: number;
    /** 0-based */
    column: number;
    endLine: number;
    endColumn: number;
}

// ─── エントリーポイント ───────────────────────────────────────────────────────

export function lintProgram(program: Program): LintDiagnostic[] {
    const diags: LintDiagnostic[] = [];

    for (const stmt of program.body) {
        lintStatement(stmt, diags);
    }

    return diags;
}

// ─── ステートメント走査 ───────────────────────────────────────────────────────

function lintStatement(stmt: Statement, out: LintDiagnostic[]): void {
    switch (stmt.type) {
        case 'VariableDeclaration':
            checkDimMultiDecl(stmt as VariableDeclaration, out);
            checkIntegerType(stmt as VariableDeclaration, out);
            break;

        case 'ProcedureDeclaration': {
            const proc = stmt as ProcedureDeclaration;
            checkParameters(proc, out);
            for (const s of proc.body) lintStatement(s, out);
            break;
        }

        case 'ClassDeclaration': {
            const cls = stmt as ClassDeclaration;
            for (const field of cls.fields) lintStatement(field, out);
            for (const proc of cls.procedures) lintStatement(proc, out);
            break;
        }

        case 'WhileStatement':
            checkWhileWend(stmt as WhileStatement, out);
            for (const s of (stmt as any).body ?? []) lintStatement(s, out);
            break;

        case 'SelectCaseStatement':
            checkSelectCaseElse(stmt as SelectCaseStatement, out);
            for (const clause of (stmt as SelectCaseStatement).cases ?? []) {
                for (const s of clause.body ?? []) lintStatement(s, out);
            }
            if ((stmt as SelectCaseStatement).elseBody) {
                for (const s of (stmt as SelectCaseStatement).elseBody!) lintStatement(s, out);
            }
            break;

        case 'GoToStatement':
            checkGoTo(stmt as GoToStatement, out);
            break;

        case 'CallStatement':
            lintExpr((stmt as any).expression, out);
            break;

        case 'AssignmentStatement':
        case 'SetStatement': {
            const s = stmt as any;
            lintExpr(s.left, out);
            lintExpr(s.right, out);
            break;
        }

        case 'IfStatement': {
            const is = stmt as any;
            lintExpr(is.condition, out);
            const cons = Array.isArray(is.consequent) ? is.consequent : (is.consequent ? [is.consequent] : []);
            const alt  = Array.isArray(is.alternate)  ? is.alternate  : (is.alternate  ? [is.alternate]  : []);
            for (const s of cons) lintStatement(s, out);
            for (const s of alt)  lintStatement(s, out);
            break;
        }

        case 'ForStatement':
        case 'ForEachStatement':
        case 'DoWhileStatement':
        case 'WithStatement': {
            const bs = stmt as any;
            for (const s of (bs.body ?? [])) lintStatement(s, out);
            break;
        }
    }
}

function lintExpr(expr: Expression, out: LintDiagnostic[]): void {
    if (!expr) return;

    switch (expr.type) {
        case 'CallExpression':
            checkSheetsNumericIndex(expr as CallExpression, out);
            lintExpr((expr as CallExpression).callee, out);
            for (const arg of (expr as CallExpression).args) lintExpr(arg, out);
            break;

        case 'Identifier':
            checkActiveObject(expr as Identifier, out);
            break;

        case 'MemberExpression': {
            const me = expr as MemberExpression;
            lintExpr(me.object, out);
            break;
        }

        case 'BinaryExpression':
        case 'LogicalExpression': {
            const be = expr as any;
            lintExpr(be.left, out);
            lintExpr(be.right, out);
            break;
        }

        case 'UnaryExpression':
            lintExpr((expr as any).argument, out);
            break;

        case 'ParenthesizedExpression':
            lintExpr((expr as any).expression, out);
            break;
    }
}

// ─── 各ルール ─────────────────────────────────────────────────────────────────

/** VBA001: Dim a, b As Long — 途中の変数に型指定なし → Variant */
function checkDimMultiDecl(vd: VariableDeclaration, out: LintDiagnostic[]): void {
    if (vd.declarations.length < 2) return;

    // 末尾以外の宣言子で型指定なし（Variant）を検出
    for (let i = 0; i < vd.declarations.length - 1; i++) {
        const decl = vd.declarations[i];
        if (!decl.objectType || decl.objectType === 'Variant') {
            const loc = decl.name.loc;
            const line   = (loc?.start.line   ?? 1) - 1;
            const col    = (loc?.start.column ?? 1) - 1;
            const endCol = (loc?.end?.column  ?? col + decl.name.name.length + 1) - 1;
            out.push({
                code: 'VBA001',
                severity: 2,
                message: `'${decl.name.name}' は型指定がなく Variant になります。各変数に型を明示してください（例: Dim ${decl.name.name} As Long）`,
                line, column: col, endLine: line, endColumn: endCol,
            });
        }
    }
}

/** VBA002: Integer 型の使用 → Long を推奨 */
function checkIntegerType(vd: VariableDeclaration, out: LintDiagnostic[]): void {
    for (const decl of vd.declarations) {
        if (decl.objectType?.toLowerCase() === 'integer') {
            const loc  = decl.name.loc;
            const line = (loc?.start.line   ?? 1) - 1;
            const col  = (loc?.start.column ?? 1) - 1;
            out.push({
                code: 'VBA002',
                severity: 3,
                message: `'${decl.name.name}' は Integer 型です。VBA では Long の方が高速で推奨されます`,
                line, column: col, endLine: line, endColumn: col + decl.name.name.length,
            });
        }
    }
}

/** VBA002 パラメーター版 */
function checkIntegerTypeParam(param: Parameter, out: LintDiagnostic[]): void {
    if (param.paramType?.toLowerCase() === 'integer') {
        const loc  = (param as any).loc;
        const line = (loc?.start.line   ?? 1) - 1;
        const col  = (loc?.start.column ?? 1) - 1;
        out.push({
            code: 'VBA002',
            severity: 3,
            message: `パラメーター '${param.name}' は Integer 型です。Long を推奨します`,
            line, column: col, endLine: line, endColumn: col + param.name.length,
        });
    }
}

/** VBA003: パラメーターに ByVal/ByRef なし → デフォルトは ByRef */
function checkParameters(proc: ProcedureDeclaration, out: LintDiagnostic[]): void {
    for (const param of proc.parameters) {
        checkIntegerTypeParam(param, out);

        if (!param.hasPassingModifier && !param.isParamArray && !param.isOptional) {
            const loc  = (param as any).loc;
            const line = (loc?.start.line   ?? 1) - 1;
            const col  = (loc?.start.column ?? 1) - 1;
            out.push({
                code: 'VBA003',
                severity: 2,
                message: `パラメーター '${param.name}' に ByVal/ByRef が明示されていません。VBA のデフォルトは ByRef です`,
                line, column: col, endLine: line, endColumn: col + param.name.length,
            });
        }
    }
}

/** VBA004: While...Wend → Do While...Loop を推奨 */
function checkWhileWend(stmt: WhileStatement, out: LintDiagnostic[]): void {
    const loc  = (stmt as any).loc;
    const line = (loc?.start.line   ?? 1) - 1;
    const col  = (loc?.start.column ?? 1) - 1;
    out.push({
        code: 'VBA004',
        severity: 3,
        message: '\'While...Wend\' は古い構文です。\'Do While...Loop\' を推奨します',
        line, column: col, endLine: line, endColumn: col + 5,
    });
}

/** VBA005: Select Case に Case Else なし */
function checkSelectCaseElse(sc: SelectCaseStatement, out: LintDiagnostic[]): void {
    if (!sc.elseBody || sc.elseBody.length === 0) {
        const loc  = (sc as any).loc;
        const line = (loc?.start.line   ?? 1) - 1;
        const col  = (loc?.start.column ?? 1) - 1;
        out.push({
            code: 'VBA005',
            severity: 3,
            message: '\'Select Case\' に \'Case Else\' がありません。想定外の値がサイレントにスルーされます',
            line, column: col, endLine: line, endColumn: col + 11,
        });
    }
}

/** VBA006: Sheets(1) / Worksheets(1) 数値インデックス */
function checkSheetsNumericIndex(ce: CallExpression, out: LintDiagnostic[]): void {
    if (ce.callee.type !== 'Identifier') return;
    const name = (ce.callee as Identifier).name.toLowerCase();
    if (name !== 'sheets' && name !== 'worksheets') return;
    if (ce.args.length === 0) return;

    const arg0 = ce.args[0];
    if (arg0.type !== 'NumberLiteral') return;

    const loc  = ce.callee.loc;
    const line = (loc?.start.line   ?? 1) - 1;
    const col  = (loc?.start.column ?? 1) - 1;
    out.push({
        code: 'VBA006',
        severity: 2,
        message: `Sheets/Worksheets の数値インデックス (${(arg0 as NumberLiteral).value}) はシート順変更で壊れます。シート名（文字列）での指定を推奨します`,
        line, column: col, endLine: line, endColumn: col + (ce.callee as Identifier).name.length,
    });
}

/** VBA007: ActiveSheet / ActiveWorkbook 直接参照 */
function checkActiveObject(id: Identifier, out: LintDiagnostic[]): void {
    const lower = id.name.toLowerCase();
    if (lower !== 'activesheet' && lower !== 'activeworkbook') return;

    const loc  = id.loc;
    const line = (loc?.start.line   ?? 1) - 1;
    const col  = (loc?.start.column ?? 1) - 1;
    out.push({
        code: 'VBA007',
        severity: 3,
        message: `'${id.name}' は現在選択されているオブジェクトに依存します。明示的な参照（変数またはシート名）を推奨します`,
        line, column: col, endLine: line, endColumn: col + id.name.length,
    });
}

/** VBA008: GoTo（エラーハンドラーラベル以外） */
function checkGoTo(stmt: GoToStatement, out: LintDiagnostic[]): void {
    const label = (stmt as any).label ?? '';
    const lower = label.toLowerCase();
    // エラーハンドラーらしいラベル名は除外
    if (lower.includes('err') || lower.includes('exit') || lower.includes('clean') || lower.includes('finally')) return;

    const loc  = (stmt as any).loc;
    const line = (loc?.start.line   ?? 1) - 1;
    const col  = (loc?.start.column ?? 1) - 1;
    out.push({
        code: 'VBA008',
        severity: 2,
        message: `'GoTo ${label}' はスパゲッティコードの原因になります。ループや条件分岐での代替を検討してください`,
        line, column: col, endLine: line, endColumn: col + 4,
    });
}
