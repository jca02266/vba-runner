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
 *   VBA009 - デッドストア → 代入値が上書きまたは未使用
 *   VBA010 - 到達不能コード → Exit Sub / GoTo などで実行されないコード
 *   VBA011 - Range 変数経由の Excel プロパティ/メソッドアクセス（Excel依存箇所の可視化）
 *   VBA012 - ByRef 明示なしのパラメーターへ代入 → 呼び出し元を意図せず書き換える罠
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
    LabelStatement,
    ClassDeclaration,
    CallExpression,
    Identifier,
    NumberLiteral,
    MemberExpression,
} from './parser';
import { findDeadStores } from './dead-store';
import { buildCFG, findUnreachableBlocks } from './cfg';

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

// ─── ループ continue ラベル収集 ───────────────────────────────────────────────

/**
 * ループ本体（For/ForEach/DoWhile/While）の末尾にある LabelStatement を収集する。
 * そのラベルへの GoTo はループの continue として使われているとみなす。
 */
export function findLoopContinueLabels(stmts: Statement[]): Set<string> {
    const result = new Set<string>();
    for (const stmt of stmts) {
        switch (stmt.type) {
            case 'ForStatement':
            case 'ForEachStatement':
            case 'DoWhileStatement':
            case 'WhileStatement': {
                const body: Statement[] = (stmt as any).body ?? [];
                if (body.length > 0 && body[body.length - 1].type === 'LabelStatement') {
                    result.add((body[body.length - 1] as LabelStatement).label.toLowerCase());
                }
                for (const l of findLoopContinueLabels(body)) result.add(l);
                break;
            }
            case 'IfStatement': {
                const is = stmt as any;
                const cons = Array.isArray(is.consequent) ? is.consequent : (is.consequent ? [is.consequent] : []);
                const alt  = Array.isArray(is.alternate)  ? is.alternate  : (is.alternate  ? [is.alternate]  : []);
                for (const l of findLoopContinueLabels(cons)) result.add(l);
                for (const l of findLoopContinueLabels(alt))  result.add(l);
                break;
            }
            case 'WithStatement':
                for (const l of findLoopContinueLabels((stmt as any).body ?? [])) result.add(l);
                break;
        }
    }
    return result;
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

function lintStatement(stmt: Statement, out: LintDiagnostic[], continueLabels: Set<string> = new Set()): void {
    switch (stmt.type) {
        case 'VariableDeclaration':
            checkDimMultiDecl(stmt as VariableDeclaration, out);
            checkIntegerType(stmt as VariableDeclaration, out);
            break;

        case 'ProcedureDeclaration': {
            const proc = stmt as ProcedureDeclaration;
            checkParameters(proc, out);
            checkParamAssignWithoutByRef(proc, out);
            checkDeadStores(proc, out);
            checkUnreachableCode(proc, out);
            const procContinueLabels = findLoopContinueLabels(proc.body);
            for (const s of proc.body) lintStatement(s, out, procContinueLabels);
            break;
        }

        case 'ClassDeclaration': {
            const cls = stmt as ClassDeclaration;
            for (const field of cls.fields) lintStatement(field, out, continueLabels);
            for (const proc of cls.procedures) lintStatement(proc, out, continueLabels);
            break;
        }

        case 'WhileStatement':
            checkWhileWend(stmt as WhileStatement, out);
            for (const s of (stmt as any).body ?? []) lintStatement(s, out, continueLabels);
            break;

        case 'SelectCaseStatement':
            checkSelectCaseElse(stmt as SelectCaseStatement, out);
            for (const clause of (stmt as SelectCaseStatement).cases ?? []) {
                for (const s of clause.body ?? []) lintStatement(s, out, continueLabels);
            }
            if ((stmt as SelectCaseStatement).elseBody) {
                for (const s of (stmt as SelectCaseStatement).elseBody!) lintStatement(s, out, continueLabels);
            }
            break;

        case 'GoToStatement':
            checkGoTo(stmt as GoToStatement, out, continueLabels);
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
            for (const s of cons) lintStatement(s, out, continueLabels);
            for (const s of alt)  lintStatement(s, out, continueLabels);
            break;
        }

        case 'ForStatement':
        case 'ForEachStatement':
        case 'DoWhileStatement':
        case 'WithStatement': {
            const bs = stmt as any;
            for (const s of (bs.body ?? [])) lintStatement(s, out, continueLabels);
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

/**
 * 文の列を走査し、代入先となっている識別子名（小文字）を収集する。
 * 対象: 単純代入 / Set 代入の左辺 Identifier、For/For Each のループ変数。
 * ネストしたブロック（If/For/While/Do/With/Select Case）も再帰的にたどる。
 */
function collectAssignedNames(stmts: Statement[], out: Set<string> = new Set()): Set<string> {
    for (const stmt of stmts) {
        switch (stmt.type) {
            case 'AssignmentStatement':
            case 'SetStatement': {
                const left = (stmt as any).left;
                if (left && left.type === 'Identifier') out.add((left as Identifier).name.toLowerCase());
                break;
            }
            case 'ForStatement': {
                const id = (stmt as any).identifier;
                if (id && id.type === 'Identifier') out.add((id as Identifier).name.toLowerCase());
                collectAssignedNames((stmt as any).body ?? [], out);
                break;
            }
            case 'ForEachStatement': {
                const v = (stmt as any).variable;
                if (v && v.type === 'Identifier') out.add((v as Identifier).name.toLowerCase());
                collectAssignedNames((stmt as any).body ?? [], out);
                break;
            }
            case 'DoWhileStatement':
            case 'WhileStatement':
            case 'WithStatement':
                collectAssignedNames((stmt as any).body ?? [], out);
                break;
            case 'IfStatement': {
                const is = stmt as any;
                const cons = Array.isArray(is.consequent) ? is.consequent : (is.consequent ? [is.consequent] : []);
                const alt  = Array.isArray(is.alternate)  ? is.alternate  : (is.alternate  ? [is.alternate]  : []);
                collectAssignedNames(cons, out);
                collectAssignedNames(alt, out);
                break;
            }
            case 'SelectCaseStatement': {
                const sc = stmt as SelectCaseStatement;
                for (const clause of sc.cases ?? []) collectAssignedNames(clause.body ?? [], out);
                if (sc.elseBody) collectAssignedNames(sc.elseBody, out);
                break;
            }
        }
    }
    return out;
}

/**
 * VBA012: パラメーターに ByRef の指定がないのに代入している
 *
 * 修飾子なし（暗黙の ByRef）のパラメーターに本体内で代入していると、
 * 呼び出し元の変数を意図せず書き換えてしまう。明示的に ByRef（書き戻す）
 * または ByVal（ローカルコピー）を指定するよう促す。
 * ByVal/ByRef が明示されている場合は対象外。
 */
function checkParamAssignWithoutByRef(proc: ProcedureDeclaration, out: LintDiagnostic[]): void {
    const assigned = collectAssignedNames(proc.body);

    for (const param of proc.parameters) {
        // ByVal / ByRef が明示されているものは対象外
        if (param.hasPassingModifier) continue;
        // ParamArray は代入対象として扱わない
        if (param.isParamArray) continue;
        if (!assigned.has(param.name.toLowerCase())) continue;

        const loc  = (param as any).loc;
        const line = (loc?.start.line   ?? 1) - 1;
        const col  = (loc?.start.column ?? 1) - 1;
        out.push({
            code: 'VBA012',
            severity: 2,
            message: `パラメーター '${param.name}' は ByRef の指定がないのに代入されています。呼び出し元の変数を意図せず書き換える可能性があります（明示的に ByRef または ByVal を指定してください）`,
            line, column: col, endLine: line, endColumn: col + param.name.length,
        });
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

/** VBA008: GoTo（エラーハンドラー・ループcontinue以外） */
function checkGoTo(stmt: GoToStatement, out: LintDiagnostic[], continueLabels: Set<string> = new Set()): void {
    const label = (stmt as any).label ?? '';
    const lower = label.toLowerCase();
    // エラーハンドラーらしいラベル名は除外
    if (lower.includes('err') || lower.includes('exit') || lower.includes('clean') || lower.includes('finally')) return;
    // ループ末尾に定義されたラベルはcontinueとして除外
    if (continueLabels.has(lower)) return;

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

/** VBA010: 到達不能コード（Exit Sub / GoTo などで実行されないコード） */
function checkUnreachableCode(proc: ProcedureDeclaration, out: LintDiagnostic[]): void {
    try {
        const cfg = buildCFG(proc);
        for (const block of findUnreachableBlocks(cfg)) {
            const first = block.stmts[0];
            const last  = block.stmts[block.stmts.length - 1];
            const startLoc = (first as any).loc?.start;
            const endLoc   = (last  as any).loc?.end ?? (last as any).loc?.start;
            if (!startLoc) continue;
            out.push({
                code: 'VBA010',
                severity: 3,
                message: 'このコードには到達できません（到達不能コード）',
                line:      startLoc.line   - 1,
                column:    startLoc.column - 1,
                endLine:   (endLoc?.line   ?? startLoc.line)   - 1,
                endColumn: (endLoc?.column ?? startLoc.column) - 1,
            });
        }
    } catch {
        // CFG 構築失敗は lint を止めない
    }
}

/** VBA009: デッドストア（代入値が上書きまたは未使用） */
function checkDeadStores(proc: ProcedureDeclaration, out: LintDiagnostic[]): void {
    let deadStores;
    try {
        deadStores = findDeadStores(proc);
    } catch {
        return; // CFG/解析エラーは lint を止めない
    }

    for (const ds of deadStores) {
        // 対象: 単純代入のみ（For/ForEach ループカウンターは除外）
        if (ds.stmtType !== 'AssignmentStatement' && ds.stmtType !== 'SetStatement') continue;

        const line      = ds.line      > 0 ? ds.line      - 1 : 0;
        const col       = ds.column    > 0 ? ds.column    - 1 : 0;
        const endCol    = ds.endColumn > 0 ? ds.endColumn - 1 : col + ds.varName.length;
        out.push({
            code: 'VBA009',
            severity: 2,
            message: `'${ds.varName}' への代入はデッドストア（使用されない代入値）`,
            line, column: col, endLine: line, endColumn: endCol,
        });
    }
}
