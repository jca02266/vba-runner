/**
 * VBA 固有の Diagnostics 警告
 *
 * VBA 初心者・移行者が踏みやすい言語仕様の罠を静的解析で検出する。
 *
 * 重大度（LSP severity）3段階:
 *   1 = 重要 (Error)   — バグや実行時エラーに直結する問題 / VS Code: 赤いエラーアイコン
 *   2 = 警告 (Warning) — 意図しない動作を引き起こしやすいコード / VS Code: 黄色い警告アイコン
 *   3 = 情報 (Info)    — スタイル・保守性の改善提案 / VS Code: 青い情報アイコン
 *
 * ルール一覧:
 *   [重要] VBA013 - Option Explicit なし → 変数名のタイポが実行時まで検出されない
 *   [警告] VBA001 - Dim 複数宣言で途中の変数に型指定なし → Variant になる罠
 *   [警告] VBA003 - パラメーターに ByVal/ByRef なし → デフォルト ByRef の罠
 *   [警告] VBA005 - Select Case に Case Else なし → 想定外の値をサイレントスルー
 *   [警告] VBA006 - Sheets(n) / Worksheets(n) 数値インデックス → シート順変更で壊れる
 *   [警告] VBA008 - GoTo（エラーハンドラー以外） → スパゲッティ化の原因
 *   [警告] VBA009 - デッドストア → 代入値が上書きまたは未使用
 *   [警告] VBA010 - 到達不能コード → Exit Sub / GoTo などで実行されないコード
 *   [警告] VBA012 - ByRef 明示なしのパラメーターへ代入 → 呼び出し元を意図せず書き換える罠
 *   [情報] VBA002 - Integer 型の使用 → Long を推奨
 *   [情報] VBA004 - While...Wend → Do While...Loop を推奨
 *   [情報] VBA007 - ActiveSheet / ActiveWorkbook 直接参照 → 何が選択されているか依存
 *   [情報] VBA011 - Range 変数経由の Excel プロパティ/メソッドアクセス（Excel依存箇所の可視化）
 *   [警告] VBA014 - 未使用ローカル変数（宣言後に一切参照されない）
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
    OptionExplicitStatement,
} from './parser';
import { findDeadStores } from './dead-store';
import { buildCFG, findUnreachableBlocks } from './cfg';

// ─── 公開型 ──────────────────────────────────────────────────────────────────

export interface LintDiagnostic {
    /** ルール ID */
    code: string;
    /** LSP severity: 1=Error 2=Warning 3=Information 4=Hint */
    severity: 1 | 2 | 3 | 4;
    /** English message (pre-formatted, used in non-extension contexts) */
    message: string;
    /** l10n key template (e.g. "Parameter '{0}' has no explicit ByVal/ByRef") */
    l10nKey: string;
    /** Substitution args for {0}, {1}, ... */
    l10nArgs: string[];
    /** 0-based */
    line: number;
    /** 0-based */
    column: number;
    endLine: number;
    endColumn: number;
}

/** {0}, {1}, … を args で置換して英語プレビュー文字列を生成する */
function fmt(key: string, args: string[]): string {
    return args.reduce((s, a, i) => s.split(`{${i}}`).join(a), key);
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

    checkOptionExplicit(program, diags);

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
            checkUnusedVariables(proc, out);
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
            const k001 = "'{0}' has no type specifier and becomes Variant. Specify a type for each variable (e.g., Dim {0} As Long)";
            const a001 = [decl.name.name, decl.name.name];
            out.push({
                code: 'VBA001', severity: 2,
                message: fmt(k001, a001), l10nKey: k001, l10nArgs: a001,
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
            const k002v = "'{0}' is declared as Integer. In VBA, Long is faster and recommended";
            const a002v = [decl.name.name];
            out.push({
                code: 'VBA002', severity: 3,
                message: fmt(k002v, a002v), l10nKey: k002v, l10nArgs: a002v,
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
        const k002p = "Parameter '{0}' is declared as Integer. Long is recommended";
        const a002p = [param.name];
        out.push({
            code: 'VBA002', severity: 3,
            message: fmt(k002p, a002p), l10nKey: k002p, l10nArgs: a002p,
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
            const k003 = "Parameter '{0}' has no explicit ByVal/ByRef. VBA defaults to ByRef";
            const a003 = [param.name];
            out.push({
                code: 'VBA003', severity: 2,
                message: fmt(k003, a003), l10nKey: k003, l10nArgs: a003,
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

        const k012 = "Parameter '{0}' is assigned without explicit ByRef, which may unintentionally modify the caller's variable (specify ByRef or ByVal explicitly)";
        const a012 = [param.name];
        const loc  = (param as any).loc;
        const line = (loc?.start.line   ?? 1) - 1;
        const col  = (loc?.start.column ?? 1) - 1;
        out.push({
            code: 'VBA012', severity: 2,
            message: fmt(k012, a012), l10nKey: k012, l10nArgs: a012,
            line, column: col, endLine: line, endColumn: col + param.name.length,
        });
    }
}

/** VBA004: While...Wend → Do While...Loop を推奨 */
function checkWhileWend(stmt: WhileStatement, out: LintDiagnostic[]): void {
    const loc  = (stmt as any).loc;
    const line = (loc?.start.line   ?? 1) - 1;
    const col  = (loc?.start.column ?? 1) - 1;
    const k004 = "'While...Wend' is obsolete. Use 'Do While...Loop' instead";
    out.push({
        code: 'VBA004', severity: 3,
        message: k004, l10nKey: k004, l10nArgs: [],
        line, column: col, endLine: line, endColumn: col + 5,
    });
}

/** VBA005: Select Case に Case Else なし */
function checkSelectCaseElse(sc: SelectCaseStatement, out: LintDiagnostic[]): void {
    if (!sc.elseBody || sc.elseBody.length === 0) {
        const k005 = "'Select Case' has no 'Case Else'. Unexpected values will be silently ignored";
        const loc  = (sc as any).loc;
        const line = (loc?.start.line   ?? 1) - 1;
        const col  = (loc?.start.column ?? 1) - 1;
        out.push({
            code: 'VBA005', severity: 2,
            message: k005, l10nKey: k005, l10nArgs: [],
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

    const k006 = "Numeric index ({0}) for Sheets/Worksheets breaks when sheet order changes. Use a sheet name instead";
    const a006 = [String((arg0 as NumberLiteral).value)];
    const loc  = ce.callee.loc;
    const line = (loc?.start.line   ?? 1) - 1;
    const col  = (loc?.start.column ?? 1) - 1;
    out.push({
        code: 'VBA006', severity: 2,
        message: fmt(k006, a006), l10nKey: k006, l10nArgs: a006,
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
    const k007 = "'{0}' depends on the currently selected object. Use an explicit reference instead";
    const a007 = [id.name];
    out.push({
        code: 'VBA007', severity: 3,
        message: fmt(k007, a007), l10nKey: k007, l10nArgs: a007,
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

    const k008 = "'GoTo {0}' leads to spaghetti code. Consider using loops or conditionals instead";
    const a008 = [label];
    const loc  = (stmt as any).loc;
    const line = (loc?.start.line   ?? 1) - 1;
    const col  = (loc?.start.column ?? 1) - 1;
    out.push({
        code: 'VBA008', severity: 2,
        message: fmt(k008, a008), l10nKey: k008, l10nArgs: a008,
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
                severity: 2,
                message: 'This code is unreachable', l10nKey: 'This code is unreachable', l10nArgs: [],
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

        const k009 = "Assignment to '{0}' is a dead store (value is never used)";
        const a009 = [ds.varName];
        const line      = ds.line      > 0 ? ds.line      - 1 : 0;
        const col       = ds.column    > 0 ? ds.column    - 1 : 0;
        const endCol    = ds.endColumn > 0 ? ds.endColumn - 1 : col + ds.varName.length;
        out.push({
            code: 'VBA009', severity: 2,
            message: fmt(k009, a009), l10nKey: k009, l10nArgs: a009,
            line, column: col, endLine: line, endColumn: endCol,
        });
    }
}

/** VBA014: 未使用ローカル変数（宣言後に一切参照されない） */
function checkUnusedVariables(proc: ProcedureDeclaration, out: LintDiagnostic[]): void {
    type DeclInfo = { name: string; loc: any };
    const locals = new Map<string, DeclInfo>();

    // 手続き本体内の Dim/Const/Static 宣言を収集（scope なし = ローカル）
    const gatherDecls = (stmts: Statement[]) => {
        for (const stmt of stmts) {
            if (stmt.type === 'VariableDeclaration' && !(stmt as VariableDeclaration).scope) {
                for (const d of (stmt as VariableDeclaration).declarations) {
                    const lower = d.name.name.toLowerCase();
                    if (!locals.has(lower)) locals.set(lower, { name: d.name.name, loc: d.name.loc });
                }
            }
            recurseBlock(stmt, gatherDecls);
        }
    };
    gatherDecls(proc.body);
    if (locals.size === 0) return;

    // 本体中に現れる識別子をすべて収集（読み取り・代入いずれも「参照」とみなす）
    const mentioned = new Set<string>();

    const scanExpr = (e: Expression | null | undefined): void => {
        if (!e) return;
        switch (e.type) {
            case 'Identifier':
                mentioned.add((e as Identifier).name.toLowerCase());
                break;
            case 'MemberExpression':
                scanExpr((e as MemberExpression).object);
                break;
            case 'CallExpression': {
                const ce = e as CallExpression;
                scanExpr(ce.callee);
                ce.args.forEach(scanExpr);
                break;
            }
            case 'BinaryExpression':
            case 'LogicalExpression': {
                const be = e as any;
                scanExpr(be.left); scanExpr(be.right);
                break;
            }
            case 'UnaryExpression':
                scanExpr((e as any).argument);
                break;
            case 'ParenthesizedExpression':
                scanExpr((e as any).expression);
                break;
        }
    };

    const scanStmt = (s: Statement): void => {
        const a = s as any;
        switch (s.type) {
            case 'AssignmentStatement':
            case 'SetStatement':
                // 単純代入の左辺識別子も「参照あり」として扱う（dead store は VBA009 が担当）
                if (a.left?.type === 'Identifier') mentioned.add(a.left.name.toLowerCase());
                else scanExpr(a.left);
                scanExpr(a.right);
                break;
            case 'CallStatement': scanExpr(a.expression); break;
            case 'IfStatement':
                scanExpr(a.condition);
                (Array.isArray(a.consequent) ? a.consequent : a.consequent ? [a.consequent] : []).forEach(scanStmt);
                (Array.isArray(a.alternate)  ? a.alternate  : a.alternate  ? [a.alternate]  : []).forEach(scanStmt);
                break;
            case 'ForStatement':
                if (a.identifier?.type === 'Identifier') mentioned.add(a.identifier.name.toLowerCase());
                scanExpr(a.start); scanExpr(a.end); scanExpr(a.step);
                (a.body ?? []).forEach(scanStmt);
                break;
            case 'ForEachStatement':
                if (a.variable?.type === 'Identifier') mentioned.add(a.variable.name.toLowerCase());
                scanExpr(a.iterable);
                (a.body ?? []).forEach(scanStmt);
                break;
            case 'WhileStatement':
            case 'DoWhileStatement':
                scanExpr(a.condition); (a.body ?? []).forEach(scanStmt); break;
            case 'WithStatement':
                scanExpr(a.object); (a.body ?? []).forEach(scanStmt); break;
            case 'SelectCaseStatement':
                scanExpr(a.discriminant);
                for (const clause of a.cases ?? []) {
                    (clause.tests ?? []).forEach(scanExpr);
                    (clause.body ?? []).forEach(scanStmt);
                }
                (a.elseBody ?? []).forEach(scanStmt);
                break;
            case 'ReturnStatement': scanExpr(a.value); break;
        }
    };

    proc.body.forEach(scanStmt);

    // 宣言されたが一度も言及されていない変数を報告
    for (const [lower, info] of locals) {
        if (mentioned.has(lower)) continue;
        const loc  = info.loc;
        const line = (loc?.start.line   ?? 1) - 1;
        const col  = (loc?.start.column ?? 1) - 1;
        const k014 = "Variable '{0}' is declared but never used";
        const a014 = [info.name];
        out.push({
            code: 'VBA014', severity: 2,
            message: fmt(k014, a014), l10nKey: k014, l10nArgs: a014,
            line, column: col, endLine: line, endColumn: col + info.name.length,
        });
    }
}

/** ブロック文の子 body を再帰的に走査するユーティリティ */
function recurseBlock(stmt: Statement, visitor: (stmts: Statement[]) => void): void {
    const a = stmt as any;
    if (a.body)       visitor(a.body);
    if (a.consequent) visitor(Array.isArray(a.consequent) ? a.consequent : [a.consequent]);
    if (a.alternate)  visitor(Array.isArray(a.alternate)  ? a.alternate  : [a.alternate]);
    if (a.cases)      for (const c of a.cases) visitor(c.body ?? []);
    if (a.elseBody)   visitor(a.elseBody);
}

/** VBA013: Option Explicit なし → 変数名のタイポが実行時まで検出されない */
function checkOptionExplicit(program: Program, out: LintDiagnostic[]): void {
    const hasExplicit = program.body.some(
        (s): s is OptionExplicitStatement => s.type === 'OptionExplicitStatement'
    );
    if (!hasExplicit) {
        out.push({
            code: 'VBA013',
            severity: 1,
            message: 'Option Explicit is missing. Typos in variable names will not be detected until runtime',
            l10nKey: 'Option Explicit is missing. Typos in variable names will not be detected until runtime', l10nArgs: [],
            line: 0, column: 0, endLine: 0, endColumn: 0,
        });
    }
}
