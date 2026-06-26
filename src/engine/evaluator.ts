import {
    Program,
    Statement,
    ForStatement,
    IfStatement,
    DoWhileStatement,
    WhileStatement,
    Expression,
    UnaryExpression,
    BinaryExpression,
    Identifier,
    NumberLiteral,
    StringLiteral,
    DateLiteral,
    AssignmentStatement,
    ProcedureDeclaration,
    VariableDeclaration,
    CallStatement,
    CallExpression,
    MemberExpression,
    ParenthesizedExpression,
    ConstDeclaration,
    SetStatement,
    EraseStatement,
    ReDimStatement,
    ArrayBound,
    ExitStatement,
    OnErrorStatement,
    TypeDeclaration,
    TypeMember,
    SelectCaseStatement,
    ForEachStatement,
    WithStatement,
    ImplicitWithObjectExpression,
    GoToStatement,
    StopStatement,
    EndStatement,
    GoSubStatement,
    ReturnStatement,
    OnGoToSubStatement,
    LSetStatement,
    RSetStatement,
    MidStatement,
    ErrorStatement,
    ClassDeclaration,
    NewExpression,
    Parser,
    ResumeStatement,
    OpenStatement,
    CloseStatement,
    PrintStatement,
    LineInputStatement,
    PutStatement,
    KillStatement,
    WriteStatement,
    InputStatement,
    GetStatement,
    SeekStatement,
    ResetStatement,
    AttributeStatement,
    DeclareStatement,
    AddressOfExpression,
    DictionaryAccessExpression,
    TypeOfIsExpression,
    NamedArgument,
    OptionPrivateModuleStatement,
    EventDeclaration,
    EnumDeclaration,
    OptionExplicitStatement,
    RaiseEventStatement,
    ImplementsDirective,
    AppActivateStatement,
    SendKeysStatement,
    LockStatement,
    UnlockStatement,
    WidthStatement,
    OptionCompareStatement,
    OptionBaseStatement,
    DefDirective,
    Parameter,
} from './parser';
import { Lexer, TokenType } from './lexer';
import { SandboxPath } from './sandbox';
import { FileSystem, MemoryFileSystem } from './filesystem';
import { checkOptionExplicit, walkProcForUndefinedCalls, UndefinedProcError } from './option-explicit-checker';
import * as path from 'path';
import {
    VbaBoolean, VbaDate, VbaDecimal, VbaErrorValue, VbaNamespaceRef,
    vbaEmpty, vbaNull, vbaNothing, vbaMissing,
    vbaTrue, vbaFalse,
    toVbaDate, fromVbaDate, parseVbaDate,
    createAutoInstancePlaceholder, isAutoInstancePlaceholder,
} from './vba-types';
import type { VbaVarType, VbaComObject } from './vba-types';
export {
    VbaBoolean, VbaDate, VbaDecimal, VbaErrorValue, VbaNamespaceRef,
    vbaEmpty, vbaNull, vbaNothing, vbaMissing,
    vbaTrue, vbaFalse,
    toVbaDate, fromVbaDate, parseVbaDate,
    createAutoInstancePlaceholder, isAutoInstancePlaceholder,
} from './vba-types';
export type { AutoInstancePlaceholder, VbaBooleanType, VbaNumericType, VbaVarType } from './vba-types';
import {
    vbaToNumber as _vbaToNumber,
    vbaToBoolean,
    vbaToString,
    vbaToDisplayString,
    vbaRound as _vbaRound,
} from './coerce';
import { VbaErrorCode, throwVbaError } from './vba-errors';
export { VbaErrorCode } from './vba-errors';

/**
 * VBARunner.spy() / Evaluator.spy() が返す呼び出し記録オブジェクト。
 * 各 call は引数の配列として記録される。
 */
export class SpyRecord {
    readonly calls: any[][] = [];
    readonly returnValues: any[] = [];

    get callCount(): number { return this.calls.length; }
    get lastCall(): any[] | null { return this.calls[this.calls.length - 1] ?? null; }

    /** 指定した引数で呼び出されたことがあるか（引数は先頭から部分一致） */
    calledWith(...args: any[]): boolean {
        return this.calls.some(callArgs =>
            args.every((arg, i) => callArgs[i] === arg)
        );
    }

    /** 呼び出し履歴をリセットする */
    reset(): void {
        this.calls.length = 0;
        this.returnValues.length = 0;
    }
}

class VbaCollection {
    private _items: { value: any, key: string | null }[] = [];
    public readonly __isVbaCollection__ = true;

    public add(item: any, key?: string, before?: any, after?: any) {
        const keyLower = (key !== undefined && key !== vbaEmpty && key !== null) ? String(key).toLowerCase() : null;
        if (keyLower && this._items.some(i => i.key === keyLower)) {
            throwVbaError(VbaErrorCode.KEY_ALREADY_EXISTS);
        }

        const newItem = { value: item, key: keyLower };

        if (before !== undefined && before !== vbaEmpty && before !== null) {
            const idx = this.findIndex(before) - 1;
            this._items.splice(idx, 0, newItem);
        } else if (after !== undefined && after !== vbaEmpty && after !== null) {
            const idx = this.findIndex(after);
            this._items.splice(idx, 0, newItem);
        } else {
            this._items.push(newItem);
        }
    }

    private findIndex(id: any): number {
        if (typeof id === 'number') {
            if (id < 1 || id > this._items.length)
                throwVbaError(VbaErrorCode.SUBSCRIPT_OUT_OF_RANGE);
            return id;
        } else if (id !== undefined && id !== null && id !== vbaEmpty) {
            const k = String(id).toLowerCase();
            const idx = this._items.findIndex(i => i.key === k);
            if (idx === -1)
                throwVbaError(VbaErrorCode.INVALID_PROCEDURE_CALL);
            return idx + 1;
        }
        throwVbaError(VbaErrorCode.INVALID_PROCEDURE_CALL);
    }

    public count() {
        return this._items.length;
    }

    public item(id: any) {
        const idx = this.findIndex(id);
        return this._items[idx - 1].value;
    }

    public get items(): any[] {
        return this._items.map(i => i.value);
    }

    public remove(id: any) {
        const idx = this.findIndex(id);
        this._items.splice(idx - 1, 1);
    }
}

export class VbaErrObject {
    public number: number = 0;
    public source: string = "";
    public description: string = "";
    public helpfile: string = "";
    public helpcontext: number = 0;
    public lastdllerror: number = 0;

    public clear() {
        this.number = 0;
        this.source = "";
        this.description = "";
        this.helpfile = "";
        this.helpcontext = 0;
        this.lastdllerror = 0;
    }

    public raise(number: number, source?: any, description?: any, helpfile?: any, helpcontext?: any) {
        this.number = number;
        if (source !== undefined && source !== vbaEmpty && source !== null) this.source = String(source);
        if (description !== undefined && description !== vbaEmpty && description !== null) this.description = String(description);
        if (helpfile !== undefined && helpfile !== vbaEmpty && helpfile !== null) this.helpfile = String(helpfile);
        if (helpcontext !== undefined && helpcontext !== vbaEmpty && helpcontext !== null) this.helpcontext = Number(helpcontext);

        throw { type: 'VbaError', number: this.number, message: this.description };
    }
}
export interface VbaTypeInfo {
    vbaType: VbaVarType;
}

export interface DebugHook {
    onBeforeStatement(
        line: number,
        callDepth: number,
        env: Environment,
        callStack: ReadonlyArray<{ name: string; moduleName: string; line: number }>
    ): void;
}

interface ExecProcBodyOptions {
    byRefArgs: { paramName: string; originalExpr: Expression }[];
    paramArrayParamName: string | null;
    paramArrayByRefExprs: Expression[];
    /** null in callProcedure, -1 in evaluateCallExpression */
    initialLastErrorIndex: number | null;
    /** callProcedure accepts Exit Property, evaluateCallExpression does not */
    acceptExitProperty: boolean;
    /** callProcedure returns property value, evaluateCallExpression does not */
    returnOnProperty: boolean;
    /** vbaEmpty for callProcedure, undefined for evaluateCallExpression */
    subReturnValue: any;
}

export class Environment {
    private variables: Map<string, any> = new Map();
    private variableTypes: Map<string, VbaTypeInfo> = new Map();
    private procedures: Map<string, ProcedureDeclaration> = new Map();
    private types: Map<string, TypeMember[]> = new Map();
    private withEventsVariables: Set<string> = new Set();
    private constantVariables: Set<string> = new Set();
    /** AST node identity tracking for duplicate Dim detection (same stmt re-run in loop is OK) */
    private dimStmtRegistry: Map<string, object> = new Map(); // varKey → VariableDeclaration AST node
    public enclosing?: Environment;

    constructor(enclosing?: Environment) {
        this.enclosing = enclosing;
    }

    set(name: string, value: any) {
        const key = name.toLowerCase();
        if (this.isConstant(key)) {
            throwVbaError(VbaErrorCode.INVALID_PROCEDURE_CALL, `Assignment to constant not allowed: '${name}'`);
        }
        if (this.variables.has(key)) {
            this.variables.set(key, this.coerceToType(key, value));
            return;
        }
        let env: Environment | undefined = this.enclosing;
        while (env) {
            if (env.variables.has(key)) {
                env.variables.set(key, env.coerceToType(key, value));
                return;
            }
            env = env.enclosing;
        }
        this.variables.set(key, value);
    }

    setLocally(name: string, value: any) {
        const key = name.toLowerCase();
        this.variables.set(key, this.coerceToType(key, value));
    }

    registerDimStmt(name: string, stmt: object): void {
        this.dimStmtRegistry.set(name.toLowerCase(), stmt);
    }

    getRegisteredDimStmt(name: string): object | undefined {
        return this.dimStmtRegistry.get(name.toLowerCase());
    }

    setVariableType(name: string, typeInfo: VbaTypeInfo) {
        this.variableTypes.set(name.toLowerCase(), typeInfo);
    }

    getVariableType(name: string): VbaTypeInfo | undefined {
        const key = name.toLowerCase();
        if (this.variableTypes.has(key)) {
            return this.variableTypes.get(key);
        }
        let env: Environment | undefined = this.enclosing;
        while (env) {
            if (env.variableTypes.has(key)) {
                return env.variableTypes.get(key);
            }
            env = env.enclosing;
        }
        return undefined;
    }

    private coerceToType(key: string, value: any): any {
        const typeInfo = this.variableTypes.get(key);
        if (!typeInfo) return value;

        // Don't coerce special VBA values
        if (value === vbaEmpty || value === undefined || value === vbaNull || value === vbaNothing || value === vbaMissing) {
            return value;
        }

        switch (typeInfo.vbaType) {
            case 'Byte': {
                const n = Environment.vbaRoundStatic(Environment.toNumeric(value));
                if (n < 0 || n > 255) Environment.throwOverflow();
                return n;
            }
            case 'Integer': {
                const n = Environment.vbaRoundStatic(Environment.toNumeric(value));
                if (n < -32768 || n > 32767) Environment.throwOverflow();
                return n;
            }
            case 'Long': {
                const n = Environment.vbaRoundStatic(Environment.toNumeric(value));
                if (n < -2147483648 || n > 2147483647) Environment.throwOverflow();
                return n;
            }
            case 'Single': {
                const n = Environment.toNumeric(value);
                const f32 = Math.fround(n);
                if (!isFinite(f32) && isFinite(n)) Environment.throwOverflow();
                return f32;
            }
            case 'Double':
                return Environment.toNumeric(value);
            case 'Currency': {
                const n = Environment.vbaRoundStatic(Environment.toNumeric(value), 4);
                if (n < -922337203685477.5808 || n > 922337203685477.5807) Environment.throwOverflow();
                return n;
            }
            case 'String':
                if (typeof value === 'number' || value instanceof VbaBoolean || value instanceof VbaDate) {
                    return String(value);
                }
                return value;
            case 'Boolean':
                if (typeof value === 'number') {
                    return value !== 0 ? vbaTrue : vbaFalse;
                }
                return value;
            default:
                return value;
        }
    }

    private static toNumeric(val: any): number { return _vbaToNumber(val); }

    private static vbaRoundStatic(val: number, decimals: number = 0): number { return _vbaRound(val, decimals); }

    private static throwOverflow(): never { throwVbaError(VbaErrorCode.OVERFLOW); }

    setWithEvents(name: string) {
        this.withEventsVariables.add(name.toLowerCase());
    }

    isWithEvents(name: string): boolean {
        const key = name.toLowerCase();
        if (this.withEventsVariables.has(key)) return true;
        if (this.enclosing) return this.enclosing.isWithEvents(name);
        return false;
    }

    setConstant(name: string, value: any) {
        const key = name.toLowerCase();
        this.variables.set(key, value);
        this.constantVariables.add(key);
    }

    isConstant(name: string): boolean {
        const key = name.toLowerCase();
        if (this.constantVariables.has(key)) return true;
        if (this.enclosing) return this.enclosing.isConstant(name);
        return false;
    }

    get(name: string): any {
        const key = name.toLowerCase();
        if (this.variables.has(key)) {
            return this.variables.get(key);
        }
        let env: Environment | undefined = this.enclosing;
        while (env) {
            if (env.variables.has(key)) {
                return env.variables.get(key);
            }
            env = env.enclosing;
        }

        // VBA without Option Explicit: implicit variable initialized to Empty
        this.variables.set(key, vbaEmpty);
        return vbaEmpty;
    }

    /** スコープチェーンを検索し、見つからなければ undefined を返す（暗黙初期化なし）。 */
    getConst(name: string): any {
        const key = name.toLowerCase();
        if (this.variables.has(key)) return this.variables.get(key);
        let env: Environment | undefined = this.enclosing;
        while (env) {
            if (env.variables.has(key)) return env.variables.get(key);
            env = env.enclosing;
        }
        return undefined;
    }

    /** resolveIdentifiers() 開始前のスナップショット取得用。全スコープのキー集合を返す。 */
    getAllKeys(): Set<string> {
        const keys = new Set<string>();
        let env: Environment | undefined = this;
        while (env) {
            for (const k of env.variables.keys()) keys.add(k);
            env = env.enclosing;
        }
        return keys;
    }

    /** このスコープのローカル変数のみを返す（デバッガー変数表示用）。 */
    getLocalVariables(): Map<string, any> {
        return new Map(this.variables);
    }

    hasVariable(name: string): boolean {
        const key = name.toLowerCase();
        if (this.variables.has(key)) return true;
        let env: Environment | undefined = this.enclosing;
        while (env) {
            if (env.variables.has(key)) return true;
            env = env.enclosing;
        }
        return false;
    }

    /** チェーンを辿らず現在スコープのみを確認する */
    hasOwnVariable(name: string): boolean {
        return this.variables.has(name.toLowerCase());
    }

    /** env チェーン全体の変数名・プロシージャ名を収集する（undefined call チェック用）。
     * プロシージャキーは "module:name" または "module:name:get" 形式のため全セグメントを追加する。 */
    collectAllNames(): Set<string> {
        const names = new Set<string>();
        let env: Environment | undefined = this;
        while (env) {
            for (const k of env.variables.keys()) names.add(k);
            for (const k of env.procedures.keys()) {
                for (const seg of k.split(':')) names.add(seg);
            }
            env = env.enclosing;
        }
        return names;
    }

    setProcedure(name: string, proc: ProcedureDeclaration) {
        const key = name.toLowerCase();
        if (proc.isProperty && proc.propertyType) {
            this.procedures.set(`${key}:${proc.propertyType}`, proc);
        } else {
            this.procedures.set(key, proc);
        }
    }

    setProcedureWithModule(name: string, proc: ProcedureDeclaration, moduleName: string) {
        const baseKey = `${moduleName.toLowerCase()}:${name.toLowerCase()}`;
        if (proc.isProperty && proc.propertyType) {
            this.procedures.set(`${baseKey}:${proc.propertyType}`, proc);
        } else {
            this.procedures.set(baseKey, proc);
        }
    }


    getProcedure(name: string, type?: 'get' | 'let' | 'set'): ProcedureDeclaration | undefined {
        const baseKey = name.toLowerCase();
        // If type is not specified, try baseKey then baseKey:get
        const keysToTry = type ? [`${baseKey}:${type}`] : [baseKey, `${baseKey}:get`];

        // Priority 1: Unqualified procedures (global scope)
        for (const key of keysToTry) {
            if (this.procedures.has(key)) {
                return this.procedures.get(key);
            }
        }

        // Priority 2: Module-qualified procedures with disambiguation
        // Collect all candidates from current environment
        const candidates: ProcedureDeclaration[] = [];
        for (const [key, proc] of this.procedures.entries()) {
            const matches = type
                ? key === `${baseKey}:${type}` || key.endsWith(`:${baseKey}:${type}`)
                : key === baseKey || key === `${baseKey}:get` ||
                  key.endsWith(`:${baseKey}`) || key.endsWith(`:${baseKey}:get`);
            if (matches) {
                candidates.push(proc);
            }
        }

        // Check for ambiguity: multiple module-qualified procedures
        if (candidates.length > 1) {
            const modules = candidates.map(p => p.moduleName).filter(m => m).join(', ');
            throwVbaError(VbaErrorCode.SUB_OR_FUNCTION_NOT_DEFINED, `Ambiguous procedure '${name}'. Found in multiple modules: ${modules}. Use module qualification (e.g., Module.${name}()) to disambiguate.`);
        }

        if (candidates.length === 1) {
            return candidates[0];
        }

        // Priority 3: Search in enclosing scopes
        let env: Environment | undefined = this.enclosing;
        while (env) {
            // Try unqualified first
            for (const key of keysToTry) {
                if (env.procedures.has(key)) {
                    return env.procedures.get(key);
                }
            }

            // Then try module-qualified with disambiguation
            const envCandidates: ProcedureDeclaration[] = [];
            for (const [key, proc] of env.procedures.entries()) {
                const matches = type
                    ? key === `${baseKey}:${type}` || key.endsWith(`:${baseKey}:${type}`)
                    : key === baseKey || key === `${baseKey}:get` ||
                      key.endsWith(`:${baseKey}`) || key.endsWith(`:${baseKey}:get`);
                if (matches) {
                    envCandidates.push(proc);
                }
            }

            if (envCandidates.length > 1) {
                const modules = envCandidates.map(p => p.moduleName).filter(m => m).join(', ');
                throwVbaError(VbaErrorCode.SUB_OR_FUNCTION_NOT_DEFINED, `Ambiguous procedure '${name}'. Found in multiple modules: ${modules}. Use module qualification (e.g., Module.${name}()) to disambiguate.`);
            }

            if (envCandidates.length === 1) {
                return envCandidates[0];
            }

            env = env.enclosing;
        }
        return undefined;
    }

    getProcedureFromModule(name: string, moduleName: string, type?: 'get' | 'let' | 'set'): ProcedureDeclaration | undefined {
        const baseKey = `${moduleName.toLowerCase()}:${name.toLowerCase()}`;
        // If type is not specified, try baseKey then baseKey:get
        const keysToTry = type ? [`${baseKey}:${type}`] : [baseKey, `${baseKey}:get`];

        for (const key of keysToTry) {
            if (this.procedures.has(key)) {
                return this.procedures.get(key);
            }
        }

        let env: Environment | undefined = this.enclosing;
        while (env) {
            for (const key of keysToTry) {
                if (env.procedures.has(key)) {
                    return env.procedures.get(key);
                }
            }
            env = env.enclosing;
        }
        return undefined;
    }

    setType(name: string, members: TypeMember[]) {
        this.types.set(name.toLowerCase(), members);
    }

    getType(name: string): TypeMember[] | undefined {
        const key = name.toLowerCase();
        if (this.types.has(key)) {
            return this.types.get(key);
        }
        let env: Environment | undefined = this.enclosing;
        while (env) {
            if (env.types.has(key)) {
                return env.types.get(key);
            }
            env = env.enclosing;
        }
        return undefined;
    }
}

export type PrintCallback = (output: string) => void;

/**
 * 組み込み関数（JS関数として登録される）の引数メタデータ。
 * VBA のユーザー定義 Sub/Function における `Parameter`（isOptional/isParamArray）に相当する
 * 情報を、組み込み関数にも持たせるための型。
 *
 * - 「必須引数のあとに Optional 引数が続く」通常の形は `registerBuiltin` + この型を使う。
 * - `InStr` のように引数の個数で意味が変わる（先頭の Optional 引数が個数によって有無する）
 *   不規則な組み込み関数は `BuiltinOverload`（VBA 自体にはないオーバーロード機構。
 *   組み込み関数専用のエンジン内部の仕組み）を使う。
 */
export interface BuiltinParamSpec {
    /** VBA 上のパラメーター名（大文字小文字無視で `:=` 名前付き引数とマッチングする） */
    name: string;
    optional?: boolean;
    isParamArray?: boolean;
}

/** `BuiltinParamSpec` と最小限の形を共有する抽象（ProcedureDeclaration.parameters も満たす） */
interface ArgBinderParam {
    isOptional: boolean;
    isParamArray: boolean;
}

/**
 * オーバーロードされた組み込み関数の1つの「形」。この中では全パラメーターが実質必須
 * （形をまたいだ Optional 表現は、より短いオーバーロードを別途登録することで行う）。
 */
export interface BuiltinOverload {
    params: BuiltinParamSpec[];
}

export class Evaluator {
    public env: Environment;          // Tier 3: Public cross-module names
    private builtinEnv!: Environment; // Tier 4: standard library
    private moduleEnvs: Map<string, Environment> = new Map(); // Tier 2: per-module vars
    private fileHandles: Map<number, {
        fd: number,
        mode: 'Input' | 'Output' | 'Append' | 'Random' | 'Binary',
        path: string,
        buffer?: Uint8Array,
        pos?: number
    }> = new Map();
    private sandbox: SandboxPath;
    private fs: FileSystem;
    private onPrint: PrintCallback;
    private currentProcBody: Statement[] | null = null;
    private currentProcedureName: string | null = null;
    private currentProcedureType: string | null = null;
    private currentSourceModule: string = '';
    private executingModuleName: string = '';
    // Maps module name (lower) -> set of variable/const names (lower) declared at module level
    private moduleVarRegistry: Map<string, Set<string>> = new Map();
    private withObjectStack: any[] = [];
    private gosubStack: number[] = []; // Stack of statement indices for GoSub/Return
    private staticVarStore: Map<string, any> = new Map(); // persistent store for Static variables
    private currentProcIsStatic: boolean = false;
    private arrayBase: number = 0;
    /** §5.2.2 Def-Directive: letter (a-z) → VBA type name */
    private defTypeMap: Map<string, string> = new Map();
    /** VarPtr/StrPtr/ObjPtr: dummy address counter (non-zero unique integers) */
    private _ptrCounter: number = 0x10000;
    /** Scope exit tracking: VBA class instances created with `New` in the current proc scope.
     *  Only these are eligible for auto-Terminate on scope exit. null outside any proc. */
    private _currentNewOwned: Set<any> | null = null;
    private staticVarsInCurrentProc: Set<string> = new Set();
    private errObj: VbaErrObject = new VbaErrObject();
    private classDefinitions: Map<string, ClassDeclaration> = new Map();
    /** §5.6.10 Tier 6: 修飾なし識別子の最終フォールバック先オブジェクト（MockApplication 等） */
    private defaultBindingObject: any = null;
    private comparisonMode: 'Binary' | 'Text' = 'Binary';
    private errorHandlerLabel: string | null = null;
    private errorHandlingMode: 'None' | 'ResumeNext' | 'GoTo' = 'None';
    private isInErrorHandler: boolean = false;
    private lastErrorIndex: number | null = null;
    private virtualRegistry: { [app: string]: { [section: string]: { [key: string]: string } } } = {};
    private dirIterator: string[] | null = null;
    private dirIndex: number = 0;
    private currentLine: number = 0;
    private nowOverride: (() => Date) | null = null;
    private optionExplicitViolations: Map<string, Map<string, number>> = new Map();
    /** 定数式評価中は true。Identifier 解決で resolveConstIdent() を使う。 */
    private inConstEval = false;
    private vbaCallStack: Array<{ name: string; moduleName: string; line: number }> = [];
    private debugHook: DebugHook | null = null;

    /**
     * resolveIdentifiers 完了後に true になる。
     * evaluateModule はこのフラグを見て、バッチロード中（false）はモジュールレベル
     * 実行文を pendingTopLevel に退避し、インタラクティブ呼び出し（true）では即時実行する。
     */
    private resolveIdentifiersDone = false;
    /** Pass 1 で配列境界を評価できなかった VariableDeclaration。Pass 2 で再評価。 */
    private pendingArrayDecls: Array<{ stmt: VariableDeclaration; moduleName: string }> = [];
    /** Pass 1 でシンボルテーブルに退避したモジュールレベル実行文。Pass 2 後に実行。 */
    private pendingTopLevel: Array<{ moduleName: string; stmts: Statement[] }> = [];
    /** vba-runner 拡張: モジュールレベル実行文（代入・For 等）をプロシージャの後にも書けるようにするか。
     *  false にすると標準 VBA 仕様どおり、プロシージャ後の実行文もコンパイルエラーになる。 */
    private allowTopLevelStatements: boolean = true;

    constructor(onPrint: PrintCallback, config: { sandboxRoot?: string, env?: Record<string, string>, fs?: FileSystem, allowTopLevelStatements?: boolean } = {}) {
        // Tier 4: builtinEnv — register standard library here first
        this.builtinEnv = new Environment();
        this.env = this.builtinEnv;         // temporarily point to builtinEnv
        this.onPrint = onPrint;
        this.sandbox = new SandboxPath(config.sandboxRoot, config.env);
        this.fs = config.fs || new MemoryFileSystem();
        if (config.allowTopLevelStatements !== undefined) this.allowTopLevelStatements = config.allowTopLevelStatements;
        this.registerStandardLibrary();     // → builtinEnv
        this.registerBuiltinExternalObjects(); // → typeLibraryNamespaces
        // Tier 3: globalEnv — cross-module public names, enclosing = builtinEnv
        this.env = new Environment(this.builtinEnv);
    }

    // Public accessor for testing/mocking
    getGlobalEnv(): Environment {
        return this.env;
    }

    setDebugHook(hook: DebugHook | null): void {
        this.debugHook = hook;
    }

    getVbaCallStack(): ReadonlyArray<{ name: string; moduleName: string; line: number }> {
        return this.vbaCallStack;
    }

    setNowFn(fn: (() => Date) | null): void {
        this.nowOverride = fn;
        this.registerDateTimeFunctions();
    }

    /**
     * 指定した VBA 組み込み関数 / ユーザー定義関数をスパイでラップし、呼び出し記録を返す。
     * returnFn を渡すと戻り値をオーバーライドできる（例: MsgBox を常に vbYes=6 にする）。
     * returnFn を省略すると元の実装をそのまま呼び出す。
     */
    spy(name: string, returnFn?: (...args: any[]) => any): SpyRecord {
        const key = name.toLowerCase();
        const original = this.env.get(key);
        const record = new SpyRecord();

        const spyFn = (...args: any[]) => {
            record.calls.push([...args]);
            let ret: any;
            if (returnFn) {
                ret = returnFn(...args);
            } else if (typeof original === 'function') {
                ret = original(...args);
            }
            record.returnValues.push(ret);
            return ret;
        };

        if (original && (original as any).__vbaParamSpec__) {
            (spyFn as any).__vbaParamSpec__ = (original as any).__vbaParamSpec__;
        }
        if (original && (original as any).__vbaOverloads__) {
            (spyFn as any).__vbaOverloads__ = (original as any).__vbaOverloads__;
        }

        this.env.set(key, spyFn);
        return record;
    }

    private getNow(): Date {
        return this.nowOverride ? this.nowOverride() : new Date();
    }

    /** Register a function under `name`, and also under `name$` when variants includes '$'. */
    private envSet(name: string, fn: any, variants: string[] = []) {
        this.env.set(name, fn);
        for (const v of variants) this.env.set(name + v, fn);
    }

    /**
     * `envSet` の上位互換: VBA パラメーターメタデータ（`BuiltinParamSpec[]`）を渡すと、
     * 呼び出し時に引数の数の検証・名前付き引数（`:=`）の解決ができるようになる
     * （`resolveCallArgs` 参照）。関数本体は今までどおり通常の位置引数として呼ばれる
     * （JS のデフォルト引数構文で Optional のデフォルト値を表現する）。
     *
     * 括弧無し参照時の自動呼び出し可否（`isAutoCallable`）はこの `params` から
     * その場で算出するため、別フラグとして持たない。
     */
    private registerBuiltin(name: string, fn: any, params: BuiltinParamSpec[], variants: string[] = []) {
        let seenOptional = false;
        const seenNames = new Set<string>();
        for (const p of params) {
            const lower = p.name.toLowerCase();
            if (seenNames.has(lower)) {
                throw new Error(`registerBuiltin('${name}'): duplicate parameter name '${p.name}'`);
            }
            seenNames.add(lower);
            if (p.optional) {
                seenOptional = true;
            } else if (seenOptional && !p.isParamArray) {
                throw new Error(`registerBuiltin('${name}'): required parameter '${p.name}' follows an optional parameter — use registerOverloadedBuiltin for irregular shapes`);
            }
        }
        (fn as any).__vbaParamSpec__ = params;
        this.envSet(name, fn, variants);
    }

    /**
     * VBA 自体にはないオーバーロード機構を、組み込み関数専用にエンジン内部で提供するための
     * 登録ヘルパー。`InStr`（`Start` が先頭にある Optional 引数）のように、引数の個数で
     * 意味が変わる不規則な組み込み関数に使う。関数本体は変更不要（位置引数のみの呼び出しは
     * 今までどおり本体の内部ロジックがそのまま処理する。名前付き引数を使った場合のみ、
     * 該当オーバーロードの順序で位置引数を再構築して渡す。`resolveCallArgs` 参照）。
     */
    private registerOverloadedBuiltin(name: string, fn: any, overloads: BuiltinOverload[], variants: string[] = []) {
        const seenShapes = new Set<string>();
        for (const overload of overloads) {
            const names = overload.params.map(p => p.name.toLowerCase());
            if (new Set(names).size !== names.length) {
                throw new Error(`registerOverloadedBuiltin('${name}'): duplicate parameter name within one overload`);
            }
            const shapeKey = `${names.length}:${[...names].sort().join(',')}`;
            if (seenShapes.has(shapeKey)) {
                throw new Error(`registerOverloadedBuiltin('${name}'): two overloads share both arity and parameter-name set — ambiguous spec`);
            }
            seenShapes.add(shapeKey);
        }
        (fn as any).__vbaOverloads__ = overloads;
        this.envSet(name, fn, variants);
    }

    private registerDateTimeFunctions() {
        const nowFunc = () => new VbaDate(toVbaDate(this.getNow()));
        this.registerBuiltin('now', nowFunc, []);

        const dateFunc = () => new VbaDate(Math.floor(toVbaDate(this.getNow())));
        this.registerBuiltin('date', dateFunc, []);

        const timeFunc = () => {
            const serial = toVbaDate(this.getNow());
            return new VbaDate(serial - Math.floor(serial));
        };
        this.registerBuiltin('time', timeFunc, []);

        const timerFunc = () => {
            const now = this.getNow();
            const midnight = new Date(now.getFullYear(), now.getMonth(), now.getDate());
            return (now.getTime() - midnight.getTime()) / 1000;
        };
        this.registerBuiltin('timer', timerFunc, []);
    }

    private registerStandardLibrary() {
        // VBA プロジェクト名前空間を事前定義。bare 使用（VarType(VBA) 等）を検出するためのセンチネル。
        this.env.set('vba', new VbaNamespaceRef('VBA', 'project'));
        this.registerCoreObjects();
        this.registerInformationFunctions();
        this.registerConversionFunctions();
        this.registerMathFunctions();
        this.registerStringFunctions();
        this.registerDateTimeFunctions();
        this.registerStdlibDateTimeFunctions();
        this.registerFileSystemFunctions();
        this.registerInteractionFunctions();
        this.registerFinancialFunctions();
        this.registerConstants();
        this.registerRegistryFunctions();
    }

    private registerCoreObjects() {
        this.env.set('debug', {
            print: (...args: any[]) => this.onPrint(args.map(a => this.toDisplayString(a)).join(' ')),
            assert: (condition: any) => {
                if (!this.isTrue(condition)) {
                    throw new Error('Assertion failed');
                }
            }
        });
        this.env.set('err', this.errObj);
        this.env.set('application', {
            wait: (time: any) => {
                this.onPrint(`[APPLICATION.WAIT] ${time}`);
                return vbaTrue;
            },
            statusbar: "",
            displayalerts: true,
            screenupdating: true,
        });
    }

    private registerInformationFunctions() {
        this.registerBuiltin('isempty', (val: any) => (val === undefined || val === null || val === vbaEmpty) ? vbaTrue : vbaFalse, [{ name: 'Expression' }]);
        this.registerBuiltin('ismissing', (val: any) => val === vbaMissing ? vbaTrue : vbaFalse, [{ name: 'ArgName' }]);
        this.registerBuiltin('isnumeric', (val: any) => {
            if (val === vbaNull) return vbaFalse;
            if (val === vbaEmpty || val === undefined) return vbaTrue;
            if (typeof val === 'number' || typeof val === 'bigint' || val instanceof VbaDecimal || val instanceof VbaBoolean || val instanceof VbaDate) return vbaTrue;
            if (typeof val === 'string') {
                const s = val.trim();
                if (s === "") return vbaFalse;
                const cleaned = s.replace(/[$,]/g, '');
                return (!isNaN(Number(cleaned)) && isFinite(Number(cleaned))) ? vbaTrue : vbaFalse;
            }
            return vbaFalse;
        }, [{ name: 'Expression' }]);
        this.registerBuiltin('isdate', (val: any) => {
            if (val instanceof VbaDate) return vbaTrue;
            if (typeof val === 'string') {
                const d = Date.parse(val);
                return !isNaN(d) ? vbaTrue : vbaFalse;
            }
            return vbaFalse;
        }, [{ name: 'Expression' }]);
        this.registerBuiltin('isobject', (val: any) => (val === vbaNothing || (val && typeof val === 'object' && !Array.isArray(val) && val !== vbaNull)) ? vbaTrue : vbaFalse, [{ name: 'Identifier' }]);
        this.registerBuiltin('iserror', (val: any) => (val instanceof VbaErrorValue) ? vbaTrue : vbaFalse, [{ name: 'Expression' }]);
        this.registerBuiltin('isnull', (val: any) => (val === vbaNull) ? vbaTrue : vbaFalse, [{ name: 'Expression' }]);
        this.registerBuiltin('isarray', (val: any) => Array.isArray(val) ? vbaTrue : vbaFalse, [{ name: 'VarName' }]);

        this.registerBuiltin('vartype', (val: any) => {
            if (val === vbaEmpty || val === undefined) return 0; // vbEmpty
            if (val === vbaNull) return 1; // vbNull
            if (val === vbaNothing) return 9; // vbObject
            if (val instanceof VbaBoolean) return 11; // vbBoolean
            if (val instanceof VbaDate) return 7; // vbDate
            if (val === vbaMissing || val instanceof VbaErrorValue) return 10; // vbError
            if (Array.isArray(val)) return 8192 + 12; // vbArray + vbVariant
            if (typeof val === 'number') return 5; // vbDouble
            if (typeof val === 'string') return 8; // vbString
            if (val instanceof VbaDecimal) return 14; // vbDecimal
            if (typeof val === 'bigint') return 20; // vbLongLong
            if (val && val.__vbaTypeName__) return 36; // vbUserDefinedType
            if (val && (val.__vbaClass__ || val.__isVbaDict__ || val.__isVbaCollection__)) return 9; // vbObject
            if (typeof val === 'object' && val !== null) return 9; // vbObject
            return 12; // vbVariant
        }, [{ name: 'VarName' }]);

        this.registerBuiltin('typename', (val: any) => {
            if (val === vbaEmpty || val === undefined) return 'Empty';
            if (val === vbaNull) return 'Null';
            if (val === vbaNothing) return 'Nothing';
            if (val === vbaMissing || val instanceof VbaErrorValue) return 'Error';
            if (val instanceof VbaBoolean) return 'Boolean';
            if (val instanceof VbaDate) return 'Date';
            if (typeof val === 'number') return 'Double';
            if (typeof val === 'string') return 'String';
            if (Array.isArray(val)) return 'Variant()';
            if (val && val.__vbaTypeName__) return val.__vbaTypeName__;
            if (val && val.__isVbaDict__) return 'Dictionary';
            if (val && val.__isVbaCollection__) return 'Collection';
            if (val && val.__vbaClass__) return val.__className__;
            if (val instanceof VbaDecimal) return 'Decimal';
            if (typeof val === 'bigint') return 'LongLong';
            return 'Object';
        }, [{ name: 'VarName' }]);

        // CallByName(object, procName, callType, args...)
        // callType: 1=VbMethod, 2=VbGet, 4=VbLet, 8=VbSet
        this.registerBuiltin('callbyname', (obj: any, procName: string, callType: number, ...args: any[]) => {
            if (obj === null || obj === undefined || obj === vbaNothing) {
                this.throwVbaError(VbaErrorCode.OBJECT_VARIABLE_NOT_SET, 'Object variable or With block variable not set');
            }
            const name = String(procName).toLowerCase();
            if (callType === 2 /* VbGet */ || callType === 1 /* VbMethod */) {
                if (obj.__vbaClass__) {
                    const classDef = obj.__classDef__ as ClassDeclaration;
                    const getter = classDef.procedures.find(
                        p => p.isProperty && p.propertyType === 'get' && p.name.name.toLowerCase() === name
                    );
                    if (getter) return this.callClassMethod(obj, getter, args);
                    const method = classDef.procedures.find(
                        p => !p.isProperty && p.name.name.toLowerCase() === name
                    );
                    if (method) return this.callClassMethod(obj, method, args);
                    return obj.__instanceEnv__.get(name);
                }
                if (typeof obj === 'object' && obj !== null) {
                    const keys = Object.keys(obj);
                    const match = keys.find(k => k.toLowerCase() === name) ?? name;
                    const val = obj[match];
                    if (typeof val === 'function') return val.apply(obj, args);
                    return val;
                }
            }
            this.throwVbaError(VbaErrorCode.OBJECT_DOESNT_SUPPORT_PROPERTY, `Object doesn't support this property or method: '${procName}'`);
        }, [
            { name: 'Object' },
            { name: 'ProcName' },
            { name: 'CallType' },
            { name: 'Args', isParamArray: true },
        ]);
    }

    private registerConversionFunctions() {
        this.registerBuiltin('cbyte', (val: any) => {
            if (val instanceof VbaBoolean) {
                return val.valueOf() ? 255 : 0;
            }
            const n = this.vbaRound(this.toVbaNumber(val));
            if (n < 0 || n > 255) this.throwVbaError(VbaErrorCode.OVERFLOW, "Overflow");
            return n;
        }, [{ name: 'Expression' }]);
        this.registerBuiltin('cint', (val: any) => {
            const n = this.vbaRound(this.toVbaNumber(val));
            if (n < -32768 || n > 32767) this.throwVbaError(VbaErrorCode.OVERFLOW, "Overflow");
            return n;
        }, [{ name: 'Expression' }]);
        this.registerBuiltin('clng', (val: any) => {
            const n = this.vbaRound(this.toVbaNumber(val));
            if (n < -2147483648 || n > 2147483647) this.throwVbaError(VbaErrorCode.OVERFLOW, "Overflow");
            return n;
        }, [{ name: 'Expression' }]);
        this.registerBuiltin('csng', (val: any) => {
            const n = this.toVbaNumber(val);
            const f32 = Math.fround(n);
            if (!isFinite(f32) && isFinite(n)) this.throwVbaError(VbaErrorCode.OVERFLOW, "Overflow");
            return f32;
        }, [{ name: 'Expression' }]);
        this.registerBuiltin('cdbl', (val: any) => this.toVbaNumber(val), [{ name: 'Expression' }]);
        this.registerBuiltin('cdate', (val: any) => {
            if (val === null || val === vbaNull || val === vbaEmpty) this.throwVbaError(VbaErrorCode.TYPE_MISMATCH, "Type mismatch");
            if (val instanceof VbaDate) return val;
            if (typeof val === 'string' && !/^\d+(\.\d+)?$/.test(val)) {
                return new VbaDate(toVbaDate(parseVbaDate(val)));
            }
            return new VbaDate(this.toVbaNumber(val));
        }, [{ name: 'Expression' }]);
        this.registerBuiltin('cvdate', (val: any) => {
            if (val === vbaNull) return vbaNull;
            return (this.env.get('cdate') as Function)(val);
        }, [{ name: 'Expression' }]);
        this.registerBuiltin('cdec', (val: any) => new VbaDecimal(this.toVbaNumber(val)), [{ name: 'Expression' }]);
        this.registerBuiltin('ccur', (val: any) => {
            const n = this.vbaRound(this.toVbaNumber(val), 4);
            if (n < -922337203685477.5808 || n > 922337203685477.5807) this.throwVbaError(VbaErrorCode.OVERFLOW, "Overflow");
            return n;
        }, [{ name: 'Expression' }]);
        const clnglngFunc = (val: any) => {
            if (val === vbaNull || val === null) this.throwVbaError(VbaErrorCode.TYPE_MISMATCH, "Type mismatch");
            if (typeof val === 'bigint') return val;
            if (typeof val === 'string') {
                const trimmed = val.trim();
                if (/^-?\d+$/.test(trimmed)) {
                    try {
                        const n = BigInt(trimmed);
                        if (n < -9223372036854775808n || n > 9223372036854775807n) this.throwVbaError(VbaErrorCode.OVERFLOW, "Overflow");
                        return n;
                    } catch {
                        this.throwVbaError(VbaErrorCode.TYPE_MISMATCH, "Type mismatch");
                    }
                }
            }
            const num = this.toVbaNumber(val);
            // Check range using number (approximate but necessary for large floats)
            // Note: 9223372036854775807 is exactly represented as a double but the next double is 9223372036854775807 + 1024
            if (num < -9223372036854775808 || num > 9223372036854775807) {
                // If it's very close to the edge, it might be a precision issue
                if (Math.abs(num - 9223372036854775807) < 1025 || Math.abs(num + 9223372036854775808) < 1025) {
                    // Keep going and let BigInt conversion handle it if it was a float that rounded to the edge
                } else {
                    this.throwVbaError(VbaErrorCode.OVERFLOW, "Overflow");
                }
            }
            const n = BigInt(this.vbaRound(num));
            if (n < -9223372036854775808n || n > 9223372036854775807n) this.throwVbaError(VbaErrorCode.OVERFLOW, "Overflow");
            return n;
        };
        this.registerBuiltin('clnglng', clnglngFunc, [{ name: 'Expression' }]);
        this.envSet('clngptr', this.env.get('clnglng')); // clnglng と同じ関数オブジェクトのため __vbaParamSpec__ も引き継ぐ
        this.registerBuiltin('cstr', (val: any) => {
            try { return vbaToString(val); } catch (e: any) {
                if (e?.type === 'VbaError') this.throwVbaError(e.number, e.message);
                throw e;
            }
        }, [{ name: 'Expression' }]);
        this.registerBuiltin('cbool', (val: any) => {
            if (val === vbaNull) this.throwVbaError(VbaErrorCode.INVALID_USE_OF_NULL, 'Invalid use of Null');
            return this.coerceToBoolean(val);
        }, [{ name: 'Expression' }]);
        this.registerBuiltin('cvar', (val: any) => val, [{ name: 'Expression' }]);
        this.registerBuiltin('cverr', (val: any) => {
            if (val instanceof VbaErrorValue) return val;
            const code = this.toVbaNumber(val);
            if (code < 0 || code > 65535) this.throwVbaError(VbaErrorCode.INVALID_PROCEDURE_CALL, "Invalid procedure call or argument");
            return new VbaErrorValue(code);
        }, [{ name: 'ErrorNumber' }]);

        const hexFn = (n: any) => {
            if (n === vbaNull) return vbaNull;
            return (Math.floor(this.toVbaNumber(n)) >>> 0).toString(16).toUpperCase();
        };
        this.registerBuiltin('hex', hexFn, [{ name: 'Number' }], ['$']);
        const octFn = (n: any) => {
            if (n === vbaNull) return vbaNull;
            return (Math.floor(this.toVbaNumber(n)) >>> 0).toString(8);
        };
        this.registerBuiltin('oct', octFn, [{ name: 'Number' }], ['$']);
        this.registerBuiltin('val', (s: any) => {
            if (typeof s !== 'string') return 0;
            const cleaned = s.trim().replace(/ /g, '');
            if (cleaned.toLowerCase().startsWith('&h')) return parseInt(cleaned.slice(2), 16) || 0;
            if (cleaned.toLowerCase().startsWith('&o')) return parseInt(cleaned.slice(2), 8) || 0;
            const match = cleaned.match(/^[+-]?\d*(\.\d*)?([eE][+-]?\d+)?/);
            return match ? parseFloat(match[0]) || 0 : 0;
        }, [{ name: 'String' }]);

        // Constants
        this.env.set('vbcrlf', "\r\n");
        this.env.set('vbcr', "\r");
        this.env.set('vblf', "\n");
        this.env.set('vbtab', "\t");
        this.env.set('vbnullstring', "");
        this.env.set('vbnullchar', "\0");
        this.env.set('vbok', 1);
        this.env.set('vbcancel', 2);
        this.env.set('vbabort', 3);
        this.env.set('vbretry', 4);
        this.env.set('vbignore', 5);
        this.env.set('vbyes', 6);
        this.env.set('vbno', 7);
        // CallByName callType constants
        this.env.set('vbmethod', 1);
        this.env.set('vbget', 2);
        this.env.set('vblet', 4);
        this.env.set('vbset', 8);

        this.env.set('vbokonly', 0);
        this.env.set('vbokcancel', 1);
        this.env.set('vbabortretryignore', 2);
        this.env.set('vbyesnocancel', 3);
        this.env.set('vbyesno', 4);
        this.env.set('vbretrycancel', 5);
        this.env.set('vbcritical', 16);
        this.env.set('vbquestion', 32);
        this.env.set('vbexclamation', 48);
        this.env.set('vbinformation', 64);
        this.env.set('vbdefaultbutton1', 0);
        this.env.set('vbdefaultbutton2', 256);
        this.env.set('vbdefaultbutton3', 512);
        this.env.set('vbdefaultbutton4', 768);
        this.env.set('vbtextcompare', 1);
        this.env.set('vbbinarycompare', 0);
        this.env.set('vbuppercase', 1);
        this.env.set('vblowercase', 2);
        this.env.set('vbpropercase', 3);
        this.env.set('vbwide', 4);
        this.env.set('vbnarrow', 8);
        this.env.set('vbkatakana', 16);
        this.env.set('vbhiragana', 32);
    }

    private registerMathFunctions() {
        this.registerBuiltin('abs', (val: any) => val === vbaNull ? vbaNull : Math.abs(this.toVbaNumber(val)), [{ name: 'Number' }]);
        this.registerBuiltin('atn', (val: any) => val === vbaNull ? vbaNull : Math.atan(this.toVbaNumber(val)), [{ name: 'Number' }]);
        this.registerBuiltin('cos', (val: any) => val === vbaNull ? vbaNull : Math.cos(this.toVbaNumber(val)), [{ name: 'Number' }]);
        this.registerBuiltin('exp', (val: any) => {
            if (val === vbaNull) return vbaNull;
            const n = this.toVbaNumber(val);
            if (n > 709.782712893) this.throwVbaError(VbaErrorCode.OVERFLOW, "Overflow");
            return Math.exp(n);
        }, [{ name: 'Number' }]);
        this.registerBuiltin('int', (val: any) => val === vbaNull ? vbaNull : Math.floor(this.toVbaNumber(val)), [{ name: 'Number' }]);
        this.registerBuiltin('fix', (val: any) => {
            if (val === vbaNull) return vbaNull;
            const n = this.toVbaNumber(val);
            return n >= 0 ? Math.floor(n) : Math.ceil(n);
        }, [{ name: 'Number' }]);
        this.registerBuiltin('log', (val: any) => {
            if (val === vbaNull) return vbaNull;
            const n = this.toVbaNumber(val);
            if (n <= 0) this.throwVbaError(VbaErrorCode.INVALID_PROCEDURE_CALL, "Invalid procedure call or argument");
            return Math.log(n);
        }, [{ name: 'Number' }]);
        this.registerBuiltin('round', (val: any, digits: any = 0) => val === vbaNull ? vbaNull : this.vbaRound(this.toVbaNumber(val), Number(digits)), [
            { name: 'Number' },
            { name: 'NumDigitsAfterDecimal', optional: true },
        ]);
        this.registerBuiltin('sgn', (val: any) => {
            if (val === vbaNull) return vbaNull;
            const n = this.toVbaNumber(val);
            return n > 0 ? 1 : n < 0 ? -1 : 0;
        }, [{ name: 'Number' }]);
        this.registerBuiltin('sin', (val: any) => val === vbaNull ? vbaNull : Math.sin(this.toVbaNumber(val)), [{ name: 'Number' }]);
        this.registerBuiltin('sqr', (val: any) => {
            if (val === vbaNull) return vbaNull;
            const n = this.toVbaNumber(val);
            if (n < 0) this.throwVbaError(VbaErrorCode.INVALID_PROCEDURE_CALL, "Invalid procedure call or argument");
            return Math.sqrt(n);
        }, [{ name: 'Number' }]);
        this.registerBuiltin('tan', (val: any) => val === vbaNull ? vbaNull : Math.tan(this.toVbaNumber(val)), [{ name: 'Number' }]);

        let rndSeed = 0.5;
        let lastRnd = 0.5;
        let rndInitialized = false;
        const rndFunc = (val?: any) => {
            if (!rndInitialized) { rndSeed = 0.5; rndInitialized = true; }
            if (val === undefined || (typeof val === 'number' && val > 0)) {
                rndSeed = (rndSeed * 214013 + 2531011) % 4294967296;
                lastRnd = rndSeed / 4294967296;
                return lastRnd;
            } else if (val === 0) {
                return lastRnd;
            } else if (val < 0) {
                const s = Math.abs(Number(val)) * 9301 + 49297;
                lastRnd = (s % 233280) / 233280;
                return lastRnd;
            }
            return lastRnd;
        };
        this.registerBuiltin('rnd', rndFunc, [{ name: 'Number', optional: true }]);
        this.registerBuiltin('randomize', (val?: any) => {
            rndInitialized = true;
            rndSeed = (val === undefined || val === null) ? (Date.now() % 4294967296) : (Math.round(Math.abs(Number(val)) * 1000) % 4294967296);
            lastRnd = rndSeed / 4294967296;
        }, [{ name: 'Number', optional: true }]);
    }

    private registerStringFunctions() {
        const ascFunc = (s: any) => String(s || '').charCodeAt(0);
        this.registerBuiltin('asc', ascFunc, [{ name: 'String' }]);
        this.registerBuiltin('ascw', ascFunc, [{ name: 'String' }]);
        const chrFunc = (n: any) => String.fromCharCode(Number(n));
        this.registerBuiltin('chr', chrFunc, [{ name: 'CharCode' }], ['$']);
        this.registerBuiltin('chrw', chrFunc, [{ name: 'CharCode' }], ['$']);
        // InStr: Start は先頭にある Optional 引数のため、引数の個数で意味が変わる
        // （registerOverloadedBuiltin — 本体のこの個数判定ロジックは変更しない）。
        const instrFunc = (...args: any[]) => {
            let start = 1, s1, s2, comp;
            if (args.length >= 4) [start, s1, s2, comp] = args;
            else if (args.length === 3 && typeof args[0] === 'number') [start, s1, s2] = args;
            else [s1, s2] = args;
            if (s1 === vbaNull || s2 === vbaNull) return vbaNull;
            const str1 = String(s1 ?? ''), str2 = String(s2 ?? '');
            const isText = (comp === 1) || (comp === undefined && this.comparisonMode === 'Text');
            const idx = isText ? str1.toLowerCase().indexOf(str2.toLowerCase(), start - 1) : str1.indexOf(str2, start - 1);
            return idx === -1 ? 0 : idx + 1;
        };
        this.registerOverloadedBuiltin('instr', instrFunc, [
            { params: [{ name: 'String1' }, { name: 'String2' }] },
            { params: [{ name: 'Start' }, { name: 'String1' }, { name: 'String2' }] },
            { params: [{ name: 'Start' }, { name: 'String1' }, { name: 'String2' }, { name: 'Compare' }] },
        ]);
        // InStrB: バイト単位での検索（VBA では文字列を UTF-16 として扱うため 1 文字 = 2 バイト）
        // start もバイト位置、戻り値もバイト位置（1-based）。InStr と同じく Start が先頭の Optional。
        const instrbFunc = (...args: any[]) => {
            let startByte = 1, s1, s2, comp;
            if (args.length >= 4) [startByte, s1, s2, comp] = args;
            else if (args.length === 3 && typeof args[0] === 'number') [startByte, s1, s2] = args;
            else [s1, s2] = args;
            if (s1 === vbaNull || s2 === vbaNull) return vbaNull;
            const str1 = String(s1 ?? ''), str2 = String(s2 ?? '');
            // バイト位置 → 文字位置に変換（1-based のまま）
            const startChar = Math.floor((Number(startByte) - 1) / 2) + 1;
            const isText = (comp === 1) || (comp === undefined && this.comparisonMode === 'Text');
            const idx = isText ? str1.toLowerCase().indexOf(str2.toLowerCase(), startChar - 1) : str1.indexOf(str2, startChar - 1);
            return idx === -1 ? 0 : idx * 2 + 1;
        };
        this.registerOverloadedBuiltin('instrb', instrbFunc, [
            { params: [{ name: 'String1' }, { name: 'String2' }] },
            { params: [{ name: 'Start' }, { name: 'String1' }, { name: 'String2' }] },
            { params: [{ name: 'Start' }, { name: 'String1' }, { name: 'String2' }, { name: 'Compare' }] },
        ]);
        this.registerBuiltin('instrrev', (s1: any, s2: any, start: any = -1, comp: any = undefined) => {
            if (s1 === vbaNull || s2 === vbaNull) return vbaNull;
            const str = String(s1 ?? ''), find = String(s2 ?? '');
            if (find === "") return (start === -1) ? str.length : Number(start);

            const startNum = Number(start);
            if (startNum !== -1 && startNum > str.length) return 0;

            const effStart = (start === -1) ? str.length : startNum;
            const isText = (comp === 1) || (comp === undefined && this.comparisonMode === 'Text');
            const idx = isText ? str.toLowerCase().lastIndexOf(find.toLowerCase(), effStart - 1) : str.lastIndexOf(find, effStart - 1);
            return idx === -1 ? 0 : idx + 1;
        }, [
            { name: 'StringCheck' },
            { name: 'StringMatch' },
            { name: 'Start', optional: true },
            { name: 'Compare', optional: true },
        ]);
        const lcaseFunc = (val: any) => val === vbaNull ? vbaNull : String(val ?? '').toLowerCase();
        this.registerBuiltin('lcase', lcaseFunc, [{ name: 'String' }], ['$']);
        const strFunc = (val: any) => {
            if (val === vbaNull) return vbaNull;
            const n = this.toVbaNumber(val);
            return n >= 0 ? " " + n : String(n);
        };
        this.registerBuiltin('str', strFunc, [{ name: 'Number' }], ['$']);

        const ucaseFunc = (val: any) => val === vbaNull ? vbaNull : String(val ?? '').toUpperCase();
        this.registerBuiltin('ucase', ucaseFunc, [{ name: 'String' }], ['$']);
        const leftFunc = (val: any, len: any) => String(val ?? '').substring(0, Number(len));
        this.registerBuiltin('left', leftFunc, [{ name: 'String' }, { name: 'Length' }], ['$']);
        const rightFunc = (val: any, len: any) => {
            const s = String(val ?? ''), l = Number(len);
            return s.substring(s.length - l);
        };
        this.registerBuiltin('right', rightFunc, [{ name: 'String' }, { name: 'Length' }], ['$']);
        const midFunc = (val: any, start: any, len?: any) => {
            const s = String(val ?? ''), st = Number(start);
            return len !== undefined ? s.substring(st - 1, st - 1 + Number(len)) : s.substring(st - 1);
        };
        this.registerBuiltin('mid', midFunc, [
            { name: 'String' },
            { name: 'Start' },
            { name: 'Length', optional: true },
        ], ['$']);
        this.registerBuiltin('len', (val: any) => val === vbaNull ? vbaNull : String(val ?? '').length, [{ name: 'String' }]);
        const ltrimFunc = (val: any) => val === vbaNull ? vbaNull : String(val ?? '').trimStart();
        this.registerBuiltin('ltrim', ltrimFunc, [{ name: 'String' }], ['$']);
        const rtrimFunc = (val: any) => val === vbaNull ? vbaNull : String(val ?? '').trimEnd();
        this.registerBuiltin('rtrim', rtrimFunc, [{ name: 'String' }], ['$']);
        const trimFunc = (val: any) => val === vbaNull ? vbaNull : String(val ?? '').trim();
        this.registerBuiltin('trim', trimFunc, [{ name: 'String' }], ['$']);
        const spaceFunc = (n: any) => ' '.repeat(Number(n));
        this.registerBuiltin('space', spaceFunc, [{ name: 'Number' }], ['$']);
        const stringFunc = (n: any, char: any) => {
            // VBA 仕様: 数値が渡された場合は文字コードとして扱う、文字列の場合は先頭文字を使う
            let c: string;
            if (typeof char === 'number') {
                c = String.fromCharCode(char);
            } else {
                const s = String(char ?? '');
                c = s.length > 0 ? s[0] : '';
            }
            return c.repeat(Number(n));
        };
        this.registerBuiltin('string', stringFunc, [{ name: 'Number' }, { name: 'Character' }], ['$']);
        this.registerBuiltin('split', (s: any, del: string = ' ') => String(s ?? '').split(del), [
            { name: 'Expression' },
            { name: 'Delimiter', optional: true },
        ]);
        this.registerBuiltin('join', (arr: any, del: string = ' ') => {
            if (!Array.isArray(arr)) return String(arr);
            // 配列の物理ストレージは LBound 分のオフセット（隠し要素）を含むため、
            // UBound/LBound と同じ vbaBase を見てスキップしないと先頭に余分な空要素が混入する
            const base = (arr as any).vbaBase || 0;
            return arr.slice(base).join(del);
        }, [
            { name: 'SourceArray' },
            { name: 'Delimiter', optional: true },
        ]);
        this.registerBuiltin('replace', (s: any, f: any, r: any) => String(s ?? '').split(String(f ?? '')).join(String(r ?? '')), [
            { name: 'Expression' },
            { name: 'Find' },
            { name: 'Replace' },
        ]);
        this.registerBuiltin('strcomp', (s1: any, s2: any, comp?: number) => {
            if (s1 === vbaNull || s2 === vbaNull) return vbaNull;
            let str1 = String(s1 ?? ''), str2 = String(s2 ?? '');
            const isText = (comp === 1) || (comp === undefined && this.comparisonMode === 'Text');
            if (isText) { str1 = str1.toLowerCase(); str2 = str2.toLowerCase(); }
            return str1 < str2 ? -1 : (str1 > str2 ? 1 : 0);
        }, [
            { name: 'String1' },
            { name: 'String2' },
            { name: 'Compare', optional: true },
        ]);
        this.registerBuiltin('strconv', (s: any, conv: any) => {
            if (s === vbaNull) return vbaNull;
            let str = String(s ?? '');
            const c = Number(conv);
            const caseConv = c & 3;
            if (caseConv === 1) str = str.toUpperCase();
            else if (caseConv === 2) str = str.toLowerCase();
            else if (caseConv === 3) str = str.split(/\s+/).map(w => w.charAt(0).toUpperCase() + w.slice(1).toLowerCase()).join(' ');

            if (c & 16) {
                str = str.replace(/[\u3041-\u3096]/g, m => String.fromCharCode(m.charCodeAt(0) + 0x60));
            } else if (c & 32) {
                str = str.replace(/[\u30A1-\u30F6]/g, m => String.fromCharCode(m.charCodeAt(0) - 0x60));
            }

            if (c & 4) {
                str = str.replace(/[!-~]/g, m => String.fromCharCode(m.charCodeAt(0) + 0xFEE0)).replace(/ /g, '　');
            } else if (c & 8) {
                str = str.replace(/[！-～]/g, m => String.fromCharCode(m.charCodeAt(0) - 0xFEE0)).replace(/　/g, ' ');
            }
            return str;
        }, [{ name: 'String' }, { name: 'Conversion' }]);
        this.registerBuiltin('strreverse', (s: any) => s === vbaNull ? vbaNull : String(s ?? '').split('').reverse().join(''), [{ name: 'Expression' }]);
        this.registerBuiltin('filter', (source: any, match: any, include: any = vbaTrue, compare: any = undefined) => {
            if (!Array.isArray(source)) this.throwVbaError(VbaErrorCode.TYPE_MISMATCH, "Type mismatch");
            const find = String(match ?? '');
            const isInclude = this.isTrue(include);
            const isText = (compare === 1) || (compare === undefined && this.comparisonMode === 'Text');
            const result = source.filter(s => {
                const str = String(s ?? '');
                const found = isText ? str.toLowerCase().includes(find.toLowerCase()) : str.includes(find);
                return isInclude ? found : !found;
            });
            (result as any).vbaBase = 0;
            return result;
        }, [
            { name: 'SourceArray' },
            { name: 'Match' },
            { name: 'Include', optional: true },
            { name: 'Compare', optional: true },
        ]);
        this.registerBuiltin('leftb', (val: any, len: any) => {
            const s = String(val ?? '');
            // In VBA, strings are UTF-16, so each char is 2 bytes.
            // LeftB(s, 2) returns first char.
            return s.substring(0, Math.floor(Number(len) / 2));
        }, [{ name: 'String' }, { name: 'Length' }]);
        this.registerBuiltin('rightb', (val: any, len: any) => {
            const s = String(val ?? '');
            const charLen = Math.floor(Number(len) / 2);
            return s.substring(s.length - charLen);
        }, [{ name: 'String' }, { name: 'Length' }]);
        const midbFunc = (val: any, start: any, len?: any) => {
            const s = String(val ?? '');
            const charStart = Math.floor((Number(start) + 1) / 2);
            if (len === undefined) return s.substring(charStart - 1);
            const charLen = Math.floor(Number(len) / 2);
            return s.substring(charStart - 1, charStart - 1 + charLen);
        };
        this.registerBuiltin('midb', midbFunc, [
            { name: 'String' },
            { name: 'Start' },
            { name: 'Length', optional: true },
        ], ['$']);
        const formatFunc = (val: any, pattern?: string) => {
            if (val === null || val === vbaNull || val === vbaEmpty) return "";
            const fmt = pattern ? String(pattern) : "";
            if (fmt === "") return String(val);
            const fmtLower = fmt.toLowerCase();
            const namedFormats = ['general number', 'currency', 'fixed', 'standard', 'percent', 'scientific', 'true/false', 'yes/no', 'on/off'];
            const dateNamedFormats = ['general date', 'long date', 'medium date', 'short date', 'long time', 'medium time', 'short time'];
            if (namedFormats.includes(fmtLower)) {
                if (typeof val === 'number') return this.formatNumber(val, fmt);
                return String(val);
            }
            if (dateNamedFormats.includes(fmtLower)) {
                const dateVal = val instanceof VbaDate ? fromVbaDate(val.value) : (typeof val === 'number' ? fromVbaDate(val) : new Date(String(val)));
                return this.formatDate(dateVal, fmt);
            }
            const isDatePattern = /y|m|d|h|n|s|am\/pm/i.test(fmt);
            if (val instanceof VbaDate) return this.formatDate(fromVbaDate(val.value), fmt);
            if (typeof val === 'number') {
                if (isDatePattern && !/^[0#,.%]+$/.test(fmt)) return this.formatDate(fromVbaDate(val), fmt);
                return this.formatNumber(val, fmt);
            }
            return String(val);
        };
        this.registerBuiltin('format', formatFunc, [
            { name: 'Expression' },
            { name: 'Format', optional: true },
        ], ['$']);
    }

    private registerStdlibDateTimeFunctions() {
        this.registerBuiltin('year',   (d: any) => parseVbaDate(d).getFullYear(), [{ name: 'Date' }]);
        this.registerBuiltin('month',  (d: any) => parseVbaDate(d).getMonth() + 1, [{ name: 'Date' }]);
        this.registerBuiltin('day',    (d: any) => parseVbaDate(d).getDate(), [{ name: 'Date' }]);
        this.registerBuiltin('hour',   (d: any) => parseVbaDate(d).getHours(), [{ name: 'Time' }]);
        this.registerBuiltin('minute', (d: any) => parseVbaDate(d).getMinutes(), [{ name: 'Time' }]);
        this.registerBuiltin('second', (d: any) => parseVbaDate(d).getSeconds(), [{ name: 'Time' }]);
        this.registerBuiltin('dateserial', (y: any, m: any, d: any) => new VbaDate(toVbaDate(new Date(Number(y), Number(m) - 1, Number(d)))), [
            { name: 'Year' }, { name: 'Month' }, { name: 'Day' },
        ]);
        this.registerBuiltin('timeserial', (h: any, n: any, s: any) => new VbaDate(toVbaDate(new Date(1899, 11, 30, Number(h), Number(n), Number(s)))), [
            { name: 'Hour' }, { name: 'Minute' }, { name: 'Second' },
        ]);
        this.registerBuiltin('weekday', (d: any) => parseVbaDate(d).getDay() + 1, [{ name: 'Date' }]);
        this.registerBuiltin('dateadd', (interval: any, number: any, date: any) => {
            const d = parseVbaDate(date);
            const n = Number(number);
            const intv = String(interval).toLowerCase();

            if (intv === 'yyyy' || intv === 'q' || intv === 'm') {
                const oldDay = d.getDate();
                if (intv === 'yyyy') d.setFullYear(d.getFullYear() + n);
                else if (intv === 'q') d.setMonth(d.getMonth() + n * 3);
                else if (intv === 'm') d.setMonth(d.getMonth() + n);

                // VBA behavior: if the day exceeds the last day of the month,
                // it is set to the last day of that month.
                if (d.getDate() !== oldDay) {
                    d.setDate(0);
                }
            } else if (intv === 'y' || intv === 'd' || intv === 'w') {
                d.setDate(d.getDate() + n);
            } else if (intv === 'ww') {
                d.setDate(d.getDate() + n * 7);
            } else if (intv === 'h') {
                d.setHours(d.getHours() + n);
            } else if (intv === 'n') {
                d.setMinutes(d.getMinutes() + n);
            } else if (intv === 's') {
                d.setSeconds(d.getSeconds() + n);
            }
            return new VbaDate(toVbaDate(d));
        }, [{ name: 'Interval' }, { name: 'Number' }, { name: 'Date' }]);
        this.registerBuiltin('datediff', (interval: any, date1: any, date2: any) => {
            const d1 = parseVbaDate(date1);
            const d2 = parseVbaDate(date2);
            const intv = String(interval).toLowerCase();
            const diffMs = d2.getTime() - d1.getTime();
            if (intv === 'yyyy') return d2.getFullYear() - d1.getFullYear();
            else if (intv === 'q') return (d2.getFullYear() - d1.getFullYear()) * 4 + Math.floor(d2.getMonth() / 3) - Math.floor(d1.getMonth() / 3);
            else if (intv === 'm') return (d2.getFullYear() - d1.getFullYear()) * 12 + d2.getMonth() - d1.getMonth();
            else if (intv === 'y' || intv === 'd' || intv === 'w') return Math.round(diffMs / 86400000);
            else if (intv === 'ww') return Math.round(diffMs / 604800000);
            else if (intv === 'h') return Math.round(diffMs / 3600000);
            else if (intv === 'n') return Math.round(diffMs / 60000);
            else if (intv === 's') return Math.round(diffMs / 1000);
            return 0;
        }, [{ name: 'Interval' }, { name: 'Date1' }, { name: 'Date2' }]);
        this.registerBuiltin('datepart', (interval: any, date: any) => {
            const d = parseVbaDate(date);
            const intv = String(interval).toLowerCase();
            if (intv === 'yyyy') return d.getFullYear();
            else if (intv === 'q') return Math.floor(d.getMonth() / 3) + 1;
            else if (intv === 'm') return d.getMonth() + 1;
            else if (intv === 'y') {
                const start = new Date(d.getFullYear(), 0, 0);
                const diff = d.getTime() - start.getTime();
                return Math.floor(diff / 86400000);
            }
            else if (intv === 'd') return d.getDate();
            else if (intv === 'w') return d.getDay() + 1;
            else if (intv === 'ww') {
                const start = new Date(d.getFullYear(), 0, 1);
                const diff = d.getTime() - start.getTime();
                return Math.floor((diff / 86400000 + start.getDay() + 6) / 7);
            }
            else if (intv === 'h') return d.getHours();
            else if (intv === 'n') return d.getMinutes();
            else if (intv === 's') return d.getSeconds();
            return 0;
        }, [{ name: 'Interval' }, { name: 'Date' }]);
        this.registerBuiltin('datevalue', (val: any) => {
            const d = parseVbaDate(val);
            return new VbaDate(Math.floor(toVbaDate(d)));
        }, [{ name: 'Date' }]);
        this.registerBuiltin('timevalue', (val: any) => {
            const d = parseVbaDate(val);
            const serial = toVbaDate(d);
            return new VbaDate(serial - Math.floor(serial));
        }, [{ name: 'Time' }]);
        this.registerBuiltin('monthname', (month: any, abbreviate: any = vbaFalse) => {
            const m = Number(month);
            if (m < 1 || m > 12) this.throwVbaError(VbaErrorCode.INVALID_PROCEDURE_CALL, "Invalid procedure call or argument");
            const names = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
            const abbrs = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
            return this.isTrue(abbreviate) ? abbrs[m - 1] : names[m - 1];
        }, [{ name: 'Month' }, { name: 'Abbreviate', optional: true }]);
        this.registerBuiltin('weekdayname', (weekday: any, abbreviate: any = vbaFalse, firstdayofweek: any = 1) => {
            const w = Number(weekday);
            let first = Number(firstdayofweek);
            if (first === 0) first = 1;
            if (w < 1 || w > 7) this.throwVbaError(VbaErrorCode.INVALID_PROCEDURE_CALL, "Invalid procedure call or argument");
            const names = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
            const abbrs = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
            const idx = (w + first - 2) % 7;
            return this.isTrue(abbreviate) ? abbrs[idx] : names[idx];
        }, [
            { name: 'Weekday' },
            { name: 'Abbreviate', optional: true },
            { name: 'FirstDayOfWeek', optional: true },
        ]);
    }

    private registerFileSystemFunctions() {
        const freeFileFunc = (range?: number) => {
            const start = (range === 1) ? 256 : 1;
            const end = (range === 1) ? 511 : 255;
            for (let i = start; i <= end; i++) if (!this.fileHandles.has(i)) return i;
            this.throwVbaError(VbaErrorCode.TOO_MANY_FILES, "Too many files");
        };
        this.registerBuiltin('freefile', freeFileFunc, [{ name: 'RangeNumber', optional: true }]);
        this.registerBuiltin('eof', (fn: any) => {
            const h = this.fileHandles.get(Number(fn));
            if (!h) this.throwVbaError(VbaErrorCode.BAD_FILE_NAME_OR_NUMBER, "Bad file name or number");
            return h.pos! >= this.fs.statSync(h.path).size ? vbaTrue : vbaFalse;
        }, [{ name: 'FileNumber' }]);
        this.registerBuiltin('lof', (fn: any) => {
            const h = this.fileHandles.get(Number(fn));
            if (!h) this.throwVbaError(VbaErrorCode.BAD_FILE_NAME_OR_NUMBER, "Bad file name or number");
            return this.fs.statSync(h.path).size;
        }, [{ name: 'FileNumber' }]);
        this.registerBuiltin('loc', (fn: any) => {
            const h = this.fileHandles.get(Number(fn));
            if (!h) this.throwVbaError(VbaErrorCode.BAD_FILE_NAME_OR_NUMBER, "Bad file name or number");
            return h.pos;
        }, [{ name: 'FileNumber' }]);
        this.registerBuiltin('seek', (fn: any) => {
            const h = this.fileHandles.get(Number(fn));
            if (!h) this.throwVbaError(VbaErrorCode.BAD_FILE_NAME_OR_NUMBER, "Bad file name or number");
            return (h.pos || 0) + 1;
        }, [{ name: 'FileNumber' }]);
        this.registerBuiltin('fileattr', (fn: any, info: any = 1) => {
            console.log(`[STUB] FileAttr #${fn}, ${info}`);
            return 1; // Default to Normal
        }, [{ name: 'FileNumber' }, { name: 'ReturnType', optional: true }]);
        this.registerBuiltin('chdrive', (drive: any) => {
            console.log(`[STUB] ChDrive "${drive}"`);
        }, [{ name: 'Drive' }]);
        this.registerBuiltin('setattr', (path: any, attr: any) => {
            console.log(`[STUB] SetAttr "${path}", ${attr}`);
        }, [{ name: 'PathName' }, { name: 'Attributes' }]);
        this.registerBuiltin('filedatetime', (path: any) => {
            const realPath = this.sandbox.toRealPath(String(path));
            const stats = this.fs.statSync(realPath);
            return new VbaDate(toVbaDate(stats.mtime));
        }, [{ name: 'PathName' }]);

        this.registerBuiltin('curdir', (_drive?: string) => this.sandbox.getCwd(), [{ name: 'Drive', optional: true }], ['$']);
        this.registerBuiltin('dir', (pathName?: string, _attributes?: number) => {
            if (pathName !== undefined && pathName !== null && pathName !== "") {
                try {
                    const realPath = this.sandbox.toRealPath(pathName);
                    const dirPath = path.dirname(realPath);
                    const filter = path.basename(realPath);
                    const files = this.fs.readdirSync(dirPath);
                    if (filter === '' || filter === '*' || filter === '*.*') {
                        this.dirIterator = files;
                    } else {
                        const regex = this.vbaWildcardToRegex(filter);
                        this.dirIterator = files.filter((f: string) => regex.test(f));
                    }
                    this.dirIndex = 0;
                } catch (e) {
                    this.dirIterator = [];
                }
            }
            if (!this.dirIterator || this.dirIndex >= this.dirIterator.length) return "";
            return this.dirIterator[this.dirIndex++];
        }, [
            { name: 'PathName', optional: true },
            { name: 'Attributes', optional: true },
        ], ['$']);
        this.registerBuiltin('filecopy', (src: string, dest: string) => this.fs.copyFileSync?.(this.sandbox.toRealPath(src), this.sandbox.toRealPath(dest)), [
            { name: 'Source' }, { name: 'Destination' },
        ]);
        this.registerBuiltin('kill', (p: string) => this.executeKill(p), [{ name: 'PathName' }]);
        this.registerBuiltin('mkdir', (p: string) => this.fs.mkdirSync(this.sandbox.toRealPath(p), { recursive: true }), [{ name: 'Path' }]);
        this.registerBuiltin('rmdir', (p: string) => this.fs.rmdirSync?.(this.sandbox.toRealPath(p)), [{ name: 'Path' }]);
        this.registerBuiltin('chdir', (p: string) => {
            try {
                this.sandbox.setCwd(p);
            } catch {
                this.throwVbaError(VbaErrorCode.PATH_NOT_FOUND, 'Path not found');
            }
        }, [{ name: 'Path' }]);
        this.registerBuiltin('filelen', (p: string) => this.fs.statSync(this.sandbox.toRealPath(p)).size, [{ name: 'PathName' }]);
    }

    private registerInteractionFunctions() {
        this.registerBuiltin('shell', (cmd: any, style: any = 1) => { this.onPrint(`[SHELL] ${cmd} (Style: ${style})`); return 1; }, [
            { name: 'PathName' },
            { name: 'WindowStyle', optional: true },
        ]);
        this.registerBuiltin('msgbox', (msg: any, _buttons: any = 0, _title: any = "") => {
            const title = _title ? ` ${_title}:` : '';
            this.onPrint(`[MSGBOX]${title} ${msg}`);
            return 1;
        }, [
            { name: 'Prompt' },
            { name: 'Buttons', optional: true },
            { name: 'Title', optional: true },
        ]);
        this.registerBuiltin('inputbox', (prompt: any, _title: any = "", def: any = "") => { this.onPrint(`[INPUTBOX] ${prompt}`); return def; }, [
            { name: 'Prompt' },
            { name: 'Title', optional: true },
            { name: 'Default', optional: true },
        ]);
        this.registerBuiltin('appactivate', (title: string, _wait?: boolean) => { this.onPrint(`[APPACTIVATE] ${title}`); }, [
            { name: 'Title' },
            { name: 'Wait', optional: true },
        ]);
        this.registerBuiltin('sendkeys', (keys: string, _wait?: boolean) => { this.onPrint(`[SENDKEYS] ${keys}`); }, [
            { name: 'String' },
            { name: 'Wait', optional: true },
        ]);
        const doEventsFunc = () => 0;
        this.registerBuiltin('doevents', doEventsFunc, []);
    }

    private registerFinancialFunctions() {
        const getRateFactor = (rate: number, nper: number) => {
            if (rate === 0) return nper;
            return (Math.pow(1 + rate, nper) - 1) / rate;
        };
        this.registerBuiltin('fv', (rate: any, nper: any, pmt: any, pv: any = 0, type: any = 0) => {
            const r = Number(rate), n = Number(nper), p = Number(pmt), v = Number(pv), t = Number(type);
            const factor = getRateFactor(r, n);
            const result = -(v * Math.pow(1 + r, n) + p * (1 + r * t) * factor);
            return result;
        }, [
            { name: 'Rate' }, { name: 'NPer' }, { name: 'Pmt' },
            { name: 'PV', optional: true }, { name: 'Type', optional: true },
        ]);
        this.registerBuiltin('pv', (rate: any, nper: any, pmt: any, fv: any = 0, type: any = 0) => {
            const r = Number(rate), n = Number(nper), p = Number(pmt), f = Number(fv), t = Number(type);
            if (r === 0) return -(f + p * n);
            const p1 = Math.pow(1 + r, n);
            const result = -(f + p * (1 + r * t) * ((p1 - 1) / r)) / p1;
            return result;
        }, [
            { name: 'Rate' }, { name: 'NPer' }, { name: 'Pmt' },
            { name: 'FV', optional: true }, { name: 'Type', optional: true },
        ]);
        this.registerBuiltin('pmt', (rate: any, nper: any, pv: any, fv: any = 0, type: any = 0) => {
            const r = Number(rate), n = Number(nper), v = Number(pv), f = Number(fv), t = Number(type);
            if (r === 0) return -(v + f) / n;
            const p1 = Math.pow(1 + r, n);
            const result = -(v * p1 + f) / ((1 + r * t) * ((p1 - 1) / r));
            return result;
        }, [
            { name: 'Rate' }, { name: 'NPer' }, { name: 'PV' },
            { name: 'FV', optional: true }, { name: 'Type', optional: true },
        ]);
        this.registerBuiltin('nper', (rate: any, pmt: any, pv: any, fv: any = 0, type: any = 0) => {
            const r = Number(rate), p = Number(pmt), v = Number(pv), f = Number(fv), t = Number(type);
            if (r === 0) return -(v + f) / p;
            const num = p * (1 + r * t) - f * r;
            const den = p * (1 + r * t) + v * r;
            return Math.log(num / den) / Math.log(1 + r);
        }, [
            { name: 'Rate' }, { name: 'Pmt' }, { name: 'PV' },
            { name: 'FV', optional: true }, { name: 'Type', optional: true },
        ]);
        this.registerBuiltin('rate', (nper: any, pmt: any, pv: any, fv: any = 0, type: any = 0, guess: any = 0.1) => {
            // Newton-Raphson approximation for Rate
            let r = Number(guess);
            const n = Number(nper), p = Number(pmt), v = Number(pv), f = Number(fv), t = Number(type);
            for (let i = 0; i < 20; i++) {
                const p1 = Math.pow(1 + r, n);
                const f_r = v * p1 + p * (1 + r * t) * ((p1 - 1) / r) + f;
                const df_r = v * n * Math.pow(1 + r, n - 1) + p * (t * ((p1 - 1) / r) + (1 + r * t) * (n * Math.pow(1 + r, n - 1) * r - (p1 - 1)) / (r * r));
                const newR = r - f_r / df_r;
                if (Math.abs(newR - r) < 1e-10) return newR;
                r = newR;
            }
            return r;
        }, [
            { name: 'NPer' }, { name: 'Pmt' }, { name: 'PV' },
            { name: 'FV', optional: true }, { name: 'Type', optional: true }, { name: 'Guess', optional: true },
        ]);
        this.registerBuiltin('sln', (cost: any, salvage: any, life: any) => {
            return (Number(cost) - Number(salvage)) / Number(life);
        }, [{ name: 'Cost' }, { name: 'Salvage' }, { name: 'Life' }]);
        this.registerBuiltin('syd', (cost: any, salvage: any, life: any, period: any) => {
            const c = Number(cost), s = Number(salvage), l = Number(life), p = Number(period);
            return ((c - s) * (l - p + 1) * 2) / (l * (l + 1));
        }, [{ name: 'Cost' }, { name: 'Salvage' }, { name: 'Life' }, { name: 'Period' }]);
        this.registerBuiltin('ddb', (cost: any, salvage: any, life: any, period: any, factor: any = 2) => {
            let c = Number(cost), s = Number(salvage), l = Number(life), p = Number(period), f = Number(factor);
            if (p <= 0 || p > l) return 0;
            let book = c;
            let dep = 0;
            for (let i = 1; i <= p; i++) {
                dep = Math.min(book * (f / l), Math.max(0, book - s));
                book -= dep;
            }
            return dep;
        }, [
            { name: 'Cost' }, { name: 'Salvage' }, { name: 'Life' }, { name: 'Period' },
            { name: 'Factor', optional: true },
        ]);
        this.registerBuiltin('irr', (values: any, guess: any = 0.1) => {
            if (!Array.isArray(values)) this.throwVbaError(VbaErrorCode.TYPE_MISMATCH, "Type mismatch");
            const v = values.map(Number);
            let r = Number(guess);
            for (let i = 0; i < 100; i++) {
                let npv = 0;
                let dnpv = 0;
                for (let t = 0; t < v.length; t++) {
                    const p1 = Math.pow(1 + r, t);
                    npv += v[t] / p1;
                    if (t > 0) dnpv -= t * v[t] / (p1 * (1 + r));
                }
                const newR = r - npv / dnpv;
                if (Math.abs(newR - r) < 1e-10) return newR;
                r = newR;
            }
            return r;
        }, [{ name: 'ValueArray' }, { name: 'Guess', optional: true }]);
        this.registerBuiltin('mirr', (values: any, finance_rate: any, reinvest_rate: any) => {
            if (!Array.isArray(values)) this.throwVbaError(VbaErrorCode.TYPE_MISMATCH, "Type mismatch");
            const v = values.map(Number);
            const fr = Number(finance_rate), rr = Number(reinvest_rate);
            const n = v.length - 1;
            let npv_neg = 0;
            let npv_pos = 0;
            for (let t = 0; t < v.length; t++) {
                if (v[t] < 0) npv_neg += v[t] / Math.pow(1 + fr, t);
                else npv_pos += v[t] / Math.pow(1 + rr, t);
            }
            const tv = npv_pos * Math.pow(1 + rr, n);
            return Math.pow(-tv / npv_neg, 1 / n) - 1;
        }, [{ name: 'ValueArray' }, { name: 'FinanceRate' }, { name: 'ReinvestRate' }]);
        this.registerBuiltin('npv', (rate: any, values: any) => {
            if (!Array.isArray(values)) this.throwVbaError(VbaErrorCode.TYPE_MISMATCH, "Type mismatch");
            const r = Number(rate);
            const v = values.map(Number);
            let result = 0;
            for (let i = 0; i < v.length; i++) {
                result += v[i] / Math.pow(1 + r, i + 1);
            }
            return result;
        }, [{ name: 'Rate' }, { name: 'ValueArray' }]);
        this.registerBuiltin('ipmt', (rate: any, per: any, nper: any, pv: any, fv: any = 0, type: any = 0) => {
            const r = Number(rate), p = Number(per), n = Number(nper), v = Number(pv), f = Number(fv), t = Number(type);
            const pmt = Number(this.env.get('pmt')(r, n, v, f, t));
            let ipmt: number;
            if (p === 1) {
                ipmt = t === 1 ? 0 : -v * r;
            } else {
                const fv_prev = Number(this.env.get('fv')(r, p - 1, pmt, v, t));
                ipmt = -fv_prev * r;
            }
            return ipmt;
        }, [
            { name: 'Rate' }, { name: 'Per' }, { name: 'NPer' }, { name: 'PV' },
            { name: 'FV', optional: true }, { name: 'Type', optional: true },
        ]);
        this.registerBuiltin('ppmt', (rate: any, per: any, nper: any, pv: any, fv: any = 0, type: any = 0) => {
            const r = Number(rate), p = Number(per), n = Number(nper), v = Number(pv), f = Number(fv), t = Number(type);
            const pmt = Number(this.env.get('pmt')(r, n, v, f, t));
            const ipmt = Number(this.env.get('ipmt')(r, p, n, v, f, t));
            return pmt - ipmt;
        }, [
            { name: 'Rate' }, { name: 'Per' }, { name: 'NPer' }, { name: 'PV' },
            { name: 'FV', optional: true }, { name: 'Type', optional: true },
        ]);
    }

    private registerConstants() {
        const errorMessages: Record<number, string> = { 5: "Invalid procedure call or argument", 6: "Overflow", 9: "Subscript out of range", 11: "Division by zero", 13: "Type mismatch", 52: "Bad file name or number", 53: "File not found", 58: "File already exists", 62: "Input past end of file", 70: "Permission denied", 76: "Path not found", 91: "Object variable not set", 94: "Invalid use of Null" };
        const errFunc = (n?: any) => errorMessages[n === undefined ? this.errObj.number : Number(n)] || "Application-defined or object-defined error";
        this.registerBuiltin('error', errFunc, [{ name: 'ErrorNumber', optional: true }], ['$']);
        this.env.setConstant('vbsunday', 1);
        this.env.setConstant('vbmonday', 2);
        this.env.setConstant('vbtuesday', 3);
        this.env.setConstant('vbwednesday', 4);
        this.env.setConstant('vbthursday', 5);
        this.env.setConstant('vbfriday', 6);
        this.env.setConstant('vbsaturday', 7);
        this.env.setConstant('vbusesystem', 0);
        this.env.setConstant('vbbinarycompare', 0);
        this.env.setConstant('vbtextcompare', 1);
        this.env.setConstant('vbempty', 0); this.env.setConstant('vbnull', 1); this.env.setConstant('vbinteger', 2); this.env.setConstant('vblong', 3); this.env.setConstant('vbsingle', 4); this.env.setConstant('vbdouble', 5); this.env.setConstant('vbcurrency', 6); this.env.setConstant('vbdate', 7); this.env.setConstant('vbstring', 8); this.env.setConstant('vbobject', 9); this.env.setConstant('vberror', 10); this.env.setConstant('vbboolean', 11); this.env.setConstant('vbvariant', 12); this.env.setConstant('vbbyte', 17); this.env.setConstant('vblonglong', 20); this.env.setConstant('vbarray', 8192);
        this.env.setConstant('vbcrlf', "\r\n"); this.env.setConstant('vbtab', "\t"); this.env.setConstant('vbcr', "\r"); this.env.setConstant('vblf', "\n"); this.env.setConstant('vbnewline', "\n"); this.env.setConstant('vbnullstring', ''); this.env.setConstant('vbnullchar', '\0'); this.env.setConstant('vbback', "\b"); this.env.setConstant('vbformfeed', "\f");
        // カスタムエラー番号の基準値（MS-VBAL §6.1.2.10）。Err.Raise vbObjectError + n の慣用句で使う。
        this.env.setConstant('vbobjecterror', -2147221504);
        this.env.setConstant('true', vbaTrue); this.env.setConstant('false', vbaFalse); this.env.setConstant('empty', vbaEmpty); this.env.setConstant('nothing', vbaNothing); this.env.setConstant('null', vbaNull);

        this.registerBuiltin('environ', (k: any) => this.sandbox.getEnv(k), [{ name: 'EnvString' }], ['$']);

        // VarPtr / StrPtr / ObjPtr — dummy pointer functions (non-zero unique Long per call)
        // Real addresses have no meaning in this engine; return stable non-zero integers.
        const ptrFn = () => (this._ptrCounter += 4);
        this.registerBuiltin('varptr', ptrFn, [{ name: 'VarName' }]);
        this.registerBuiltin('strptr', ptrFn, [{ name: 'VarName' }]);
        this.registerBuiltin('objptr', ptrFn, [{ name: 'VarName' }]);
        this.registerBuiltin('createobject', (id: string) => this.createExternalObject(id), [{ name: 'Class' }]);
        const getObjectFunc = (pathname?: string, classId?: string) => {
            if (pathname) {
                return {
                    __vbaTypeName__: 'Object',
                    Path: pathname,
                    Name: pathname.split(/[\\\/]/).pop()
                };
            }
            if (classId) {
                return this.createExternalObject(classId);
            }
            return vbaNothing;
        };
        this.registerBuiltin('getobject', getObjectFunc, [
            { name: 'PathName', optional: true },
            { name: 'Class', optional: true },
        ]);
        this.registerBuiltin('iif', (c: any, t: any, f: any) => this.isTrue(c) ? t : f, [
            { name: 'Expr' }, { name: 'TruePart' }, { name: 'FalsePart' },
        ]);
        this.registerBuiltin('choose', (i: any, ...c: any[]) => { const idx = Math.floor(Number(i)); return (idx >= 1 && idx <= c.length) ? c[idx - 1] : vbaNull; }, [
            { name: 'Index' },
            { name: 'Choice', isParamArray: true },
        ]);
        // Switch(Expr1, Value1, [Expr2, Value2]...) はペア単位の ParamArray のため、
        // 個数検証は行わず今までどおり可変長で受け取る（isParamArray で checkArgCountGeneric をスキップ）。
        this.registerBuiltin('switch', (...args: any[]) => { for (let i = 0; i < args.length; i += 2) if (this.isTrue(args[i])) return args[i + 1]; return vbaNull; }, [
            { name: 'VarExpr', isParamArray: true },
        ]);
        this.registerBuiltin('array', (...args: any[]) => { const a = [...args]; (a as any).vbaBase = 0; return a; }, [
            { name: 'Arglist', isParamArray: true },
        ]);
        this.registerBuiltin('lbound', (a: any, dim: any = 1) => {
            if (!Array.isArray(a)) this.throwVbaError(VbaErrorCode.SUBSCRIPT_OUT_OF_RANGE, "Subscript out of range");
            const dimIndex = Number(dim) - 1;
            if ((a as any).__vbaDimensions__) {
                if (dimIndex < 0 || dimIndex >= (a as any).__vbaDimensions__.length) {
                    this.throwVbaError(VbaErrorCode.SUBSCRIPT_OUT_OF_RANGE, "Subscript out of range");
                }
                return (a as any).__vbaDimensions__[dimIndex].lower;
            }
            if (dimIndex > 0) this.throwVbaError(VbaErrorCode.SUBSCRIPT_OUT_OF_RANGE, "Subscript out of range");
            return (a as any).vbaBase || 0;
        }, [{ name: 'ArrayName' }, { name: 'Dimension', optional: true }]);
        this.registerBuiltin('ubound', (a: any, dim: any = 1) => {
            if (!Array.isArray(a)) this.throwVbaError(VbaErrorCode.SUBSCRIPT_OUT_OF_RANGE, "Subscript out of range");
            const dimIndex = Number(dim) - 1;
            if ((a as any).__vbaDimensions__) {
                if (dimIndex < 0 || dimIndex >= (a as any).__vbaDimensions__.length) {
                    this.throwVbaError(VbaErrorCode.SUBSCRIPT_OUT_OF_RANGE, "Subscript out of range");
                }
                return (a as any).__vbaDimensions__[dimIndex].upper;
            }
            if (dimIndex > 0) this.throwVbaError(VbaErrorCode.SUBSCRIPT_OUT_OF_RANGE, "Subscript out of range");
            return ((a as any).vbaBase || 0) + a.length - 1;
        }, [{ name: 'ArrayName' }, { name: 'Dimension', optional: true }]);
    }

    private registerRegistryFunctions() {
        this.registerBuiltin('savesetting', (app: string, sec: string, key: string, val: any) => {
            if (!this.virtualRegistry[app]) this.virtualRegistry[app] = {};
            if (!this.virtualRegistry[app][sec]) this.virtualRegistry[app][sec] = {};
            this.virtualRegistry[app][sec][key] = String(val);
        }, [
            { name: 'AppName' }, { name: 'Section' }, { name: 'Key' }, { name: 'Setting' },
        ]);
        this.registerBuiltin('getsetting', (app: string, sec: string, key: string, def: any = "") => {
            return (this.virtualRegistry[app]?.[sec]?.[key]) ?? String(def);
        }, [
            { name: 'AppName' }, { name: 'Section' }, { name: 'Key' }, { name: 'Default', optional: true },
        ]);
        this.registerBuiltin('getallsettings', (app: string, sec: string) => {
            const s = this.virtualRegistry[app]?.[sec];
            if (!s) return vbaEmpty;
            const res = Object.entries(s).map(([k, v]) => [k, v]);
            (res as any).vbaBase = 0;
            return res;
        }, [{ name: 'AppName' }, { name: 'Section' }]);
        this.registerBuiltin('deletesetting', (app: string, sec?: string, key?: string) => {
            if (!sec) delete this.virtualRegistry[app];
            else if (!key) delete this.virtualRegistry[app]?.[sec];
            else delete this.virtualRegistry[app]?.[sec]?.[key];
        }, [
            { name: 'AppName' },
            { name: 'Section', optional: true },
            { name: 'Key', optional: true },
        ]);
    }



    private triggerTerminate(obj: any) {
        if (obj && obj.__vbaClass__) {
            // VBA spec: Class_Terminate executes at most once during the lifetime of an object
            if (obj.__terminateCalled__) {
                return;
            }
            obj.__terminateCalled__ = true;

            const classDef = obj.__classDef__ as ClassDeclaration;
            const terminateProc = classDef.procedures.find(p => p.name.name.toLowerCase() === 'class_terminate');
            if (terminateProc) {
                try {
                    this.callClassMethod(obj, terminateProc, []);
                } catch (e) {
                    // Errors in Terminate are typically suppressed or handled specially in VBA,
                    // but for now let's just log or ignore to prevent crashing the lifecycle
                    console.error("Error in Class_Terminate:", e);
                }
            }
        }
    }

    private instantiateType(typeName: string): any {
        const typeMembers = this.env.getType(typeName);
        if (!typeMembers) return 0;

        const instance: any = {
            __vbaTypeName__: typeName
        };
        for (const member of typeMembers) {
            const mt = member.memberType;
            const mtLower = mt.toLowerCase();
            if (mtLower === 'string') {
                instance[member.name.toLowerCase()] = '';
            } else if (mtLower === 'boolean') {
                instance[member.name.toLowerCase()] = 0; // vbaFalse
            } else if (this.env.getType(mt)) {
                instance[member.name.toLowerCase()] = this.instantiateType(mt);
            } else {
                instance[member.name.toLowerCase()] = 0;
            }
        }
        return instance;
    }

    public get(name: string): any {
        return this.env.get(name);
    }

    public set(name: string, value: any): void {
        this.env.set(name, value);
    }

    /**
     * §5.6.10 Tier 6 のデフォルト束縛オブジェクトを設定する。
     * Tier 1〜4 で解決できなかった修飾なし識別子を、このオブジェクトのプロパティ/メソッドとして検索する。
     * Excel VBA の Application 相当。MockApplication を渡すと Range("A1") / Cells(1,1) / ActiveSheet 等が動作する。
     * null を渡すと解除する。
     */
    public setDefaultBindingObject(obj: any): void {
        this.defaultBindingObject = obj;
    }

    /** 名前をグローバル定数として登録する。VBA コード側からの代入は Error 5 になる。 */
    public setConstant(name: string, value: any): void {
        this.env.setConstant(name, value);
    }

    /** precheckProc のみを実行し、本体は実行しない（assertCompileErrorPreproc 用）。 */
    public checkProcedure(name: string): void {
        const procName = name.toLowerCase();
        const proc = this.env.getProcedure(procName);
        if (!proc) throw new Error(`Procedure '${name}' not found`);
        this.precheckProc(proc);
    }

    public callProcedure(name: string, args: any[], type?: 'get' | 'let' | 'set', moduleName?: string): any {
        let procName = name.toLowerCase();
        let extractedModuleName = moduleName;

        // Check if name contains module qualifier (e.g., "ModuleName.ProcedureName")
        if (!extractedModuleName && procName.includes('.')) {
            const parts = procName.split('.');
            if (parts.length === 2) {
                extractedModuleName = parts[0];
                procName = parts[1];
            }
        }

        let proc: ProcedureDeclaration | undefined;

        // Try module-qualified lookup first if module name is specified
        if (extractedModuleName) {
            proc = this.env.getProcedureFromModule(procName, extractedModuleName, type);
        } else {
            // Fall back to non-qualified lookup
            proc = this.env.getProcedure(procName, type);
        }

        if (!proc) {
            // Fall back to built-in functions stored as closures in env
            const builtin = this.env.get(procName);
            if (typeof builtin === 'function') {
                return builtin(...args);
            }
            this.throwVbaError(VbaErrorCode.SUB_OR_FUNCTION_NOT_DEFINED, `Sub or Function not defined: '${name}'${extractedModuleName ? ` in module '${extractedModuleName}'` : ''}`);
        }

        if (this.vbaCallStack.length === 0) this.currentLine = 0;
        this.vbaCallStack.push({ name: proc.name.name, moduleName: proc.moduleName ?? '', line: this.currentLine });
        try {

        this.precheckProc(proc);

        // Validate argument count
        this.checkArgCount(proc, args);

        // Create a new local environment: Tier 1, enclosing = Tier 2 (moduleEnv)
        const procModuleKey = (proc.moduleName ?? '').toLowerCase();
        const parentEnv = this.moduleEnvs.get(procModuleKey) ?? this.env;
        const localEnv = new Environment(parentEnv);

        // Map arguments to parameter names
        for (let i = 0; i < proc.parameters.length; i++) {
            const param = proc.parameters[i];
            const paramName = param.name;

            if (param.isParamArray) {
                // Capture all remaining arguments into a Variant array
                const remainingArgs = args.slice(i);
                (remainingArgs as any).vbaBase = 0; // ParamArray is always 0-based
                localEnv.setLocally(paramName, remainingArgs);
                break;
            }

            let argValue: any;
            if (i < args.length) {
                argValue = args[i];
            } else if (param.defaultValue) {
                argValue = this.evaluateExpression(param.defaultValue);
            } else {
                // checkArgCount済みかつパラメーター順序（必須が先・Optionalが後）が保証されているため、
                // ここに来る時点で param は必ず Optional（defaultValue なし）。
                argValue = vbaMissing;
            }
            // Register parameter type metadata (but not for array parameters)
            if (param.paramType && !param.isArray) {
                const typeMap: Record<string, VbaVarType> = {
                    'byte': 'Byte', 'integer': 'Integer', 'long': 'Long',
                    'single': 'Single', 'double': 'Double', 'currency': 'Currency',
                    'string': 'String', 'boolean': 'Boolean', 'date': 'Date',
                };
                const mapped = typeMap[param.paramType.toLowerCase()];
                if (mapped) {
                    localEnv.setVariableType(paramName, { vbaType: mapped });
                }
            }
            localEnv.setLocally(paramName, argValue);
        }

        const result = this.execProcBody(proc, localEnv, {
            byRefArgs: [],
            paramArrayParamName: null,
            paramArrayByRefExprs: [],
            initialLastErrorIndex: null,
            acceptExitProperty: true,
            returnOnProperty: true,
            subReturnValue: vbaEmpty,
        });

        // ByRef パラメーターの値を呼び出し元の args 配列へ書き戻す。
        // VBARunner.run() のように TypeScript から直接 callProcedure を呼ぶ場合、
        // JS の数値・文字列・Boolean はプリミティブで参照を共有しないため、
        // ByRef パラメーターへの代入（Sub 内の `n = n + 1` 等）が呼び出し元の
        // args 配列に反映されない問題があった。VBA の既定の引数渡しは ByRef
        // （明示的な ByVal がない限り）なので、ByVal 以外のパラメーターは
        // 呼び出し後の最終値を args[i] に書き戻す。
        for (let i = 0; i < proc.parameters.length; i++) {
            const param = proc.parameters[i];
            if (param.isParamArray || param.isByVal || i >= args.length) continue;
            args[i] = localEnv.get(param.name.toLowerCase());
        }

        return result;

        } finally {
            this.vbaCallStack.pop();
        }
    }

    /** Tier 2: モジュール専用 Environment を取得または生成する。enclosing = Tier 3 (this.env) */
    private getOrCreateModuleEnv(moduleName: string): Environment {
        const key = (moduleName || 'module1').toLowerCase();
        if (!this.moduleEnvs.has(key)) {
            this.moduleEnvs.set(key, new Environment(this.env));
        }
        return this.moduleEnvs.get(key)!;
    }

    // OE check shared between callProcedure and evaluateCallExpression.
    // §5.6.10 Tier 6: names resolvable via defaultBindingObject are not implicit variables.
    private precheckProc(proc: ProcedureDeclaration): void {
        try {
            const procKey = proc.name.name.toLowerCase();
            if (this.optionExplicitViolations.has(procKey)) {
                const violations = this.optionExplicitViolations.get(procKey)!;
                const stillMissing = [...violations.entries()].filter(([n]) => {
                    if (this.env.hasVariable(n)) return false;
                    if (this.typeLibraryNamespaces.has(n)) return false;
                    if (this.defaultBindingObject &&
                            this.resolveObjectMemberKey(this.defaultBindingObject, n) !== undefined) return false;
                    return true;
                });
                if (stillMissing.length > 0) {
                    const names = stillMissing.map(([n]) => n).join(', ');
                    const firstLine = stillMissing[0][1] || undefined;
                    this.throwCompileError(VbaErrorCode.OPTION_EXPLICIT_VIOLATION,
                        `Variable not declared in '${proc.name.name}' (Option Explicit): ${names}`,
                        firstLine, proc.moduleName ?? undefined);
                }
            }
            this.checkSubAsValueInProc(proc);
            this.checkUndefinedCallsInProc(proc);
            this.checkConstantArrayBoundsInProc(proc);
            this.checkDuplicateDimInProc(proc);
            this.checkGoToLabelsInProc(proc);
            this.checkCallArgCountsInProc(proc);
        } catch (e: any) {
            if (e._precheckRaw) {
                const line = e.vbaLine;
                const mod = e.vbaModule ?? this.executingModuleName ?? this.currentSourceModule ?? null;
                const msg = line !== undefined ? `Compile error: ${e.message} (line ${line})` : `Compile error: ${e.message}`;
                const formatted: any = new Error(msg);
                formatted.type = 'VbaError';
                formatted.number = e.number;
                formatted.vbaLine = line;
                formatted.vbaModule = mod;
                formatted.vbaStack = [...this.vbaCallStack].reverse();
                throw formatted;
            }
            throw e;
        }
    }

    private checkUndefinedCallsInProc(proc: ProcedureDeclaration): void {
        const knownNames = this.env.collectAllNames();
        // moduleEnvs (Tier 2) の名前も既知として追加
        for (const mEnv of this.moduleEnvs.values()) {
            for (const n of mEnv.collectAllNames()) knownNames.add(n);
        }
        // builtinEnv の名前も追加（collectAllNames は chain を辿るが builtinEnv は enclosing なので含まれる）
        // defaultBindingObject のメンバーを既知として追加（Sheets, Range 等）
        if (this.defaultBindingObject) {
            let cur = this.defaultBindingObject;
            while (cur && cur !== Object.prototype && cur !== Function.prototype) {
                for (const k of Object.getOwnPropertyNames(cur)) knownNames.add(k.toLowerCase());
                cur = Object.getPrototypeOf(cur);
            }
        }
        const errs: UndefinedProcError[] = [];
        walkProcForUndefinedCalls(proc, new Set(), knownNames, errs);
        if (errs.length > 0) {
            this.throwCompileError(VbaErrorCode.SUB_OR_FUNCTION_NOT_DEFINED,
                `Sub or Function not defined: '${errs[0].name}'`,
                errs[0].line || undefined,
                proc.moduleName ?? undefined);
        }
    }

    // Walk proc body and throw compile error if a Sub is used in value context (e.g. v = MySub).
    private checkSubAsValueInProc(proc: ProcedureDeclaration): void {
        const walkStmts = (stmts: Statement[]) => { for (const s of stmts) walkStmt(s); };
        const walkStmt = (stmt: Statement) => {
            switch (stmt.type) {
                case 'AssignmentStatement': {
                    const assign = stmt as AssignmentStatement;
                    const name = this.subNameInValueExpr(assign.right);
                    if (name !== null) {
                        const p = this.env.getProcedure(name);
                        if (p && !p.isFunction && !p.isProperty) {
                            const line = (assign.right as any).loc?.start.line ?? undefined;
                            this.throwCompileError(VbaErrorCode.TYPE_MISMATCH,
                                `Function or variable expected: '${name}'`,
                                line, proc.moduleName ?? undefined);
                        }
                    }
                    break;
                }
                case 'IfStatement': {
                    const s = stmt as IfStatement;
                    walkStmts(s.consequent);
                    if (Array.isArray(s.alternate)) walkStmts(s.alternate);
                    else if (s.alternate) walkStmt(s.alternate);
                    break;
                }
                case 'ForStatement':     walkStmts((stmt as ForStatement).body); break;
                case 'ForEachStatement': walkStmts((stmt as ForEachStatement).body); break;
                case 'DoWhileStatement': walkStmts((stmt as DoWhileStatement).body); break;
                case 'WhileStatement':   walkStmts((stmt as WhileStatement).body); break;
                case 'WithStatement':    walkStmts((stmt as WithStatement).body); break;
                case 'SelectCaseStatement': {
                    const s = stmt as SelectCaseStatement;
                    for (const c of s.cases) walkStmts(c.body);
                    if (s.elseBody) walkStmts(s.elseBody);
                    break;
                }
            }
        };
        walkStmts(proc.body);
    }

    private checkDuplicateDimInProc(proc: ProcedureDeclaration): void {
        const seen = new Map<string, number>(); // varKey → first-occurrence line
        const walkStmts = (stmts: Statement[]) => { for (const s of stmts) walkStmt(s); };
        const walkStmt = (stmt: Statement) => {
            switch (stmt.type) {
                case 'VariableDeclaration': {
                    const decl = stmt as VariableDeclaration;
                    for (const d of decl.declarations) {
                        const key = d.name.name.toLowerCase();
                        if (seen.has(key)) {
                            const line = d.name.loc?.start.line;
                            this.throwCompileError(VbaErrorCode.INVALID_PROCEDURE_CALL,
                                `Variable '${d.name.name}' is already declared in this scope (duplicate declaration)`,
                                line, proc.moduleName ?? undefined);
                        }
                        seen.set(key, d.name.loc?.start.line ?? 0);
                    }
                    break;
                }
                case 'IfStatement': {
                    const s = stmt as IfStatement;
                    walkStmts(s.consequent);
                    if (Array.isArray(s.alternate)) walkStmts(s.alternate);
                    else if (s.alternate) walkStmt(s.alternate);
                    break;
                }
                case 'ForStatement':        walkStmts((stmt as ForStatement).body); break;
                case 'ForEachStatement':    walkStmts((stmt as ForEachStatement).body); break;
                case 'DoWhileStatement':    walkStmts((stmt as DoWhileStatement).body); break;
                case 'WhileStatement':      walkStmts((stmt as WhileStatement).body); break;
                case 'WithStatement':       walkStmts((stmt as WithStatement).body); break;
                case 'SelectCaseStatement': {
                    const s = stmt as SelectCaseStatement;
                    for (const c of s.cases) walkStmts(c.body);
                    if (s.elseBody) walkStmts(s.elseBody);
                    break;
                }
            }
        };
        walkStmts(proc.body);
    }

    private checkGoToLabelsInProc(proc: ProcedureDeclaration): void {
        const labels = new Set<string>();
        const gotos: Array<{ label: string; line: number }> = [];

        const walkStmts = (stmts: Statement[]) => { for (const s of stmts) walkStmt(s); };
        const walkStmt = (stmt: Statement) => {
            switch (stmt.type) {
                case 'LabelStatement':
                    labels.add((stmt as any).label.toLowerCase());
                    break;
                case 'GoToStatement':
                    gotos.push({ label: (stmt as GoToStatement).label, line: stmt.loc?.start.line ?? 0 });
                    break;
                case 'GoSubStatement':
                    gotos.push({ label: (stmt as GoSubStatement).label, line: stmt.loc?.start.line ?? 0 });
                    break;
                case 'OnGoToSubStatement':
                    for (const lbl of (stmt as OnGoToSubStatement).labels)
                        gotos.push({ label: lbl, line: stmt.loc?.start.line ?? 0 });
                    break;
                case 'IfStatement': {
                    const s = stmt as IfStatement;
                    walkStmts(s.consequent);
                    if (Array.isArray(s.alternate)) walkStmts(s.alternate);
                    else if (s.alternate) walkStmt(s.alternate);
                    break;
                }
                case 'ForStatement':        walkStmts((stmt as ForStatement).body); break;
                case 'ForEachStatement':    walkStmts((stmt as ForEachStatement).body); break;
                case 'DoWhileStatement':    walkStmts((stmt as DoWhileStatement).body); break;
                case 'WhileStatement':      walkStmts((stmt as WhileStatement).body); break;
                case 'WithStatement':       walkStmts((stmt as WithStatement).body); break;
                case 'SelectCaseStatement': {
                    const s = stmt as SelectCaseStatement;
                    for (const c of s.cases) walkStmts(c.body);
                    if (s.elseBody) walkStmts(s.elseBody);
                    break;
                }
            }
        };
        walkStmts(proc.body);

        for (const g of gotos) {
            if (!labels.has(g.label.toLowerCase())) {
                this.throwCompileError(VbaErrorCode.SUB_OR_FUNCTION_NOT_DEFINED,
                    `Label not defined: '${g.label}'`,
                    g.line || undefined, proc.moduleName ?? undefined);
            }
        }
    }

    private checkCallArgCountsInProc(proc: ProcedureDeclaration): void {
        const checkCall = (ce: CallExpression): void => {
            if (ce.callee.type !== 'Identifier') return;
            const name = (ce.callee as Identifier).name.toLowerCase();
            const target = this.env.getProcedure(name);
            if (!target) return;
            const hasParamArray = target.parameters.some(p => p.isParamArray);
            if (hasParamArray) return;
            const maxParams = target.parameters.length;
            const minParams = target.parameters.filter(p => !p.isOptional && p.defaultValue == null).length;
            const argCount = ce.args.length;
            const line = ce.loc?.start.line;
            if (argCount > maxParams) {
                this.throwCompileError(VbaErrorCode.WRONG_NUMBER_OF_ARGUMENTS,
                    'Wrong number of arguments or invalid property assignment',
                    line, proc.moduleName ?? undefined);
            }
            if (argCount < minParams) {
                this.throwCompileError(VbaErrorCode.ARGUMENT_NOT_OPTIONAL,
                    'Argument not optional',
                    line, proc.moduleName ?? undefined);
            }
        };

        const visitExpr = (expr: Expression): void => {
            switch (expr.type) {
                case 'CallExpression': {
                    const ce = expr as CallExpression;
                    checkCall(ce);
                    for (const a of ce.args) visitExpr(a);
                    break;
                }
                case 'BinaryExpression': {
                    const b = expr as BinaryExpression;
                    visitExpr(b.left); visitExpr(b.right);
                    break;
                }
                case 'UnaryExpression': visitExpr((expr as UnaryExpression).argument); break;
                case 'ParenthesizedExpression': visitExpr((expr as ParenthesizedExpression).expression); break;
                case 'MemberExpression': visitExpr((expr as MemberExpression).object); break;
            }
        };

        const walkStmts = (stmts: Statement[]) => { for (const s of stmts) walkStmt(s); };
        const walkStmt = (stmt: Statement) => {
            switch (stmt.type) {
                case 'CallStatement':      visitExpr((stmt as CallStatement).expression); break;
                case 'AssignmentStatement': {
                    const a = stmt as AssignmentStatement;
                    visitExpr(a.left); visitExpr(a.right);
                    break;
                }
                case 'SetStatement': {
                    const s = stmt as SetStatement;
                    visitExpr(s.left); visitExpr(s.right);
                    break;
                }
                case 'IfStatement': {
                    const s = stmt as IfStatement;
                    visitExpr(s.condition);
                    walkStmts(s.consequent);
                    if (Array.isArray(s.alternate)) walkStmts(s.alternate);
                    else if (s.alternate) walkStmt(s.alternate);
                    break;
                }
                case 'ForStatement': {
                    const f = stmt as ForStatement;
                    visitExpr(f.start); visitExpr(f.end);
                    if (f.step) visitExpr(f.step);
                    walkStmts(f.body);
                    break;
                }
                case 'ForEachStatement': {
                    const f = stmt as ForEachStatement;
                    visitExpr(f.collection); walkStmts(f.body);
                    break;
                }
                case 'DoWhileStatement': {
                    const d = stmt as DoWhileStatement;
                    if (d.condition) visitExpr(d.condition);
                    walkStmts(d.body);
                    break;
                }
                case 'WhileStatement': {
                    const w = stmt as WhileStatement;
                    visitExpr(w.condition); walkStmts(w.body);
                    break;
                }
                case 'WithStatement': {
                    const w = stmt as WithStatement;
                    visitExpr(w.expression); walkStmts(w.body);
                    break;
                }
                case 'SelectCaseStatement': {
                    const s = stmt as SelectCaseStatement;
                    visitExpr(s.expression);
                    for (const c of s.cases) {
                        for (const r of c.ranges) visitExpr('value' in r ? (r as any).value : ((r as any).start));
                        walkStmts(c.body);
                    }
                    if (s.elseBody) walkStmts(s.elseBody);
                    break;
                }
            }
        };
        walkStmts(proc.body);
    }

    private checkConstantArrayBoundsInProc(proc: ProcedureDeclaration): void {
        const walkStmts = (stmts: Statement[]) => { for (const s of stmts) walkStmt(s); };
        const walkStmt = (stmt: Statement) => {
            switch (stmt.type) {
                case 'VariableDeclaration': {
                    const decl = stmt as VariableDeclaration;
                    for (const d of decl.declarations) {
                        if (!d.isArray || !d.arrayBounds || d.arrayBounds.length === 0) continue;
                        for (const bound of d.arrayBounds) {
                            this.validateConstantExpr(bound.upper);
                            if (bound.lower) this.validateConstantExpr(bound.lower);
                        }
                    }
                    break;
                }
                case 'IfStatement': {
                    const s = stmt as IfStatement;
                    walkStmts(s.consequent);
                    if (Array.isArray(s.alternate)) walkStmts(s.alternate);
                    else if (s.alternate) walkStmt(s.alternate);
                    break;
                }
                case 'ForStatement':        walkStmts((stmt as ForStatement).body); break;
                case 'ForEachStatement':    walkStmts((stmt as ForEachStatement).body); break;
                case 'DoWhileStatement':    walkStmts((stmt as DoWhileStatement).body); break;
                case 'WhileStatement':      walkStmts((stmt as WhileStatement).body); break;
                case 'WithStatement':       walkStmts((stmt as WithStatement).body); break;
                case 'SelectCaseStatement': {
                    const s = stmt as SelectCaseStatement;
                    for (const c of s.cases) walkStmts(c.body);
                    if (s.elseBody) walkStmts(s.elseBody);
                    break;
                }
            }
        };
        walkStmts(proc.body);
    }

    private subNameInValueExpr(expr: Expression): string | null {
        if (expr.type === 'Identifier') return (expr as Identifier).name;
        if (expr.type === 'CallExpression') {
            const ce = expr as CallExpression;
            if (ce.callee.type === 'Identifier') return (ce.callee as Identifier).name;
        }
        return null;
    }

    // Shared procedure body execution.
    // Caller must push vbaCallStack before calling and pop in its own finally.
    private execProcBody(proc: ProcedureDeclaration, localEnv: Environment, opts: ExecProcBodyOptions): any {
        const procNameKey = proc.name.name.toLowerCase();

        // Initialize return variable for function/property
        if (proc.isFunction || proc.isProperty) {
            localEnv.setLocally(proc.name.name, vbaEmpty);
            // 配列を返す関数（As String() 等）は戻り値変数をスカラー型として
            // coerce してはいけない（CStr() 等が配列に適用され壊れるため）
            if (proc.returnType && !proc.returnsArray) {
                const retTypeMap: Record<string, VbaVarType> = {
                    'byte': 'Byte', 'integer': 'Integer', 'long': 'Long',
                    'single': 'Single', 'double': 'Double', 'currency': 'Currency',
                    'longlong': 'LongLong', 'longptr': 'LongPtr',
                    'string': 'String', 'boolean': 'Boolean', 'date': 'Date',
                };
                const mapped = retTypeMap[proc.returnType.toLowerCase()];
                if (mapped) localEnv.setVariableType(proc.name.name, { vbaType: mapped });
            }
        }

        const previousEnv = this.env;
        const previousErrorHandler = this.errorHandlerLabel;
        const previousErrorHandlingMode = this.errorHandlingMode;
        const previousIsInErrorHandler = this.isInErrorHandler;
        const previousLastErrorIndex = this.lastErrorIndex;
        const previousProcBody = this.currentProcBody;
        const previousProcedureName = this.currentProcedureName;
        const previousProcedureType = this.currentProcedureType;
        const previousExecutingModule = this.executingModuleName;
        const previousProcIsStatic = this.currentProcIsStatic;
        const previousStaticVars = this.staticVarsInCurrentProc;
        const previousNewOwned = this._currentNewOwned;

        this.env = localEnv;
        this.errorHandlerLabel = null;
        this.errorHandlingMode = 'None';
        this.isInErrorHandler = false;
        this.lastErrorIndex = opts.initialLastErrorIndex;
        this.currentProcBody = proc.body;
        this.currentProcedureName = proc.name.name;
        this.currentProcedureType = proc.propertyType || (proc.isFunction ? 'function' : 'sub');
        this.executingModuleName = proc.moduleName ?? '';
        this.currentProcIsStatic = proc.isStatic ?? false;
        this.staticVarsInCurrentProc = new Set();
        this._currentNewOwned = new Set();

        try {
            this.executeStatements(proc.body, 0);
        } catch (e: any) {
            if (e && e.type === 'Exit' && (
                e.target === 'Sub' ||
                e.target === 'Function' ||
                (opts.acceptExitProperty && e.target === 'Property')
            )) {
                // valid exit, swallowed
            } else {
                throw e;
            }
        } finally {
            this.env = previousEnv;

            // ByRef writeback (evaluateCallExpression only; empty for callProcedure)
            for (const ref of opts.byRefArgs) {
                const updatedVal = localEnv.get(ref.paramName);
                try {
                    this.evaluateAssignmentToVariable(ref.originalExpr, updatedVal);
                } catch {
                    // r-value: silently ignored
                }
            }
            if (opts.paramArrayParamName !== null) {
                const updatedArray = localEnv.get(opts.paramArrayParamName) as any[];
                if (Array.isArray(updatedArray)) {
                    for (let j = 0; j < opts.paramArrayByRefExprs.length; j++) {
                        try {
                            this.evaluateAssignmentToVariable(opts.paramArrayByRefExprs[j], updatedArray[j]);
                        } catch { }
                    }
                }
            }

            for (const varName of this.staticVarsInCurrentProc) {
                const key = `${procNameKey}:${varName}`;
                this.staticVarStore.set(key, localEnv.get(varName));
            }

            // Scope exit: fire Class_Terminate for local VBA objects that were created with
            // `New` in this scope and not returned. Objects fetched from outer scopes, dicts,
            // or function calls are NOT tracked here and are safe from premature termination.
            {
                const owned = this._currentNewOwned;
                if (owned && owned.size > 0) {
                    const paramNames = new Set(proc.parameters.map(p => p.name.toLowerCase()));
                    const returnVarName = (proc.isFunction || proc.isProperty) ? proc.name.name.toLowerCase() : null;
                    const returnVal = returnVarName !== null ? localEnv.get(returnVarName) : undefined;
                    for (const [name, val] of localEnv.getLocalVariables()) {
                        if (paramNames.has(name)) continue;
                        if (returnVarName !== null && name === returnVarName) continue;
                        if (this.staticVarsInCurrentProc.has(name)) continue;
                        if (val !== null && val !== undefined && val === returnVal) continue;
                        if (!owned.has(val)) continue;
                        this.triggerTerminate(val);
                    }
                }
            }

            this.errorHandlerLabel = previousErrorHandler;
            this.errorHandlingMode = previousErrorHandlingMode;
            this.isInErrorHandler = previousIsInErrorHandler;
            this.lastErrorIndex = previousLastErrorIndex;
            this.currentProcBody = previousProcBody;
            this.currentProcedureName = previousProcedureName;
            this.currentProcedureType = previousProcedureType;
            this.executingModuleName = previousExecutingModule;
            this.currentProcIsStatic = previousProcIsStatic;
            this.staticVarsInCurrentProc = previousStaticVars;
            this._currentNewOwned = previousNewOwned;
        }

        if (proc.isFunction || (opts.returnOnProperty && proc.isProperty)) {
            return localEnv.get(proc.name.name);
        }
        return opts.subReturnValue;
    }

    public setSourceModule(moduleName: string) {
        if (moduleName === '' && !this.currentSourceModule) {
            throw new Error("Cannot set empty module name when no module is currently set. Provide an explicit module name (e.g., 'Module1') or use Attribute VB_Name in VBA source.");
        }
        const effectiveModuleName = moduleName || this.currentSourceModule;
        if (effectiveModuleName && effectiveModuleName.length > 31) {
            throw new Error(`Module name '${effectiveModuleName}' exceeds the maximum length of 31 characters (MS-VBAL §5.2). Current length: ${effectiveModuleName.length}`);
        }
        this.currentSourceModule = effectiveModuleName;
        if (effectiveModuleName) {
            // モジュール名を VbaNamespaceRef として登録。bare 使用（VarType(Mod1) 等）を検出するためのセンチネル。
            // すでに変数として宣言されている名前は上書きしない。
            if (!this.env.hasVariable(effectiveModuleName.toLowerCase())) {
                this.env.set(effectiveModuleName.toLowerCase(), new VbaNamespaceRef(effectiveModuleName, 'module'));
            }
        }
    }

    /**
     * モジュールレベルでシンボルテーブルに登録すべき宣言文かどうかを返す。
     * false の場合はモジュールレベル実行文として pendingTopLevel に退避する。
     */
    private isModuleLevelDeclaration(stmt: Statement): boolean {
        switch (stmt.type) {
            case 'ProcedureDeclaration':
            case 'VariableDeclaration':
            case 'ConstDeclaration':
            case 'TypeDeclaration':
            case 'ClassDeclaration':
            case 'EnumDeclaration':
            case 'OptionExplicitStatement':
            case 'OptionBaseStatement':
            case 'OptionCompareStatement':
            case 'OptionPrivateModuleStatement':
            case 'DefDirective':
            case 'AttributeStatement':
            case 'DeclareStatement':
            case 'EventDeclaration':
            case 'ImplementsDirective':
            case 'LabelStatement':
                return true;
            default:
                return false;
        }
    }

    /**
     * Pass 1: 1モジュール分の AST を登録する（シンボルテーブル構築）。
     *
     * - ConstDeclaration      → スキップ。Pass 2（resolveIdentifiers）で依存順に評価。
     * - VariableDeclaration   → 配列境界なしは即時登録。配列境界ありは pendingArrayDecls
     *                           に退避し Pass 2 で Const 解決後に評価。
     * - モジュールレベル実行文 → resolveIdentifiers 完了前（バッチロード中）は
     *                           pendingTopLevel に退避。Pass 2 後に実行。
     *                           resolveIdentifiers 完了後（インタラクティブ呼び出し）は
     *                           即時実行（evalExpression や追加モジュールロード等）。
     */
    public evaluateModule(program: Program) {
        if (this.currentSourceModule && !this.env.hasVariable(this.currentSourceModule.toLowerCase())) {
            this.env.set(this.currentSourceModule.toLowerCase(), new VbaNamespaceRef(this.currentSourceModule, 'module'));
        }

        const batchMode = !this.resolveIdentifiersDone;
        const topLevelStmts: Statement[] = [];

        for (const stmt of program.body) {
            // ConstDeclaration は常に Pass 2 に委ねる。
            // バッチロード後のインタラクティブ呼び出しでは即時評価する。
            if (stmt.type === 'ConstDeclaration') {
                if (!batchMode) this.evaluateConstDeclaration(stmt as ConstDeclaration);
                continue;
            }

            // VariableDeclaration: 配列境界あり → Pass 2 まで評価を延期してシンボル登録のみ。
            // 境界なし → 即時登録（初期値は型デフォルトのみで Const 参照なし）。
            if (batchMode && stmt.type === 'VariableDeclaration') {
                const varDecl = stmt as VariableDeclaration;
                if (varDecl.declarations.some(d => d.isArray && d.arrayBounds && d.arrayBounds.length > 0)) {
                    this.pendingArrayDecls.push({ stmt: varDecl, moduleName: this.currentSourceModule || '' });
                    continue;
                }
            }

            // 宣言文でないモジュールレベル実行文はバッチ中は退避し、Pass 2 後にまとめて実行。
            if (batchMode && !this.isModuleLevelDeclaration(stmt)) {
                topLevelStmts.push(stmt);
                continue;
            }

            this.evaluateStatement(stmt);
        }

        if (batchMode && topLevelStmts.length > 0) {
            this.pendingTopLevel.push({ moduleName: this.currentSourceModule || '', stmts: topLevelStmts });
        }
    }

    /**
     * `+`/`-` の連鎖を再帰的に左へ辿り、最終的な葉が呼び出し可能な形
     * （`Identifier`/`CallExpression`/`MemberExpression`）かどうかを判定する。
     * `evalExpression()` の高速パスで、`x + 1` のような裸の式が VBA の
     * 暗黙 Call 文と曖昧な形になっていないかを検出するために使う。
     * 括弧で明示的に囲まれた場合（`(x) + 1`）は曖昧性が解消されるため false。
     */
    private isCallableLeftmostLeaf(expr: Expression): boolean {
        let node: Expression = expr;
        while (true) {
            switch (node.type) {
                case 'Identifier':
                case 'MemberExpression':
                    return true;
                case 'CallExpression':
                    // foo(args) + 1 は括弧で引数リストが確定しており、
                    // 暗黙 Call 文との曖昧性はない。
                    return false;
                case 'BinaryExpression': {
                    const op = (node as BinaryExpression).operator;
                    if (op !== '+' && op !== '-') return false;
                    node = (node as BinaryExpression).left;
                    continue;
                }
                default:
                    return false;
            }
        }
    }

    /** @deprecated Use {@link evaluateModule} instead. */
    public evaluate(program: Program) {
        this.evaluateModule(program);
        this.resolveIdentifiers([{ ast: program, moduleName: this.currentSourceModule }]);
    }

    public evalExpression(exprString: string): any {
        // eval() はどのファイルにも属さない専用のトップレベル（独立した呼び出しフレーム）として
        // 評価する。callProcedure が手続き呼び出しごとに退避・復元している状態と同様に、
        // currentSourceModule（module-level Dim の書き込み先解決に使う）と On Error 関連の状態
        // （errorHandlerLabel 等）を退避することで、
        //   - ロード済みファイルの currentSourceModule が残っていることによる変数スコープのズレ
        //   - 直前の eval() 呼び出しで設定した On Error 状態が後続の eval() に漏れ残る問題
        // の両方を防ぐ。
        const savedSourceModule = this.currentSourceModule;
        const savedErrorHandlerLabel = this.errorHandlerLabel;
        const savedErrorHandlingMode = this.errorHandlingMode;
        const savedIsInErrorHandler = this.isInErrorHandler;
        const savedLastErrorIndex = this.lastErrorIndex;
        this.currentSourceModule = '';
        this.errorHandlerLabel = null;
        this.errorHandlingMode = 'None';
        this.isInErrorHandler = false;
        this.lastErrorIndex = null;
        try {
            const lexer = new Lexer(exprString);
            const tokens = lexer.tokenize();

            // 単一式としてパースできるか試す（パース段階のみを try/catch で囲む）。
            // 実行（callProcedure/evaluateExpression）はこの外側で行うことで、
            // 実行中に発生した本物のランタイムエラーが catch に飲み込まれて
            // 「Parse error: syntax error」のような無関係なフォールバックエラーに
            // 化けてしまう問題を防ぐ（例: 式の中で呼んだ関数が Error 91 を投げても
            // 黒く握りつぶされ、文として再解析した際の無関係な構文エラーだけが見える）。
            let fastPathExpr: Expression | null = null;
            try {
                const exprParser = new Parser(tokens);
                const expr = exprParser.parseExpressionPublic();
                const nextToken = (exprParser as any).peek();

                // 単一式として全体を消費したか判定する。直後が Newline の場合は
                // 「末尾の改行のみ（その後は EOF）」であることまで確認する。
                // そうしないと複数行コードの1行目だけを式として誤評価し、残りの行を
                // 黒く無視してしまう（例: "x = 10\nDebug.Print 1" の "x = 10" を
                // 等価比較式として評価して返し、Debug.Print 以降を実行しない）。
                let fullyConsumed = !nextToken || nextToken.type === TokenType.EOF;
                if (!fullyConsumed && nextToken.type === TokenType.Newline) {
                    let aheadOffset = 0;
                    while ((exprParser as any).peek(aheadOffset).type === TokenType.Newline) aheadOffset++;
                    fullyConsumed = (exprParser as any).peek(aheadOffset).type === TokenType.EOF;
                }

                // VBA の statement 文法では、識別子（や `obj.Member`・`arr(i)` などの呼び出し
                // 可能な形）で始まり `=`/`+`/`-` が続く裸の文は、単独行でも複数文中でも常に
                // 「代入文」または「暗黙の Call 文」として解釈される（`x = 10` は代入、
                // `x + 1` は `x` を引数 `+1` で呼ぶ Call 文。`<`/`&`/`And` 等はこの曖昧性を
                // 持たず、statement としては Parse error になるため対象外）。
                // ここで式として確定させてしまうと、代入が一切実行されなかったり
                // （`x = 10` を比較式として評価）、呼び出し可能な手続きが暗黙的に
                // Empty(0) 扱いされて誤った算術結果を返したりする
                // （`Helper + 1` で `Helper` が必須引数を持つ場合など）。
                // 括弧で明示的に囲った場合（`(x) + 1`）は曖昧性が解消されるため対象外。
                const isStatementAmbiguous = expr.type === 'BinaryExpression' && (
                    (expr as BinaryExpression).operator === '=' ||
                    (((expr as BinaryExpression).operator === '+' || (expr as BinaryExpression).operator === '-')
                        && this.isCallableLeftmostLeaf((expr as BinaryExpression).left))
                );

                if (fullyConsumed && !isStatementAmbiguous) {
                    fastPathExpr = expr;
                }
            } catch {
                // パース失敗。フォールバック（文として解析）に委ねる。
            }

            if (fastPathExpr) {
                // If it is just an Identifier matching a Procedure, run it as a CallStatement
                if (fastPathExpr.type === 'Identifier') {
                    const name = (fastPathExpr as Identifier).name;
                    if (this.env.getProcedure(name)) {
                        this.callProcedure(name, []);
                        return undefined;
                    }
                }
                // Otherwise evaluate as a typical expression returning a value
                return this.evaluateExpression(fastPathExpr);
            }

            // Fallback: parse and evaluate as a statement sequence.
            // executeStatements()（手続き本体の実行に使う、On Error Resume Next/GoTo・GoTo・
            // Resume に対応した実行ループ）を経由することで、eval() 内の複数文でも実際の
            // Sub/Function 呼び出しと同じ On Error セマンティクスが働くようにする。
            // （evalExpression は VBARunner.eval() からしか呼ばれず、常に resolveIdentifiers
            //  完了後に実行されるため、Pass 1 バッチロード用の分岐は経由しない）
            const stmtParser = new Parser(tokens);
            const program = stmtParser.parse();
            // resolveIdentifiers（Pass 2）は最初の _ensureResolved() 時点のモジュール群しか
            // 解析しないため、eval() でその後に定義したプロシージャは Option Explicit の
            // 静的チェック対象に一度も載らない。ここで都度チェックし、違反があれば
            // optionExplicitViolations に追加登録する（既存の登録は維持）。
            const { violatedProcedures } = checkOptionExplicit(program);
            for (const [procName, undeclared] of violatedProcedures) {
                if (undeclared.size > 0) {
                    this.optionExplicitViolations.set(procName, undeclared);
                }
            }
            this.executeStatements(program.body, 0);
            return undefined;
        } finally {
            this.currentSourceModule = savedSourceModule;
            this.errorHandlerLabel = savedErrorHandlerLabel;
            this.errorHandlingMode = savedErrorHandlingMode;
            this.isInErrorHandler = savedIsInErrorHandler;
            this.lastErrorIndex = savedLastErrorIndex;
        }
    }

    private evaluateStatement(stmt: Statement) {
        if (stmt.line !== undefined) {
            this.currentLine = stmt.line;
            if (this.debugHook && this.vbaCallStack.length > 0) {
                this.debugHook.onBeforeStatement(stmt.line, this.vbaCallStack.length, this.env, [...this.vbaCallStack]);
            }
        }
        switch (stmt.type) {
            case 'ForStatement':
                this.evaluateForStatement(stmt as ForStatement);
                break;
            case 'ForEachStatement':
                this.evaluateForEachStatement(stmt as ForEachStatement);
                break;
            case 'IfStatement':
                this.evaluateIfStatement(stmt as IfStatement);
                break;
            case 'DoWhileStatement':
                this.evaluateDoWhileStatement(stmt as DoWhileStatement);
                break;
            case 'WhileStatement':
                this.evaluateWhileStatement(stmt as WhileStatement);
                break;
            case 'AssignmentStatement':
                this.evaluateAssignmentStatement(stmt as AssignmentStatement);
                break;
            case 'ProcedureDeclaration': {
                const procDecl = stmt as ProcedureDeclaration;
                // Use Module1 as default module name for global scope
                const moduleName = this.currentSourceModule || 'Module1';
                procDecl.moduleName = moduleName;
                const procName = procDecl.name.name;
                this.env.setProcedureWithModule(procName, procDecl, moduleName);
                break;
            }
            case 'VariableDeclaration':
                this.evaluateVariableDeclaration(stmt as VariableDeclaration);
                break;
            case 'CallStatement':
                this.evaluateCallStatement(stmt as CallStatement);
                break;
            case 'ConstDeclaration':
                this.evaluateConstDeclaration(stmt as ConstDeclaration);
                break;
            case 'SetStatement':
                this.evaluateSetStatement(stmt as SetStatement);
                break;
            case 'OnErrorStatement':
                this.evaluateOnErrorStatement(stmt as OnErrorStatement);
                break;
            case 'ResumeStatement':
                this.evaluateResumeStatement(stmt as ResumeStatement);
                break;
            case 'EraseStatement':
                this.evaluateEraseStatement(stmt as EraseStatement);
                break;
            case 'ReDimStatement':
                this.evaluateReDimStatement(stmt as ReDimStatement);
                break;
            case 'ExitStatement':
                this.evaluateExitStatement(stmt as ExitStatement);
                break;
            case 'EnumDeclaration':
                this.evaluateEnumDeclaration(stmt as EnumDeclaration);
                break;
            case 'TypeDeclaration':
                this.evaluateTypeDeclaration(stmt as TypeDeclaration);
                break;
            case 'ClassDeclaration':
                this.evaluateClassDeclaration(stmt as ClassDeclaration);
                break;
            case 'OptionCompareStatement':
                this.evaluateOptionCompareStatement(stmt as OptionCompareStatement);
                break;
            case 'AttributeStatement':
                this.evaluateAttributeStatement(stmt as AttributeStatement);
                break;
            case 'DeclareStatement':
                this.evaluateDeclareStatement(stmt as DeclareStatement);
                break;
            case 'OpenStatement':
                this.evaluateOpenStatement(stmt as OpenStatement);
                break;
            case 'CloseStatement':
                this.evaluateCloseStatement(stmt as CloseStatement);
                break;
            case 'PrintStatement':
                this.evaluatePrintStatement(stmt as PrintStatement);
                break;
            case 'LineInputStatement':
                this.evaluateLineInputStatement(stmt as LineInputStatement);
                break;
            case 'PutStatement':
                this.evaluatePutStatement(stmt as PutStatement);
                break;
            case 'KillStatement':
                this.evaluateKillStatement(stmt as KillStatement);
                break;
            case 'WriteStatement':
                this.evaluateWriteStatement(stmt as WriteStatement);
                break;
            case 'InputStatement':
                this.evaluateInputStatement(stmt as InputStatement);
                break;
            case 'GetStatement':
                this.evaluateGetStatement(stmt as GetStatement);
                break;
            case 'SeekStatement':
                this.evaluateSeekStatement(stmt as SeekStatement);
                break;
            case 'ResetStatement':
                this.evaluateResetStatement(stmt as ResetStatement);
                break;
            case 'LabelStatement':
                // No-op for now. Label execution just passes through.
                break;
            case 'SelectCaseStatement':
                this.evaluateSelectCaseStatement(stmt as SelectCaseStatement);
                break;
            case 'WithStatement':
                this.evaluateWithStatement(stmt as WithStatement);
                break;
            case 'GoToStatement':
                this.evaluateGoToStatement(stmt as GoToStatement);
                break;
            case 'StopStatement':
                this.evaluateStopStatement(stmt as StopStatement);
                break;
            case 'EndStatement':
                this.evaluateEndStatement(stmt as EndStatement);
                break;
            case 'GoSubStatement':
                this.evaluateGoSubStatement(stmt as GoSubStatement);
                break;
            case 'ReturnStatement':
                this.evaluateReturnStatement(stmt as ReturnStatement);
                break;
            case 'OnGoToSubStatement':
                this.evaluateOnGoToSubStatement(stmt as OnGoToSubStatement);
                break;
            case 'OptionExplicitStatement':
                this.evaluateOptionExplicitStatement(stmt as OptionExplicitStatement);
                break;
            case 'OptionBaseStatement':
                this.evaluateOptionBaseStatement(stmt as OptionBaseStatement);
                break;
            case 'DefDirective':
                this.evaluateDefDirective(stmt as DefDirective);
                break;
            case 'OptionPrivateModuleStatement':
                this.evaluateOptionPrivateModuleStatement(stmt as OptionPrivateModuleStatement);
                break;
            case 'EventDeclaration':
                this.evaluateEventDeclaration(stmt as EventDeclaration);
                break;
            case 'RaiseEventStatement':
                this.evaluateRaiseEventStatement(stmt as RaiseEventStatement);
                break;
            case 'LSetStatement':
                this.evaluateLSetStatement(stmt as LSetStatement);
                break;
            case 'RSetStatement':
                this.evaluateRSetStatement(stmt as RSetStatement);
                break;
            case 'MidStatement':
                this.evaluateMidStatement(stmt as MidStatement);
                break;
            case 'ErrorStatement':
                this.evaluateErrorStatement(stmt as ErrorStatement);
                break;
            case 'ImplementsDirective':
                this.evaluateImplementsDirective(stmt as ImplementsDirective);
                break;
            case 'AppActivateStatement':
                this.evaluateAppActivateStatement(stmt as AppActivateStatement);
                break;
            case 'SendKeysStatement':
                this.evaluateSendKeysStatement(stmt as SendKeysStatement);
                break;
            case 'LockStatement':
                this.evaluateLockStatement(stmt as LockStatement);
                break;
            case 'UnlockStatement':
                this.evaluateUnlockStatement(stmt as UnlockStatement);
                break;
            case 'WidthStatement':
                this.evaluateWidthStatement(stmt as WidthStatement);
                break;
            default:
                throw new Error(`Execution error: Unknown statement type ${stmt.type}`);
        }
    }

    private evaluateForStatement(stmt: ForStatement) {
        const startValue = this.evaluateExpression(stmt.start);
        const endValue = this.evaluateExpression(stmt.end);
        const stepValue = stmt.step ? this.evaluateExpression(stmt.step) : 1;
        const varName = stmt.identifier.name;

        // Initialize block scope variable if it doesn't exist
        if (this.env.get(varName) === vbaEmpty) { // Check against vbaEmpty
            this.env.set(varName, startValue);
        } else {
            this.env.setLocally(varName, startValue);
        }

        const condition = () => stepValue > 0 ? this.env.get(varName) <= endValue : this.env.get(varName) >= endValue;

        while (condition()) {
            try {
                this.executeStatements(stmt.body, 0, false);
            } catch (e: any) {
                if (e && e.type === 'Exit' && e.target === 'For') break;
                if (e && e.type === 'GoTo') {
                    const idx = this.findLabelInBody(stmt.body, e.label);
                    if (idx >= 0) {
                        try {
                            this.executeStatements(stmt.body, idx + 1, false);
                        } catch (e2: any) {
                            if (e2 && e2.type === 'Exit' && e2.target === 'For') break;
                            throw e2;
                        }
                    } else {
                        throw e;
                    }
                } else {
                    throw e;
                }
            }
            // Increment/decrement loop variable
            this.env.setLocally(varName, this.env.get(varName) + stepValue);
        }
    }

    private evaluateForEachStatement(stmt: ForEachStatement) {
        const collection = this.evaluateExpression(stmt.collection);
        const varName = stmt.variable.name;

        let elements: any[];
        if (Array.isArray(collection)) {
            const dimensions = (collection as any).__vbaDimensions__;
            if (dimensions && dimensions.length > 1) {
                elements = this.enumerateMultiDimArray(collection, dimensions);
            } else {
                elements = this.flattenArray(collection);
            }
        } else if (collection && collection.__isVbaDict__) {
            elements = Array.from((collection.__map__ as Map<any, any>).keys());
        } else if (collection && collection.__isVbaCollection__) {
            if (typeof collection[Symbol.iterator] === 'function') {
                elements = Array.from(collection as Iterable<any>);
            } else {
                elements = Array.isArray(collection.items) ? collection.items : [];
            }
        } else if (collection && typeof collection.items !== 'undefined') {
            elements = Array.isArray(collection.items) ? collection.items : [];
        } else {
            this.throwVbaError(VbaErrorCode.TYPE_MISMATCH, "Type mismatch: 'For Each' requires a collection or array");
        }

        for (const element of elements) {
            this.env.set(varName, element);
            try {
                this.executeStatements(stmt.body, 0, false);
            } catch (e: any) {
                if (e && e.type === 'Exit' && e.target === 'For') break;
                if (e && e.type === 'GoTo') {
                    const idx = this.findLabelInBody(stmt.body, e.label);
                    if (idx >= 0) {
                        try {
                            this.executeStatements(stmt.body, idx + 1, false);
                        } catch (e2: any) {
                            if (e2 && e2.type === 'Exit' && e2.target === 'For') break;
                            throw e2;
                        }
                    } else {
                        throw e;
                    }
                } else {
                    throw e;
                }
            }
        }
    }

    private enumerateMultiDimArray(arr: any, dimensions: { lower: number, upper: number }[]): any[] {
        const result: any[] = [];
        const indices = dimensions.map(d => d.lower);

        while (true) {
            let element: any = arr;
            for (const idx of indices) {
                element = element[idx];
            }
            if (element !== undefined) {
                result.push(element);
            }

            let dimIdx = 0;
            while (dimIdx < dimensions.length) {
                indices[dimIdx]++;
                if (indices[dimIdx] <= dimensions[dimIdx].upper) {
                    break;
                }
                indices[dimIdx] = dimensions[dimIdx].lower;
                dimIdx++;
            }

            if (dimIdx === dimensions.length) {
                break;
            }
        }

        return result;
    }

    private flattenArray(arr: any[]): any[] {
        const result: any[] = [];
        for (const item of arr) {
            if (Array.isArray(item)) {
                result.push(...this.flattenArray(item));
            } else {
                result.push(item);
            }
        }
        return result;
    }

    private toVbaNumber(val: any): number {
        try {
            return _vbaToNumber(val);
        } catch (e: any) {
            if (e?.type === 'VbaError') this.throwVbaError(e.number, e.message);
            throw e;
        }
    }

    private vbaRound(val: number, decimals: number = 0): number { return _vbaRound(val, decimals); }

    private throwVbaError(number: number, message: string, overrideLine?: number, overrideModule?: string): never {
        const line = overrideLine ?? (this.currentLine || undefined);
        const mod = overrideModule ?? (this.executingModuleName || this.currentSourceModule || null);
        const msg = line !== undefined ? `Run-time error '${number}': ${message} (line ${line})` : `Run-time error '${number}': ${message}`;
        const err: any = new Error(msg);
        err.type = 'VbaError';
        err.number = number;
        err.vbaLine = line;
        err.vbaModule = mod;
        err.vbaStack = [...this.vbaCallStack].reverse();
        // Err.Description は "Run-time error 'N': ..." の枠組みテキストを含まず、
        // 生のメッセージ（例: "Type mismatch"）だけを保持する（実 VBA 仕様）。
        // TypeScript 側に投げる例外の message には引き続き枠組み付きの文字列を使う
        // （eval()/run() 呼び出し元向けの分かりやすさのため、既存の挙動を維持）。
        err.vbaBareMessage = message;
        throw err;
    }

    /** コンパイルエラーをプレフィックスなしで throw する。呼び出し元の catch ブロックで
     *  "Compile error:" を一元付与する（precheckProc / resolveIdentifiers / evaluateVariableDeclaration）。 */
    private throwCompileError(number: number, message: string, line?: number, module?: string): never {
        const err: any = new Error(message);
        err._precheckRaw = true;
        err.number = number;
        err.vbaLine = line;
        err.vbaModule = module;
        throw err;
    }

    private evaluateIfStatement(stmt: IfStatement) {
        const conditionVal = this.evaluateExpression(stmt.condition);
        if (this.isTrue(conditionVal)) {
            this.executeStatements(stmt.consequent, 0, false);
        } else if (stmt.alternate) {
            if (Array.isArray(stmt.alternate)) {
                this.executeStatements(stmt.alternate as Statement[], 0, false);
            } else {
                this.evaluateIfStatement(stmt.alternate as IfStatement);
            }
        }
    }

    private evaluateSelectCaseStatement(stmt: SelectCaseStatement) {
        const selectVal = this.evaluateExpression(stmt.expression);

        for (const caseClause of stmt.cases) {
            let matched = false;
            for (const range of caseClause.ranges) {
                if (range.kind === 'expression') {
                    matched = selectVal === this.evaluateExpression(range.value);
                } else if (range.kind === 'to') {
                    const start = this.evaluateExpression(range.start);
                    const end = this.evaluateExpression(range.end);
                    matched = selectVal >= start && selectVal <= end;
                } else {
                    // comparison: selectVal <op> value
                    const val = this.evaluateExpression(range.value);
                    switch (range.operator) {
                        case '=':  matched = selectVal === val; break;
                        case '<>': matched = selectVal !== val; break;
                        case '<':  matched = selectVal < val;   break;
                        case '>':  matched = selectVal > val;   break;
                        case '<=': matched = selectVal <= val;  break;
                        case '>=': matched = selectVal >= val;  break;
                    }
                }
                if (matched) break;
            }
            if (matched) {
                this.executeStatements(caseClause.body, 0, false);
                return;
            }
        }

        if (stmt.elseBody) {
            this.executeStatements(stmt.elseBody, 0, false);
        }
    }

    private evaluateDoWhileStatement(stmt: DoWhileStatement) {
        const checkCondition = (): boolean => {
            if (stmt.condition === undefined) return true; // infinite
            const val = this.evaluateExpression(stmt.condition);
            const truthy = this.isTrue(val);
            return stmt.conditionType === 'until' ? !truthy : truthy;
        };

        // eslint-disable-next-line no-constant-condition
        while (true) {
            // pre-condition check
            if (stmt.conditionPosition === 'pre' && !checkCondition()) break;

            try {
                this.executeStatements(stmt.body, 0, false);
            } catch (e: any) {
                if (e && e.type === 'Exit' && e.target === 'Do') break;
                if (e && e.type === 'GoTo') {
                    const idx = this.findLabelInBody(stmt.body, e.label);
                    if (idx >= 0) {
                        try {
                            this.executeStatements(stmt.body, idx + 1, false);
                        } catch (e2: any) {
                            if (e2 && e2.type === 'Exit' && e2.target === 'Do') break;
                            throw e2;
                        }
                    } else {
                        throw e;
                    }
                } else {
                    throw e;
                }
            }

            // post-condition check
            if (stmt.conditionPosition === 'post' && !checkCondition()) break;

            // infinite loop: no condition, no break unless Exit Do
            if (stmt.conditionPosition === undefined) continue;
        }
    }

    private evaluateWhileStatement(stmt: WhileStatement) {
        while (this.isTrue(this.evaluateExpression(stmt.condition))) {
            try {
                this.executeStatements(stmt.body, 0, false);
            } catch (e: any) {
                if (e && e.type === 'GoTo') {
                    const idx = this.findLabelInBody(stmt.body, e.label);
                    if (idx >= 0) {
                        this.executeStatements(stmt.body, idx + 1, false);
                    } else {
                        throw e;
                    }
                } else {
                    throw e;
                }
            }
        }
    }

    private evaluateWithStatement(stmt: WithStatement) {
        const obj = this.evaluateExpression(stmt.expression);
        this.withObjectStack.push(obj);
        try {
            this.executeStatements(stmt.body, 0, false);
        } finally {
            this.withObjectStack.pop();
        }
    }

    private findLabelInBody(body: Statement[], label: string): number {
        const lower = label.toLowerCase();
        return body.findIndex(s => s.type === 'LabelStatement' && (s as any).label.toLowerCase() === lower);
    }

    private evaluateGoToStatement(stmt: GoToStatement) {
        throw { type: 'GoTo', label: stmt.label };
    }

    private evaluateStopStatement(_stmt: StopStatement) {
        console.log('STOP Statement encountered');
        // implementation-defined: just log for now
    }

    private evaluateEndStatement(_stmt: EndStatement) {
        throw { type: 'Terminate' };
    }

    private evaluateGoSubStatement(stmt: GoSubStatement) {
        throw { type: 'GoSub', label: stmt.label };
    }

    private evaluateReturnStatement(_stmt: ReturnStatement) {
        throw { type: 'Return' };
    }

    private evaluateOnGoToSubStatement(stmt: OnGoToSubStatement) {
        const val = this.evaluateExpression(stmt.expression);
        const idx = Math.floor(Number(val));

        if (idx < 0 || idx > 255) {
            this.throwVbaError(VbaErrorCode.INVALID_PROCEDURE_CALL, `Invalid procedure call or argument (On...GoTo/GoSub index ${idx})`);
        }

        if (idx >= 1 && idx <= stmt.labels.length) {
            const label = stmt.labels[idx - 1];
            if (stmt.isGoSub) {
                throw { type: 'GoSub', label };
            } else {
                throw { type: 'GoTo', label };
            }
        }
    }

    private evaluateLSetStatement(stmt: LSetStatement) {
        const val = String(this.evaluateExpression(stmt.right) || '');
        if (stmt.left.type === 'Identifier') {
            const name = (stmt.left as Identifier).name;
            const target = String(this.env.get(name) || '');
            const result = val.padEnd(target.length, ' ').substring(0, target.length);
            this.env.set(name, result);
        } else {
            this.throwVbaError(VbaErrorCode.TYPE_MISMATCH, 'Type mismatch');
        }
    }

    private evaluateRSetStatement(stmt: RSetStatement) {
        const val = String(this.evaluateExpression(stmt.right) || '');
        if (stmt.left.type === 'Identifier') {
            const name = (stmt.left as Identifier).name;
            const target = String(this.env.get(name) || '');
            const result = val.padStart(target.length, ' ').substring(0, target.length);
            this.env.set(name, result);
        } else {
            this.throwVbaError(VbaErrorCode.TYPE_MISMATCH, 'Type mismatch');
        }
    }

    private evaluateMidStatement(stmt: MidStatement) {
        const target = String(this.evaluateExpression(stmt.target) ?? '');
        let start = Number(this.evaluateExpression(stmt.start));
        const value = String(this.evaluateExpression(stmt.value) ?? '');
        let rawLength = stmt.length !== null ? Number(this.evaluateExpression(stmt.length)) : null;
        // MidB uses UTF-16 byte offsets (2 bytes per char): convert to character positions
        if (stmt.isByte) {
            start = Math.ceil(start / 2);
            if (rawLength !== null) rawLength = Math.ceil(rawLength / 2);
        }
        if (start < 1 || start > target.length) {
            throwVbaError(VbaErrorCode.INVALID_PROCEDURE_CALL);
        }
        const effectiveStart = start - 1;
        const replaceLen = rawLength !== null
            ? Math.min(rawLength, target.length - effectiveStart, value.length)
            : Math.min(value.length, target.length - effectiveStart);
        const newStr = target.substring(0, effectiveStart)
            + value.substring(0, replaceLen)
            + target.substring(effectiveStart + replaceLen);
        this.evaluateAssignmentToVariable(stmt.target, newStr);
    }

    private evaluateErrorStatement(stmt: ErrorStatement) {
        const errNum = this.evaluateExpression(stmt.errorNumber);
        const errObj = this.env.get('err');
        if (errObj) {
            errObj.raise(errNum);
        } else {
            this.throwVbaError(Number(errNum), String(errNum));
        }
    }

    private evaluateEventDeclaration(_stmt: EventDeclaration) {
        // Top-level event declarations are just metadata for the class
    }

    private evaluateRaiseEventStatement(stmt: RaiseEventStatement) {
        const eventName = stmt.eventName.name.toLowerCase();
        // Look for the instance in the current environment (Me)
        const me = this.env.get('Me');
        if (me && me.__events__) {
            const handlers = me.__events__.get(eventName);
            if (handlers) {
                const args = stmt.args.map((a: any) => this.evaluateExpression(a));
                for (const handler of handlers) {
                    handler(...args);
                }
            }
        }
    }

    private evaluateImplementsDirective(_stmt: ImplementsDirective) {
        // No-op for now. Used for interface compliance metadata.
    }

    private evaluateAppActivateStatement(stmt: AppActivateStatement) {
        const title = String(this.evaluateExpression(stmt.title));
        console.log(`[STUB] AppActivate "${title}"`);
    }

    private evaluateSendKeysStatement(stmt: SendKeysStatement) {
        const keys = String(this.evaluateExpression(stmt.keys));
        console.log(`[STUB] SendKeys "${keys}"`);
    }

    private evaluateLockStatement(stmt: LockStatement) {
        const fileNum = Number(this.evaluateExpression(stmt.fileNumber));
        console.log(`[STUB] Lock #${fileNum}`);
    }

    private evaluateUnlockStatement(stmt: UnlockStatement) {
        const fileNum = Number(this.evaluateExpression(stmt.fileNumber));
        console.log(`[STUB] Unlock #${fileNum}`);
    }

    private evaluateWidthStatement(stmt: WidthStatement) {
        const fileNum = Number(this.evaluateExpression(stmt.fileNumber));
        const width = Number(this.evaluateExpression(stmt.width));
        console.log(`[STUB] Width #${fileNum}, ${width}`);
    }

    private evaluateAssignmentStatement(stmt: AssignmentStatement) {

        let val = this.evaluateExpression(stmt.right);

        // FEATURE: Implicit default Value getter (result = obj -> result = obj.Value)
        // If RHS evaluates to an object with a default Value property, call the getter
        if (val && typeof val === 'object' && val.__vbaClass__ && !(val instanceof VbaDate) && !(val instanceof VbaBoolean)) {
            const classDef = val.__classDef__ as ClassDeclaration;
            if (classDef) {
                const valueGetter = classDef.procedures.find(
                    p => p.isProperty && p.propertyType === 'get' && p.name.name.toLowerCase() === 'value'
                );
                if (valueGetter) {
                    // Call the Value property getter to extract the value
                    val = this.callClassMethod(val, valueGetter, []);
                }
            }
        }

        // FEATURE: Implicit default Value getter for non-VBA-class mock objects
        // __vbaDefault__ = true のオブジェクトはデフォルトプロパティを持つとみなし、
        // Value getter を使って実値を取り出す。setter パスも同じ Value を使うので対称になる。
        // __vbaDefault__ ガードがあるため VbaDate 等の内部型は影響を受けない。
        if (val && val.__vbaDefault__ === true) {
            const valueKey = this.resolveObjectMemberKey(val, 'value');
            if (valueKey !== undefined) val = (val as any)[valueKey];
        }

        this.evaluateAssignmentToVariable(stmt.left, val);
    }

    /**
     * VBA Let-coercion: 宣言型に合わせて値を強制変換する。
     * 現在は Boolean のみサポート。他の型（Integer/Long/...）の coercion は後続作業。
     */
    private coerceToDeclaredType(val: any, vbaType: string): any {
        // Null は数値・Boolean 型では「Invalid use of Null (Error 94)」
        if (val === vbaNull) {
            const errorTypes = new Set(['Boolean', 'Byte', 'Integer', 'Long', 'LongLong', 'LongPtr', 'Single', 'Double', 'Currency', 'Date']);
            if (errorTypes.has(vbaType)) {
                this.throwVbaError(VbaErrorCode.INVALID_USE_OF_NULL, 'Invalid use of Null');
            }
            return val;
        }
        switch (vbaType) {
            case 'Boolean':
                return this.coerceToBoolean(val);
            case 'Byte':
                return this.env.get('cbyte')(val);
            case 'Integer':
                return this.env.get('cint')(val);
            case 'Long':
                return this.env.get('clng')(val);
            case 'LongLong':
            case 'LongPtr':
                return this.env.get('clnglng')(val);
            case 'Single':
                return this.env.get('csng')(val);
            case 'Double':
                return this.env.get('cdbl')(val);
            case 'Currency':
                return this.env.get('ccur')(val);
            case 'String':
                return this.env.get('cstr')(val);
            case 'Date':
                return this.env.get('cdate')(val);
            default:
                return val;
        }
    }

    private coerceToBoolean(val: any): VbaBoolean { return vbaToBoolean(val); }

    private evaluateAssignmentToVariable(left: Expression, val: any) {
        if (left.type === 'Identifier') {
            const name = (left as Identifier).name;
            let variable = this.env.get(name);

            // Resolve auto-instance placeholder if needed
            variable = this.resolveAutoInstance(left as Identifier, variable);

            // Check if this is a class instance with a default Value property
            // FEATURE: Implicit default Value assignment (obj = value -> obj.Value = value)
            if (variable && typeof variable === 'object' && variable.__vbaClass__) {
                const classDef = variable.__classDef__ as ClassDeclaration;
                if (classDef) {
                    const valueSetter = classDef.procedures.find(
                        p => p.isProperty && (p.propertyType === 'let' || p.propertyType === 'set') && p.name.name.toLowerCase() === 'value'
                    );
                    if (valueSetter) {
                        // Call the Value property setter with the assigned value
                        this.callClassMethod(variable, valueSetter, [val]);
                        return;
                    }
                }
            }

            const oldVal = this.env.get(name);
            // B-3: Only on explicit Nothing assignment (see evaluateSetStatement for rationale)
            if (val === vbaNothing && oldVal !== val) this.triggerTerminate(oldVal);

            const procNameLower = name.toLowerCase();

            // Special case: assigning to current function/property name (return value)
            if (this.currentProcedureName && this.currentProcedureName.toLowerCase() === procNameLower) {
                if (this.currentProcedureType === 'function' || this.currentProcedureType === 'get') {
                    // 関数の戻り値も宣言された返り値型に Let-coercion される
                    const retType = (this.currentProcedureType === 'function' || this.currentProcedureType === 'get')
                        ? this.env.getVariableType(name)
                        : null;
                    const coerced = retType ? this.coerceToDeclaredType(val, retType.vbaType) : val;
                    this.env.setLocally(name, coerced);
                    return;
                }
            }

            const proc = this.env.getProcedure(name, 'let');
            if (proc) {
                this.callProcedure(name, [val], 'let');
                return;
            }

            // Let-coercion: 宣言型に合わせて値を変換
            const typeInfo = this.env.getVariableType(name);
            if (typeInfo) {
                val = this.coerceToDeclaredType(val, typeInfo.vbaType);
            }
            this.env.set(name, val);
        } else if (left.type === 'CallExpression') {
            // Array/Dictionary assignment: arr(0) = val OR dict("key") = val
            const call = left as CallExpression;
            if (call.callee.type === 'Identifier') {
                const name = (call.callee as Identifier).name;
                const lowerName = name.toLowerCase();

                if (lowerName === 'mid' || lowerName === 'mid$' || lowerName === 'midb' || lowerName === 'midb$') {
                    // Mid(s, start, [len]) = val
                    const targetName = (call.args[0] as Identifier).name; // Must be identifier for assignment
                    let start = this.evaluateExpression(call.args[1]) as number;
                    let length = call.args.length > 2 ? this.evaluateExpression(call.args[2]) as number : -1;

                    const isByte = lowerName.startsWith('midb');
                    if (isByte) {
                        // Convert byte position to char position (approximate: 2 bytes per char)
                        start = Math.floor((start + 1) / 2);
                        if (length !== -1) length = Math.floor(length / 2);
                    }

                    const sourceStr = String(this.env.get(targetName) || '');
                    const replacement = String(val || '');

                    let replaceLen = length === -1 ? replacement.length : length;
                    // Mid statement rules: replacement is limited by length of source and specified length
                    replaceLen = Math.min(replaceLen, replacement.length, sourceStr.length - start + 1);

                    const head = sourceStr.substring(0, start - 1);
                    const tail = sourceStr.substring(start - 1 + replaceLen);
                    const mid = replacement.substring(0, replaceLen);

                    this.env.set(targetName, head + mid + tail);
                    return;
                }

                const target = this.env.get(name);

                if (Array.isArray(target)) {
                    const dims = (target as any).__vbaDimensions__ as { lower: number, upper: number }[] | undefined;
                    if (dims && call.args.length !== dims.length) this.throwVbaError(VbaErrorCode.SUBSCRIPT_OUT_OF_RANGE, 'Subscript out of range');
                    // VBA index == JS index. Multi-dimensional: arr(i, j) = val -> arr[i][j] = val
                    let current = target;
                    for (let i = 0; i < call.args.length - 1; i++) {
                        const d = this.evaluateExpression(call.args[i]) as number;
                        if (dims) {
                            const { lower, upper } = dims[i];
                            if (d < lower || d > upper) this.throwVbaError(VbaErrorCode.SUBSCRIPT_OUT_OF_RANGE, 'Subscript out of range');
                        }
                        if (!current[d]) {
                            current[d] = [];
                        }
                        current = current[d];
                    }
                    const lastIdx = this.evaluateExpression(call.args[call.args.length - 1]) as number;
                    if (dims) {
                        const { lower, upper } = dims[call.args.length - 1];
                        if (lastIdx < lower || lastIdx > upper) this.throwVbaError(VbaErrorCode.SUBSCRIPT_OUT_OF_RANGE, 'Subscript out of range');
                    }
                    current[lastIdx] = val;
                } else if (target && target.__isVbaDict__) {
                    // Treat as Dictionary assignment dict("key") = val
                    const key = String(this.evaluateExpression(call.args[0]));
                    target.__map__.set(key, val);
                } else if (target && target.__vbaClass__) {
                    // Default property assignment: obj(args) = val -> obj.Item(args) = val
                    const classDef = target.__classDef__ as ClassDeclaration;
                    const setter = classDef.procedures.find(
                        p => p.isProperty && (p.propertyType === 'let' || p.propertyType === 'set') && p.name.name.toLowerCase() === 'item'
                    );
                    if (setter) {
                        const argsVals = call.args.map(a => this.evaluateExpression(a));
                        this.callClassMethod(target, setter, [...argsVals, val]);
                    } else {
                        this.throwVbaError(VbaErrorCode.OBJECT_DOESNT_SUPPORT_PROPERTY, "Object doesn't support this property or method");
                    }
                } else {
                    this.throwVbaError(VbaErrorCode.SUBSCRIPT_OUT_OF_RANGE, 'Subscript out of range');
                }
            } else if (call.callee.type === 'MemberExpression') {
                // obj.Method(key) = val  →  Property Let / Dictionary item set
                const memberCallee = call.callee as MemberExpression;
                const obj = this.resolveAutoInstance(memberCallee.object, this.evaluateExpression(memberCallee.object));
                const methodName = memberCallee.property.name.toLowerCase();
                if (obj && obj.__isVbaDict__) {
                    // dict.Item(key) = val
                    const key = String(this.evaluateExpression(call.args[0]));
                    obj.__map__.set(key, val);
                } else if (obj && obj.__vbaClass__) {
                    const classDef = obj.__classDef__ as ClassDeclaration;
                    const setter = classDef.procedures.find(
                        p => p.isProperty && (p.propertyType === 'let' || p.propertyType === 'set') && p.name.name.toLowerCase() === methodName
                    );
                    if (setter) {
                        const argsVals = call.args.map(a => this.evaluateExpression(a));
                        this.callClassMethod(obj, setter, [...argsVals, val]);
                    } else {
                        // Property Let/Set がなければ、配列フィールドへの外部インデックス代入を試みる
                        // 例: obj.Items(0) = val
                        const instanceEnvForArr = obj.__instanceEnv__ as Environment;
                        const fieldArr = instanceEnvForArr.get(methodName);
                        if (Array.isArray(fieldArr)) {
                            const dims = (fieldArr as any).__vbaDimensions__ as { lower: number, upper: number }[] | undefined;
                            if (dims && call.args.length !== dims.length) this.throwVbaError(VbaErrorCode.SUBSCRIPT_OUT_OF_RANGE, 'Subscript out of range');
                            let current = fieldArr;
                            for (let i = 0; i < call.args.length - 1; i++) {
                                const d = this.evaluateExpression(call.args[i]) as number;
                                if (dims) {
                                    const { lower, upper } = dims[i];
                                    if (d < lower || d > upper) this.throwVbaError(VbaErrorCode.SUBSCRIPT_OUT_OF_RANGE, 'Subscript out of range');
                                }
                                if (!current[d]) current[d] = [];
                                current = current[d];
                            }
                            const lastIdx = this.evaluateExpression(call.args[call.args.length - 1]) as number;
                            if (dims) {
                                const { lower, upper } = dims[call.args.length - 1];
                                if (lastIdx < lower || lastIdx > upper) this.throwVbaError(VbaErrorCode.SUBSCRIPT_OUT_OF_RANGE, 'Subscript out of range');
                            }
                            current[lastIdx] = val;
                        } else {
                            this.throwVbaError(VbaErrorCode.OBJECT_DOESNT_SUPPORT_PROPERTY, "Object doesn't support this property or method");
                        }
                    }
                } else if (obj && typeof obj === 'object') {
                    // obj.Method(args) = val: Method を呼び出し、結果が __vbaDefault__ を持つ場合は
                    // そのデフォルトプロパティ (Value) に代入する（例: ws.Range("A1:D1") = arr）
                    // メソッドは case-insensitive で解決する（VBA は大文字小文字不問）
                    const methodKey = this.resolveObjectMemberKey(obj, methodName);
                    const method = methodKey !== undefined ? (obj as any)[methodKey] : undefined;
                    if (typeof method === 'function') {
                        const argsVals = call.args.map(a => this.evaluateExpression(a));
                        const result = method.call(obj, ...argsVals);
                        if (result && result.__vbaDefault__ === true) {
                            const valueKey = this.resolveObjectMemberKey(result, 'value') ?? 'Value';
                            result[valueKey] = val;
                            return;
                        }
                    }
                    const key = String(this.evaluateExpression(call.args[0]));
                    obj[key] = val;
                } else {
                    this.throwVbaError(VbaErrorCode.INVALID_PROCEDURE_CALL, 'Invalid procedure call or argument');
                }
            } else if (call.callee.type === 'CallExpression') {
                // outer("sub")("x") = val  →  evaluate outer("sub") to get inner dict, then assign
                const innerObj = this.evaluateExpression(call.callee);
                if (innerObj && innerObj.__isVbaDict__) {
                    const key = String(this.evaluateExpression(call.args[0]));
                    innerObj.__map__.set(key, val);
                } else if (innerObj && typeof innerObj === 'object') {
                    const key = String(this.evaluateExpression(call.args[0]));
                    innerObj[key] = val;
                } else {
                    this.throwVbaError(VbaErrorCode.SUBSCRIPT_OUT_OF_RANGE, 'Subscript out of range');
                }
            } else {
                this.throwVbaError(VbaErrorCode.INVALID_PROCEDURE_CALL, 'Invalid procedure call or argument');
            }
        } else if (left.type === 'MemberExpression') {
            const member = left as MemberExpression;
            const obj = this.resolveAutoInstance(member.object, this.evaluateExpression(member.object));
            const propName = member.property.name.toLowerCase();
            if (obj && obj.__vbaClass__) {
                const classDef = obj.__classDef__ as ClassDeclaration;
                const instanceEnv = obj.__instanceEnv__ as Environment;
                const setter = classDef.procedures.find(
                    p => p.isProperty && (p.propertyType === 'let' || p.propertyType === 'set') && p.name.name.toLowerCase() === propName
                );
                if (setter) {
                    this.callClassMethod(obj, setter, [val]);
                } else {
                    instanceEnv.set(propName, val);
                }
            } else if (obj && typeof obj === 'object') {
                // VBA は大文字小文字不問: アクセサ(setter)・プロトタイプも辿って実キーを解決する
                const resolvedProp = this.resolveObjectMemberKey(obj, propName) ?? propName;
                obj[resolvedProp] = val;
            } else {
                if (obj === null || obj === undefined || obj === vbaNothing) {
                    this.throwVbaError(VbaErrorCode.OBJECT_VARIABLE_NOT_SET, 'Object variable or With block variable not set');
                } else {
                    this.throwVbaError(VbaErrorCode.OBJECT_REQUIRED, 'Object required');
                }
            }
        } else if (left.type === 'ImplicitWithObjectExpression') {
            if (this.withObjectStack.length === 0) {
                this.throwVbaError(VbaErrorCode.OBJECT_VARIABLE_NOT_SET, 'Object variable or With block variable not set');
            }
            const obj = this.withObjectStack[this.withObjectStack.length - 1];
            const member = left as ImplicitWithObjectExpression;
            const propName = member.property.name.toLowerCase();
            if (obj && typeof obj === 'object') {
                obj[propName] = val;
            } else {
                if (obj === null || obj === undefined || obj === vbaNothing) {
                    this.throwVbaError(VbaErrorCode.OBJECT_VARIABLE_NOT_SET, 'Object variable or With block variable not set');
                } else {
                    this.throwVbaError(VbaErrorCode.OBJECT_REQUIRED, 'Object required');
                }
            }
        } else {
            this.throwVbaError(VbaErrorCode.INVALID_PROCEDURE_CALL, 'Invalid procedure call or argument');
        }
    }

    private evaluateVariableDeclaration(stmt: VariableDeclaration) {
        const isStaticDecl = stmt.isStatic || this.currentProcIsStatic;
        for (const decl of stmt.declarations) {
            const varName = decl.name.name;
            if (decl.isWithEvents) {
                this.env.setWithEvents(varName);
            }
            const varKey = varName.toLowerCase();
            const staticKey = `${this.currentProcedureName?.toLowerCase()}:${varKey}`;

            // Register type metadata for typed declarations
            // If no explicit type, apply Def-Directive mapping by first letter (§5.2.2)
            const effectiveType = decl.objectType
                ?? this.defTypeMap.get(varName.charAt(0).toLowerCase()) ?? null;
            if (effectiveType && !decl.isArray) {
                const typeMap: Record<string, VbaVarType> = {
                    'byte': 'Byte', 'integer': 'Integer', 'long': 'Long',
                    'single': 'Single', 'double': 'Double', 'currency': 'Currency',
                    'longlong': 'LongLong',
                    'longptr': 'LongPtr',
                    'string': 'String', 'boolean': 'Boolean', 'date': 'Date',
                };
                const mapped = typeMap[effectiveType.toLowerCase()];
                if (mapped) {
                    this.env.setVariableType(varName, { vbaType: mapped });
                }
            }

            // For static variables, restore persisted value if available
            if (isStaticDecl && this.staticVarStore.has(staticKey)) {
                this.env.setLocally(varName, this.staticVarStore.get(staticKey));
                this.staticVarsInCurrentProc.add(varKey);
                continue;
            }

            let initialValue: any = vbaEmpty;
            // Typed numeric/string variables get VBA-spec default values (also applies Def-Directive type)
            if (effectiveType) {
                const t = effectiveType.toLowerCase();
                if (['integer', 'long', 'single', 'double', 'currency', 'byte', 'longlong', 'longptr'].includes(t)) {
                    initialValue = 0;
                } else if (t === 'string') {
                    initialValue = '';
                } else if (t === 'boolean') {
                    initialValue = 0; // vbaFalse
                }
            }
            if (decl.isArray) {
                if (decl.arrayBounds && decl.arrayBounds.length > 0) {
                    initialValue = this.createMultiDimArray(decl.arrayBounds, initialValue);
                    (initialValue as any).vbaFixed = true;
                } else {
                    initialValue = [];
                    (initialValue as any).vbaBase = this.arrayBase;
                    (initialValue as any).vbaFixed = false;
                }
                // UDT 型配列: ReDim 時に要素を初期化できるよう型名を保持する
                if (decl.objectType && this.env.getType(decl.objectType)) {
                    (initialValue as any).__vbaElementTypeName__ = decl.objectType;
                    // B-4: For fixed-size UDT arrays, initialize each element with a proper UDT instance
                    if ((initialValue as any).vbaFixed) {
                        this.fillArrayWithUdtInstances(initialValue, decl.objectType);
                    }
                }
            } else if (decl.isNew && decl.objectType === 'Collection') {
                initialValue = new VbaCollection();
            } else if (decl.isNew && decl.objectType && (
                this.classDefinitions.has(decl.objectType.toLowerCase()) ||
                this.externalObjectFactories.has(decl.objectType.toLowerCase())
            )) {
                // Auto-Instantiation: 宣言時点ではインスタンス化せず、placeholder を入れる。
                // メンバアクセスやメソッド呼び出し時に遅延インスタンス化される。
                initialValue = createAutoInstancePlaceholder(decl.objectType);
                this.autoInstanceVars.set(varName.toLowerCase(), decl.objectType);
            } else if (decl.objectType) {
                const t = decl.objectType.toLowerCase();
                if (!['integer', 'long', 'single', 'double', 'currency', 'byte', 'string', 'boolean', 'longlong', 'longptr'].includes(t)) {
                    if (this.env.getType(decl.objectType)) {
                        // UDT 型: 各メンバを既定値で初期化したインスタンスを生成
                        initialValue = this.instantiateType(decl.objectType);
                    } else if (
                        this.classDefinitions.has(t) ||
                        this.externalObjectFactories.has(t) ||
                        t === 'object' || t === 'collection'
                    ) {
                        // クラスオブジェクト型（New なし）: VBA 仕様により Nothing
                        initialValue = vbaNothing;
                    } else if (t !== 'variant' && t !== 'date') {
                        // 外部COMオブジェクト型（Worksheet, Workbook 等）: VBA 仕様により Nothing
                        initialValue = vbaNothing;
                    }
                    // Variant / Date は vbaEmpty のまま
                }
            }
            // Inside a procedure, declare locally to avoid corrupting outer scopes in recursive calls.
            // At module level (no currentProcedureName), use set() so the value lands in the global env.
            // Detect duplicate Dim in same scope (VBA compile error).
            // Same Dim stmt re-executed in a loop is NOT a duplicate — check by AST node identity.
            if (this.currentProcedureName && this.env.hasOwnVariable(varKey)) {
                const prevNode = this.env.getRegisteredDimStmt(varKey);
                if (prevNode !== stmt) {
                    this.throwVbaError(VbaErrorCode.INVALID_PROCEDURE_CALL,
                        `Variable '${varName}' is already declared in this scope (duplicate declaration)`);
                }
                // Same stmt re-executed (loop body) — skip re-initialization
                continue;
            }
            this.env.registerDimStmt(varKey, stmt);
            if (this.currentProcedureName) {
                // Tier 1: proc-local variable
                this.env.setLocally(varName, initialValue);
            } else {
                // Module-level variable — 単一コピーで管理（二重登録しない）
                // Public/Friend → globalEnv (Tier 3): 全モジュールのチェーンから到達可能
                // Private/Dim with module name → moduleEnv (Tier 2): 自モジュールのチェーンでのみ発見
                // Private/Dim without module name → globalEnv (後方互換: setSourceModule なし旧 API)
                const modName = this.currentSourceModule || 'module1';
                if (stmt.scope === 'public' || stmt.scope === 'friend' || !this.currentSourceModule) {
                    this.env.setLocally(varName, initialValue);
                } else {
                    this.getOrCreateModuleEnv(this.currentSourceModule).setLocally(varName, initialValue);
                }
                this.registerModuleVar(modName, varName);
            }
            if (isStaticDecl) {
                this.staticVarsInCurrentProc.add(varKey);
            }
        }
    }

    private evaluateTypeDeclaration(stmt: TypeDeclaration) {
        this.env.setType(stmt.name, stmt.members);
    }

    private evaluateOptionCompareStatement(stmt: OptionCompareStatement) {
        this.comparisonMode = stmt.mode;
    }

    private evaluateOptionExplicitStatement(_stmt: OptionExplicitStatement) {
        // No-op at runtime. Parser should handle validation.
    }

    private evaluateOptionBaseStatement(stmt: OptionBaseStatement) {
        this.arrayBase = stmt.base;
    }

    private evaluateDefDirective(stmt: DefDirective) {
        for (const { from, to } of stmt.ranges) {
            const fromCode = from.charCodeAt(0);
            const toCode   = to.charCodeAt(0);
            for (let c = fromCode; c <= toCode; c++) {
                this.defTypeMap.set(String.fromCharCode(c), stmt.vbaType);
            }
        }
    }

    private evaluateOptionPrivateModuleStatement(_stmt: OptionPrivateModuleStatement) {
        // No-op for now. Affects visibility in multi-module environment.
    }

    public registerClass(name: string, classDef: ClassDeclaration) {
        this.classDefinitions.set(name.toLowerCase(), classDef);
    }

    private evaluateClassDeclaration(stmt: ClassDeclaration) {
        this.registerClass(stmt.name, stmt);
    }

    private instantiateClass(className: string): any {
        // 外部登録ファクトリ（class 名で別名登録されたもの）を先に確認
        // - registerComObject(factory) で factory().__progId__ が登録されている
        // - これにより VBA の「参照設定」相当（New ClassName / Dim x As ClassName）が動く
        const factory = this.externalObjectFactories.get(className.toLowerCase());
        if (factory) return factory();

        const classDef = this.classDefinitions.get(className.toLowerCase());
        if (!classDef) {
            this.throwVbaError(VbaErrorCode.ACTIVEX_CANT_CREATE_OBJECT, `Class '${className}' not found`);
        }

        return this.createInstanceFromDef(classDef!);
    }

    private createInstanceFromDef(classDef: ClassDeclaration): any {
        // Create instance environment rooted at the global env
        const instanceEnv = new Environment(this.env);

        // Initialize public/private fields with default values
        for (const fieldDecl of classDef.fields) {
            for (const decl of fieldDecl.declarations) {
                let defaultVal: any = vbaEmpty;
                const mt = (decl.objectType || '').toLowerCase();
                // Dim 変数の既定値判定（evaluateVariableDeclaration）と同じ型集合に揃える。
                // 元々は string/integer/long/double/single の数種類しか見ておらず、
                // currency/byte/longlong/longptr は 0 にならず、boolean も False(0) では
                // なく Empty のままになっていた（Date/Variant は Dim 側でも Empty のままが
                // 既定の挙動のため対象外）。
                if (mt === 'string') defaultVal = '';
                else if (['integer', 'long', 'single', 'double', 'currency', 'byte', 'longlong', 'longptr'].includes(mt)) defaultVal = 0;
                else if (mt === 'boolean') defaultVal = 0; // vbaFalse
                else if (decl.objectType && this.env.getType(decl.objectType)) {
                    // UDT (Type ... End Type) フィールド: 既定値では Empty のままになり、
                    // Class_Initialize 等でのメンバー代入が Error 91 になっていた。
                    // Dim 変数の UDT 初期化と同じ instantiateType() を使う。
                    defaultVal = this.instantiateType(decl.objectType);
                } else if (decl.objectType && (
                    this.classDefinitions.has(mt) || this.externalObjectFactories.has(mt) ||
                    mt === 'object' || mt === 'collection'
                )) {
                    // オブジェクト型（クラス名・Object・Collection・WithEvents 含む）フィールドは
                    // New を伴わない宣言の既定値が Empty のままになっていた。実 VBA では
                    // オブジェクト参照型は未代入時 Nothing になる（Empty ではない）ため、
                    // `Is Nothing` 判定や WithEvents バインディングの前提が崩れていた。
                    // Dim 変数の既定値判定（evaluateVariableDeclaration）と同じ扱いに揃える。
                    defaultVal = vbaNothing;
                }
                if (decl.isArray) {
                    if (decl.arrayBounds && decl.arrayBounds.length > 0) {
                        defaultVal = this.createMultiDimArray(decl.arrayBounds, defaultVal);
                        (defaultVal as any).vbaFixed = true;
                    } else {
                        defaultVal = [];
                        (defaultVal as any).vbaBase = this.arrayBase;
                        (defaultVal as any).vbaFixed = false;
                    }
                    if (decl.objectType && this.env.getType(decl.objectType)) {
                        (defaultVal as any).__vbaElementTypeName__ = decl.objectType;
                        if ((defaultVal as any).vbaFixed) {
                            this.fillArrayWithUdtInstances(defaultVal, decl.objectType);
                        }
                    }
                }
                instanceEnv.setLocally(decl.name.name, defaultVal);
                if (decl.isWithEvents) instanceEnv.setWithEvents(decl.name.name);
            }
        }

        const instance: any = {
            __vbaClass__: true,
            __className__: classDef.name,
            __vbaTypeName__: classDef.name,
            __classDef__: classDef,
            __instanceEnv__: instanceEnv,
            __events__: new Map<string, ((...args: any[]) => void)[]>(),
        };

        // Initialize Events
        for (const stmt of classDef.body) {
            if (stmt.type === 'EventDeclaration') {
                const eventDecl = stmt as EventDeclaration;
                instance.__events__.set(eventDecl.name.name.toLowerCase(), []);
            }
        }

        // Set Me in instance env pointing to the instance itself
        instanceEnv.setLocally('Me', instance);

        // B-1: Evaluate class-level Const declarations into instanceEnv
        {
            const prevEnv = this.env;
            this.env = instanceEnv;
            for (const stmt of classDef.body) {
                if (stmt.type === 'ConstDeclaration') {
                    const constVal = this.evaluateConstValue((stmt as ConstDeclaration).value);
                    instanceEnv.setConstant((stmt as ConstDeclaration).name.name, constVal);
                }
            }
            this.env = prevEnv;
        }

        // Call Class_Initialize if defined
        const initProc = classDef.procedures.find(p => p.name.name.toLowerCase() === 'class_initialize');
        if (initProc) {
            this.callClassMethod(instance, initProc, []);
        }

        return instance;
    }

    /**
     * __mocks__ VBA モジュール評価後に呼び出す。
     * classesBefore に含まれない（新しく追加された）クラス定義を
     * classDefinitions から externalObjectFactories へ昇格させる。
     * externalObjectFactories は instantiateClass で classDefinitions より先に参照されるため、
     * ビルトインクラスやユーザー定義クラスより確実に優先される。
     */
    public promoteMockVbaClasses(classesBefore: ReadonlySet<string>): void {
        for (const [name, classDef] of this.classDefinitions) {
            if (!classesBefore.has(name)) {
                const def = classDef;
                this.externalObjectFactories.set(name, () => this.createInstanceFromDef(def));
                this.classDefinitions.delete(name);
            }
        }
    }

    /** classDefinitions に登録済みのクラス名一覧（小文字）を返す。 */
    public getRegisteredClassNames(): Set<string> {
        return new Set(this.classDefinitions.keys());
    }

    /**
     * JS/TS モックから呼び出す。指定した名前でグローバル env を上書きする。
     * registerStandardLibrary より後に呼ぶことでビルトインを差し替えられる。
     */
    public setBuiltinOverride(name: string, fn: any): void {
        this.env.set(name.toLowerCase(), fn);
    }

    /**
     * 引数の数を検証する汎用ロジック。`ProcedureDeclaration.parameters`（ユーザー定義 Proc）と
     * `BuiltinParamSpec[]`（組み込み関数）の両方をこの最小形（`ArgBinderParam[]`）に変換して
     * 共有する。VBA の `Optional` 引数の判定基準は、ユーザー定義 Proc では
     * `!isOptional && defaultValue == null`、組み込み関数では `!optional` であり、
     * その違いは呼び出し側で吸収する（このメソッド自体は判定済みの `isOptional` を見るだけ）。
     */
    private checkArgCountGeneric(params: ArgBinderParam[], providedCount: number): void {
        const hasParamArray = params.some(p => p.isParamArray);
        if (hasParamArray) return;

        const maxParams = params.length;
        const minParams = params.filter(p => !p.isOptional).length;

        if (providedCount > maxParams) {
            this.throwVbaError(VbaErrorCode.WRONG_NUMBER_OF_ARGUMENTS, 'Wrong number of arguments or invalid property assignment');
        }
        if (providedCount < minParams) {
            this.throwVbaError(VbaErrorCode.ARGUMENT_NOT_OPTIONAL, 'Argument not optional');
        }
    }

    private checkArgCount(proc: ProcedureDeclaration, args: any[]): void {
        // 「required」とみなされない（= minParams に数えない）のは isOptional または
        // defaultValue を持つパラメーター。checkArgCountGeneric 側の isOptional フラグに変換する。
        const params: ArgBinderParam[] = proc.parameters.map(p => ({
            isOptional: !!p.isOptional || p.defaultValue != null,
            isParamArray: !!p.isParamArray,
        }));
        this.checkArgCountGeneric(params, args.length);
    }

    /**
     * 省略スロット（`Foo(1, , 3)` の中間カンマで作られる `MissingArgument` ノード）が、
     * Optional でも defaultValue もない必須パラメーターの位置に来ている場合に検出する。
     * VBA 本来はコンパイルエラーだが、本インタープリターでは他の引数数エラーと同様
     * 実行時エラー（449 Argument not optional）として扱う。
     * 名前付き引数は gap の対象外のため、`argExprs[i]` を `params[i]` に素直に対応させてよい
     * （gap は常に純粋な位置引数の文脈でのみ発生する）。
     */
    private checkNoGapOnRequiredParam(params: Parameter[], argExprs: Expression[]): void {
        for (let i = 0; i < argExprs.length && i < params.length; i++) {
            if (argExprs[i].type !== 'MissingArgument') continue;
            const p = params[i];
            if (!p.isOptional && p.defaultValue == null && !p.isParamArray) {
                this.throwVbaError(VbaErrorCode.ARGUMENT_NOT_OPTIONAL, 'Argument not optional');
            }
        }
    }

    /** `argExprs` を名前付き/位置引数に振り分けて評価する（組み込み関数バインダー2種の共通処理）。 */
    private splitCallArgs(argExprs: Expression[]): { namedArgs: Map<string, any>; positionalArgs: any[] } {
        const namedArgs = new Map<string, any>();
        const positionalArgs: any[] = [];
        for (const argExpr of argExprs) {
            if (argExpr.type === 'NamedArgument') {
                const namedArg = argExpr as NamedArgument;
                namedArgs.set(namedArg.name.toLowerCase(), this.evaluateExpression(namedArg.value));
            } else if (argExpr.type === 'MissingArgument') {
                positionalArgs.push(undefined);
            } else {
                positionalArgs.push(this.resolveAutoInstance(argExpr, this.evaluateExpression(argExpr)));
            }
        }
        return { namedArgs, positionalArgs };
    }

    /**
     * `BuiltinParamSpec[]`（必須引数のあとに Optional 引数が続く通常形）を使って、
     * 組み込み関数呼び出しの引数を検証・解決する。ByRef・localEnv は組み込み関数には
     * 存在しないため、ユーザー定義 Proc 用のバインドループより意図的に単純にしている。
     */
    private bindCallArguments(spec: BuiltinParamSpec[], argExprs: Expression[]): any[] {
        const { namedArgs, positionalArgs } = this.splitCallArgs(argExprs);
        const params: ArgBinderParam[] = spec.map(p => ({ isOptional: !!p.optional, isParamArray: !!p.isParamArray }));
        this.checkArgCountGeneric(params, positionalArgs.length + namedArgs.size);

        const result: any[] = [];
        for (let i = 0; i < spec.length; i++) {
            const p = spec[i];
            if (p.isParamArray) {
                result.push(...positionalArgs.slice(i));
                break;
            }
            const nameLower = p.name.toLowerCase();
            if (namedArgs.has(nameLower)) {
                result.push(namedArgs.get(nameLower));
            } else if (i < positionalArgs.length) {
                result.push(positionalArgs[i]);
            } else {
                result.push(undefined); // JS 側のデフォルト引数構文に委ねる
            }
        }
        return result;
    }

    /**
     * `BuiltinOverload[]`（VBA にはない、組み込み関数専用のオーバーロード機構）を使って、
     * 引数の数・名前から該当する1つの「形」を選び、その順序で位置引数配列を再構築する。
     *
     * 名前付き引数を使わない（= 全て位置引数の）呼び出しは、個数の整合性チェックのみ行い、
     * 引数配列はそのまま無加工で返す — 関数本体（例: InStr の `typeof args[0]==='number'`
     * 判定）が今までどおり実際の解釈を行う。名前付き引数が1つでもある場合のみ、該当する
     * オーバーロードを選んで引数を並べ替える。
     */
    private bindOverloadedCallArguments(name: string, overloads: BuiltinOverload[], argExprs: Expression[]): any[] {
        const { namedArgs, positionalArgs } = this.splitCallArgs(argExprs);
        const totalCount = positionalArgs.length + namedArgs.size;
        const arities = overloads.map(o => o.params.length);

        if (namedArgs.size === 0) {
            if (!arities.includes(totalCount)) {
                if (totalCount < Math.min(...arities)) {
                    this.throwVbaError(VbaErrorCode.ARGUMENT_NOT_OPTIONAL, 'Argument not optional');
                }
                this.throwVbaError(VbaErrorCode.WRONG_NUMBER_OF_ARGUMENTS, 'Wrong number of arguments or invalid property assignment');
            }
            return positionalArgs;
        }

        const candidates = overloads.filter(o => {
            if (o.params.length !== totalCount) return false;
            const paramNames = o.params.map(p => p.name.toLowerCase());
            for (const namedKey of namedArgs.keys()) {
                if (!paramNames.includes(namedKey)) return false;
            }
            // 位置引数が占める先頭スロットの名前が、名前付き引数と重複していないこと
            for (let i = 0; i < positionalArgs.length; i++) {
                if (namedArgs.has(paramNames[i])) return false;
            }
            // 残り（位置引数で埋まらないスロット）が全て名前付き引数で埋まること
            for (let i = positionalArgs.length; i < o.params.length; i++) {
                if (!namedArgs.has(paramNames[i])) return false;
            }
            return true;
        });

        if (candidates.length === 0) {
            if (totalCount < Math.min(...arities)) {
                this.throwVbaError(VbaErrorCode.ARGUMENT_NOT_OPTIONAL, 'Argument not optional');
            }
            this.throwVbaError(VbaErrorCode.WRONG_NUMBER_OF_ARGUMENTS, 'Wrong number of arguments or invalid property assignment');
        }
        if (candidates.length > 1) {
            // registerOverloadedBuiltin の登録時チェックで弾かれるはずなので、ここに来るのは
            // スペックの作成ミス（開発時バグ）。VBA エラーではなく内部エラーとして投げる。
            throw new Error(`Ambiguous overload for built-in '${name}' with ${totalCount} argument(s)`);
        }

        const overload = candidates[0];
        const result: any[] = [];
        for (let i = 0; i < overload.params.length; i++) {
            if (i < positionalArgs.length) {
                result.push(positionalArgs[i]);
            } else {
                result.push(namedArgs.get(overload.params[i].name.toLowerCase()));
            }
        }
        return result;
    }

    /**
     * VBA 構文から JS 関数を呼ぶすべての箇所（グローバル組み込み関数・`VBA.Func(...)` 修飾呼び出し・
     * `obj.Method(...)` メソッド呼び出し・式の結果を呼ぶ汎用フォールバック）で共有する引数解決。
     * `fn` に `__vbaOverloads__`/`__vbaParamSpec__` が付いていれば新しい検証・名前解決を行い、
     * 付いていなければ（未移行の組み込み関数）今までどおり単純な位置引数評価のみを行う。
     */
    private resolveCallArgs(fn: Function, argExprs: Expression[], nameForError: string): any[] {
        const overloads = (fn as any).__vbaOverloads__ as BuiltinOverload[] | undefined;
        if (overloads) {
            return this.bindOverloadedCallArguments(nameForError, overloads, argExprs);
        }
        const spec = (fn as any).__vbaParamSpec__ as BuiltinParamSpec[] | undefined;
        if (spec) {
            return this.bindCallArguments(spec, argExprs);
        }
        return argExprs.map(a => this.resolveAutoInstance(a, this.evaluateExpression(a)));
    }

    /**
     * 括弧無し参照（`Now`、`Rnd` など）時に自動呼び出ししてよいか
     * （= 必須引数が0個か）を `__vbaParamSpec__`/`__vbaOverloads__` から判定する。
     * どちらも付いていない（未移行の組み込み関数）場合は対象外。
     */
    private isAutoCallable(fn: Function): boolean {
        const overloads = (fn as any).__vbaOverloads__ as BuiltinOverload[] | undefined;
        if (overloads) {
            return overloads.some(o => o.params.length === 0);
        }
        const spec = (fn as any).__vbaParamSpec__ as BuiltinParamSpec[] | undefined;
        if (spec) {
            return !spec.some(p => p.isParamArray) && spec.every(p => p.optional);
        }
        return false;
    }

    private callClassMethod(instance: any, proc: ProcedureDeclaration, args: any[]): any {
        const instanceEnv = instance.__instanceEnv__ as Environment;
        const localEnv = new Environment(instanceEnv);

        // Make Me available as the current instance
        localEnv.setLocally('Me', instance);

        // Validate argument count
        this.checkArgCount(proc, args);

        // Map arguments to parameters
        for (let i = 0; i < proc.parameters.length; i++) {
            const param = proc.parameters[i];
            const paramName = param.name;

            if (param.isParamArray) {
                const remainingArgs = args.slice(i);
                (remainingArgs as any).vbaBase = 0;
                localEnv.setLocally(paramName, remainingArgs);
                break;
            }

            let argValue: any;
            if (i < args.length) {
                argValue = args[i];
            } else if (param.defaultValue) {
                argValue = this.evaluateExpression(param.defaultValue);
            } else {
                // checkArgCount済みかつパラメーター順序（必須が先・Optionalが後）が保証されているため、
                // ここに来る時点で param は必ず Optional（defaultValue なし）。
                argValue = vbaMissing;
            }
            localEnv.setLocally(paramName, argValue);
        }

        if (proc.isFunction || proc.isProperty) {
            localEnv.setLocally(proc.name.name, vbaEmpty);
            // 配列を返す関数（As String() 等）は戻り値変数をスカラー型として
            // coerce してはいけない（CStr() 等が配列に適用され壊れるため）
            if (proc.returnType && !proc.returnsArray) {
                const retTypeMap: Record<string, VbaVarType> = {
                    'byte': 'Byte', 'integer': 'Integer', 'long': 'Long',
                    'single': 'Single', 'double': 'Double', 'currency': 'Currency',
                    'longlong': 'LongLong', 'longptr': 'LongPtr',
                    'string': 'String', 'boolean': 'Boolean', 'date': 'Date',
                };
                const mapped = retTypeMap[proc.returnType.toLowerCase()];
                if (mapped) localEnv.setVariableType(proc.name.name, { vbaType: mapped });
            }
        }

        const previousEnv = this.env;
        const previousErrorHandler = this.errorHandlerLabel;
        const previousErrorHandlingMode = this.errorHandlingMode;
        const previousIsInErrorHandler = this.isInErrorHandler;
        const previousLastErrorIndex = this.lastErrorIndex;
        const previousProcBody = this.currentProcBody;
        const previousProcedureName = this.currentProcedureName;
        const previousProcedureType = this.currentProcedureType;
        const previousProcIsStatic = this.currentProcIsStatic;
        const previousStaticVars = this.staticVarsInCurrentProc;
        const previousNewOwnedCM = this._currentNewOwned;
        this.env = localEnv;
        this.errorHandlerLabel = null;
        this.errorHandlingMode = 'None';
        this.isInErrorHandler = false;
        this.lastErrorIndex = null;
        this.currentProcBody = proc.body;
        this.currentProcedureName = proc.name.name;
        this.currentProcedureType = proc.propertyType || (proc.isFunction ? 'function' : 'sub');
        this.currentProcIsStatic = false;
        this.staticVarsInCurrentProc = new Set();
        this._currentNewOwned = new Set();

        try {
            this.executeStatements(proc.body, 0);
        } catch (e: any) {
            if (e && e.type === 'Exit') {
                if (
                    (e.target === 'Function' && proc.isFunction) ||
                    (e.target === 'Sub' && !proc.isFunction && !proc.isProperty) ||
                    (e.target === 'Property' && proc.isProperty)
                ) {
                    // valid exit
                } else {
                    throw e;
                }
            } else {
                throw e;
            }
        } finally {
            this.env = previousEnv;
            this.errorHandlerLabel = previousErrorHandler;
            this.errorHandlingMode = previousErrorHandlingMode;
            this.isInErrorHandler = previousIsInErrorHandler;
            this.lastErrorIndex = previousLastErrorIndex;

            // Scope exit: fire Class_Terminate for local VBA objects created with `New` in
            // this scope that were not returned to the caller.
            {
                const owned = this._currentNewOwned;
                if (owned && owned.size > 0) {
                    const paramNames = new Set(proc.parameters.map(p => p.name.toLowerCase()));
                    paramNames.add('me');
                    const returnVarName = (proc.isFunction || proc.isProperty) ? proc.name.name.toLowerCase() : null;
                    const returnVal = returnVarName !== null ? localEnv.get(returnVarName) : undefined;
                    for (const [name, val] of localEnv.getLocalVariables()) {
                        if (paramNames.has(name)) continue;
                        if (returnVarName !== null && name === returnVarName) continue;
                        if (val !== null && val !== undefined && val === returnVal) continue;
                        if (!owned.has(val)) continue;
                        this.triggerTerminate(val);
                    }
                }
            }

            this.currentProcBody = previousProcBody;
            this.currentProcedureName = previousProcedureName;
            this.currentProcedureType = previousProcedureType;
            this.currentProcIsStatic = previousProcIsStatic;
            this.staticVarsInCurrentProc = previousStaticVars;
            this._currentNewOwned = previousNewOwnedCM;
        }

        if (proc.isFunction || proc.isProperty) {
            return localEnv.get(proc.name.name.toLowerCase());
        }
        return vbaEmpty;
    }

    private evaluateEnumDeclaration(stmt: EnumDeclaration) {
        let currentValue = 0;
        const enumObj: any = {};
        for (const member of stmt.members) {
            if (member.value) {
                currentValue = this.evaluateExpression(member.value);
            }
            const memberName = member.name.name;
            // Set directly in environment for flat access (common in VBA)
            this.env.set(memberName, currentValue);
            // Also store in enum object for EnumName.MemberName access
            enumObj[memberName] = currentValue;
            currentValue++;
        }
        // Register the Enum name itself
        this.env.set(stmt.name.name, enumObj);
    }

    private evaluateCallStatement(stmt: CallStatement) {
        this.evaluateExpression(stmt.expression);
    }

    /**
     * 全モジュールロード後に呼ぶ。モジュールレベル定数を依存グラフでトポロジカルソートして
     * 正しい順序で再評価する。循環参照はエラーとして検出する。
     */
    /**
     * Pass 2: 全モジュールのロード完了後に識別子を解決する。
     * モジュールレベル定数を依存グラフ順に再評価し、全モジュール名が確定した状態で
     * Option Explicit の精密チェックを再実行する。
     * evalVBASingle / evalVBAModules から必ず1回だけ呼ぶこと。
     */
    public resolveIdentifiers(modules: Array<{ ast: Program; moduleName: string }>): void {
        try {
        // 全モジュールレベル定数を「module:name」修飾キーで収集する。
        const allConsts = new Map<string, { stmt: ConstDeclaration; moduleName: string; name: string }>();
        const collect = (stmt: ConstDeclaration, moduleName: string) => {
            const name = stmt.name.name.toLowerCase();
            allConsts.set(`${moduleName.toLowerCase()}:${name}`, { stmt, moduleName, name });
        };
        for (const { ast, moduleName } of modules) {
            for (const stmt of ast.body) {
                if (stmt.type === 'ConstDeclaration') {
                    collect(stmt as ConstDeclaration, moduleName);
                } else if (stmt.type === 'ClassDeclaration') {
                    for (const member of (stmt as ClassDeclaration).body) {
                        if (member.type === 'ConstDeclaration') {
                            collect(member as ConstDeclaration, moduleName);
                        }
                    }
                }
            }
        }

        // 依存グラフを構築。非修飾名は同一モジュール優先、なければ他モジュールを探す。
        // 他モジュールの Private Const を参照するとエラー（VBA 仕様: Private は同一モジュール内のみ）。
        const resolveDep = (depName: string, fromModule: string): string | undefined => {
            const sameModule = `${fromModule.toLowerCase()}:${depName}`;
            if (allConsts.has(sameModule)) return sameModule;
            for (const key of allConsts.keys()) {
                if (key.endsWith(`:${depName}`)) return key;
            }
            return undefined;
        };
        const deps = new Map<string, Set<string>>();
        for (const [qkey, { stmt, moduleName }] of allConsts) {
            const exprDeps = this.collectConstExprDeps(stmt.value);
            const resolved = new Set<string>();
            for (const d of exprDeps) {
                const r = resolveDep(d, moduleName);
                if (r) {
                    // 他モジュールの Private Const への参照はエラー
                    const { moduleName: depModule, stmt: depStmt } = allConsts.get(r)!;
                    if (depModule.toLowerCase() !== moduleName.toLowerCase() &&
                        (depStmt as any).scope === 'private') {
                        throw new Error(
                            `Constant expression required: '${stmt.name.name}' references private constant '${depStmt.name.name}' in module '${depModule}'`
                        );
                    }
                    resolved.add(r);
                }
            }
            deps.set(qkey, resolved);
        }

        // グローバルトポロジカルソート（循環参照を検出）
        const order = this.topologicalSortConsts(allConsts, deps);

        // 正しい順序で評価（evaluateConstValue が resolveConstIdent を使うため未定義名はエラー）
        for (const qkey of order) {
            const { stmt, moduleName } = allConsts.get(qkey)!;
            const prev = this.currentSourceModule;
            this.currentSourceModule = moduleName;
            this.evaluateConstDeclaration(stmt);
            this.currentSourceModule = prev;
        }

        // Pass 1 で延期した配列境界付き VariableDeclaration を Const 解決後に評価する。
        // evaluateModule はバッチ中は配列境界付き Dim を pendingArrayDecls に退避している。
        for (const { stmt, moduleName } of this.pendingArrayDecls) {
            const prev = this.currentSourceModule;
            this.currentSourceModule = moduleName;
            for (const decl of stmt.declarations) {
                if (!decl.arrayBounds) continue;
                for (const bound of decl.arrayBounds) {
                    this.validateConstantExpr(bound.upper);
                    if (bound.lower) this.validateConstantExpr(bound.lower);
                }
            }
            this.evaluateVariableDeclaration(stmt);
            this.currentSourceModule = prev;
        }
        this.pendingArrayDecls = [];

        // 同一モジュール内の重複プロシージャ名チェック（Pass 2）
        for (const { ast, moduleName } of modules) {
            const seen = new Set<string>();
            for (const stmt of ast.body) {
                if (stmt.type !== 'ProcedureDeclaration') continue;
                const proc = stmt as ProcedureDeclaration;
                const key = (proc.isProperty && proc.propertyType)
                    ? `${proc.name.name.toLowerCase()}:${proc.propertyType}`
                    : proc.name.name.toLowerCase();
                if (seen.has(key)) {
                    const line = proc.name.loc?.start.line ?? 0;
                    this.throwCompileError(VbaErrorCode.INVALID_PROCEDURE_CALL, `Duplicate procedure name '${proc.name.name}' in module '${moduleName}'`, line);
                }
                seen.add(key);
            }
        }

        // モジュールレベル宣言の位置チェック（Pass 2）
        // 標準 VBA 仕様: Dim/Const/Type/Enum 等のモジュールレベル宣言も含め、プロシージャ
        // （End Sub/End Function/End Property）の後に書けるのはコメントのみ。それ以外を
        // 書くと「End Sub、End Function または End Property 以降には、コメントのみが
        // 記述できます」というコンパイルエラーになる。
        // vba-runner 拡張: evalVBASingle/evalVBAModules 等は allowTopLevelStatements が
        // true（デフォルト）になっており、REPL・テストスクリプト用に Dim も含めた
        // モジュールレベル文をプロシージャの前後どこに書いても許容する
        // （先頭にない Dim もモジュールレベル実行文と同様に扱う）。
        // allowTopLevelStatements: false を指定したときのみ標準 VBA 相当の挙動になる。
        for (const { ast } of modules) {
            let seenProcedure = false;
            for (const stmt of ast.body) {
                if (stmt.type === 'ProcedureDeclaration') {
                    seenProcedure = true;
                    continue;
                }
                // vba-runner 拡張: 1つの .bas 内に複数 Class をインライン定義できる
                // （__mocks__.bas 等）。標準 VBA に存在しない構文のため位置チェック対象外。
                if (stmt.type === 'ClassDeclaration') continue;
                if (!seenProcedure) continue;
                if (!this.allowTopLevelStatements) {
                    const line = stmt.loc?.start.line ?? stmt.line ?? 0;
                    this.throwCompileError(VbaErrorCode.INVALID_PROCEDURE_CALL, `Only comments may appear after End Sub, End Function, or End Property`, line);
                }
            }
        }

        // Option Explicit チェックはここ（Pass 2）に一本化している。
        // 全モジュール名を knownModuleNames として渡す精密モードで実行することで、
        // コール式の bare identifier オブジェクト（obj.Method() の obj）が本物のモジュール名か
        // 未宣言変数かを判定できる。
        const knownModuleNames = new Set<string>(
            modules.map(m => m.moduleName.toLowerCase()).filter(n => n !== '')
        );
        // Option Explicit チェック: Pass 2 では純粋な静的 AST 解析結果をキャッシュするだけ。
        // env（ev.set() 等による実行時注入）との突合せは precheckProc で行うことで責務を分離する。
        // resolveIdentifiers が複数回呼ばれても冪等になるよう、毎回クリアして再構築する。
        // 実際の throw は callProcedure（§5.6.10）で行い、
        // コールスタックが積まれた状態でエラーを報告する。
        this.optionExplicitViolations.clear();
        for (const { ast } of modules) {
            const { violatedProcedures } = checkOptionExplicit(ast, knownModuleNames);
            for (const [procName, undeclared] of violatedProcedures) {
                if (undeclared.size > 0) {
                    this.optionExplicitViolations.set(procName, undeclared);
                }
            }
        }

        // バッチロード完了。以降の evaluateModule 呼び出し（evalExpression 等）は即時実行する。
        this.resolveIdentifiersDone = true;

        // Pass 1 でシンボルテーブルに退避したモジュールレベル実行文を順番に実行する。
        // Const・配列 Dim がすべて確定した状態で実行されるため定数参照が正しく解決される。
        for (const { moduleName, stmts } of this.pendingTopLevel) {
            const prev = this.currentSourceModule;
            this.currentSourceModule = moduleName;
            for (const stmt of stmts) {
                this.evaluateStatement(stmt);
            }
            this.currentSourceModule = prev;
        }
        this.pendingTopLevel = [];

        } catch (e: any) {
            if (e._precheckRaw) {
                const line = e.vbaLine;
                const mod = e.vbaModule ?? this.executingModuleName ?? this.currentSourceModule ?? null;
                const msg = line !== undefined ? `Compile error: ${e.message} (line ${line})` : `Compile error: ${e.message}`;
                const formatted: any = new Error(msg);
                formatted.type = 'VbaError';
                formatted.number = e.number;
                formatted.vbaLine = line;
                formatted.vbaModule = mod;
                formatted.vbaStack = [...this.vbaCallStack].reverse();
                throw formatted;
            }
            throw e;
        }
    }

    /** @deprecated Use {@link resolveIdentifiers} instead. */
    public reEvaluateModuleConstsAll(modules: Array<{ ast: Program; moduleName: string }>): void {
        this.resolveIdentifiers(modules);
    }

    private collectConstExprDeps(expr: Expression): Set<string> {
        const deps = new Set<string>();
        const walk = (e: Expression): void => {
            switch (e.type) {
                case 'Identifier':
                    deps.add((e as Identifier).name.toLowerCase());
                    break;
                case 'BinaryExpression': {
                    const b = e as BinaryExpression;
                    walk(b.left); walk(b.right);
                    break;
                }
                case 'UnaryExpression':
                    walk((e as UnaryExpression).argument);
                    break;
                case 'ParenthesizedExpression':
                    walk((e as ParenthesizedExpression).expression);
                    break;
                case 'CallExpression': {
                    const ce = e as CallExpression;
                    if (ce.callee.type === 'Identifier') {
                        deps.add((ce.callee as Identifier).name.toLowerCase());
                    }
                    for (const arg of ce.args) walk(arg);
                    break;
                }
            }
        };
        walk(expr);
        return deps;
    }

    private topologicalSortConsts(
        consts: Map<string, unknown>,
        deps: Map<string, Set<string>>
    ): string[] {
        type State = 'unvisited' | 'visiting' | 'done';
        const state = new Map<string, State>();
        const order: string[] = [];
        const path: string[] = [];

        for (const name of consts.keys()) state.set(name, 'unvisited');

        const visit = (name: string): void => {
            const s = state.get(name);
            if (s === 'done') return;
            if (s === 'visiting') {
                const cycle = [...path.slice(path.indexOf(name)), name].join(' → ');
                throw new Error(`Circular reference in constant declarations: ${cycle}`);
            }
            state.set(name, 'visiting');
            path.push(name);
            for (const dep of deps.get(name) ?? []) visit(dep);
            path.pop();
            state.set(name, 'done');
            order.push(name);
        };

        for (const name of consts.keys()) visit(name);
        return order;
    }

    /** 単一モジュールの定数を再評価する（後方互換用）。循環検出は行わない。 */
    public reEvaluateModuleConsts(program: Program, moduleName: string): void {
        const prev = this.currentSourceModule;
        this.currentSourceModule = moduleName;
        for (const stmt of program.body) {
            if (stmt.type === 'ConstDeclaration') {
                this.evaluateConstDeclaration(stmt as ConstDeclaration);
            } else if (stmt.type === 'ClassDeclaration') {
                for (const member of (stmt as ClassDeclaration).body) {
                    if (member.type === 'ConstDeclaration') {
                        this.evaluateConstDeclaration(member as ConstDeclaration);
                    }
                }
            }
        }
        this.currentSourceModule = prev;
    }

    /**
     * 定数式中の識別子を解決する。
     * env に見つからない場合は "Constant expression required" エラー。
     * 他モジュールの Public Const はグローバルトポロジカルソートにより評価済みなので参照可能。
     */
    private resolveConstIdent(name: string): any {
        const val = this.env.getConst(name.toLowerCase());
        if (val === undefined) {
            throw new Error(`Constant expression required: '${name}' is not defined`);
        }
        return val;
    }

    /** 定数式を評価する。識別子は resolveConstIdent() で解決し、未定義名はエラー。 */
    private evaluateConstValue(expr: Expression): any {
        this.inConstEval = true;
        try {
            return this.evaluateExpression(expr);
        } finally {
            this.inConstEval = false;
        }
    }

    private evaluateConstDeclaration(stmt: ConstDeclaration) {
        const value = this.evaluateConstValue(stmt.value);
        const name = stmt.name.name;
        if (!this.currentProcedureName && this.currentSourceModule) {
            // モジュールレベル定数は VBA では Public がデフォルト → globalEnv (Tier 3) のみ
            const modName = this.currentSourceModule;
            this.env.setConstant(name, value);
            this.env.setConstant(`${modName}:${name}`, value);
            this.registerModuleVar(modName, name);
        } else {
            this.env.setConstant(name, value);
        }
    }

    private registerModuleVar(moduleName: string, varName: string) {
        const key = moduleName.toLowerCase();
        if (!this.moduleVarRegistry.has(key)) {
            this.moduleVarRegistry.set(key, new Set());
        }
        this.moduleVarRegistry.get(key)!.add(varName.toLowerCase());
    }

    /**
     * `Set <fieldName> = obj`（WithEvents フィールドへの代入）で `<fieldName>_<EventName>`
     * ハンドラーを配線する。bare identifier 代入（`Set Source = New Pub`）・
     * member access 経由の代入（`Set s.Source = New Pub`）の両方から呼ばれる。
     * `classDef`/`instance` はハンドラーの検索先・実行時の `Me` に使う
     * （bare identifier の場合は呼び出し元の `Me`、member access の場合は代入先オブジェクト自身）。
     */
    private bindWithEventsHandlers(fieldName: string, value: any, classDef: ClassDeclaration | undefined, instance: any) {
        for (const eventName of value.__events__.keys()) {
            const handlerName = `${fieldName}_${eventName}`;
            const handlerNameLower = handlerName.toLowerCase();
            const handlers = value.__events__.get(eventName);
            if (classDef) {
                const classProc = classDef.procedures.find(
                    p => p.name.name.toLowerCase() === handlerNameLower
                );
                if (classProc) {
                    const capturedInstance = instance;
                    const capturedProc = classProc;
                    handlers.push((...args: any[]) => this.callClassMethod(capturedInstance, capturedProc, args));
                    continue;
                }
            }
            const handler = this.env.getProcedure(handlerName);
            if (handler) {
                handlers.push((...args: any[]) => this.callProcedure(handlerName, args));
            }
        }
    }

    private evaluateSetStatement(stmt: SetStatement) {
        let value = this.evaluateExpression(stmt.right);

        // Track objects created with `New` in the current proc scope for scope exit cleanup.
        if (stmt.right.type === 'NewExpression' && value && value.__vbaClass__ && this._currentNewOwned) {
            this._currentNewOwned.add(value);
        }

        // VBA requires Set target to be an object (or Nothing)
        if (value !== null && value !== vbaNothing && typeof value !== 'object') {
            this.throwVbaError(VbaErrorCode.OBJECT_REQUIRED, 'Object required');
        }
        // If the right side evaluates to a variable name (string), resolve it
        if (typeof value === 'string' && stmt.right.type === 'Identifier') {
            value = this.env.get(value);
        }

        if (stmt.left.type === 'Identifier') {
            const name = (stmt.left as Identifier).name;
            const procNameLower = name.toLowerCase();

            // Special case: assigning to current function/property name (return value)
            if (this.currentProcedureName && this.currentProcedureName.toLowerCase() === procNameLower) {
                if (this.currentProcedureType === 'function' || this.currentProcedureType === 'get') {
                    this.env.setLocally(name, value);
                    return;
                }
            }

            let oldVal = this.env.get(name);
            // B-3: Only trigger Terminate on explicit Nothing assignment (Set x = Nothing).
            // Without reference counting we cannot know if the old object is still reachable
            // elsewhere, so triggering on every reassignment causes premature Terminate.
            //
            // `Dim w As New T` + `Set w = Nothing` without any prior access: real VBA still
            // instantiates the object (Class_Initialize) before destroying it (Class_Terminate).
            // Resolve the placeholder first so both lifecycle hooks fire correctly.
            if (value === vbaNothing && isAutoInstancePlaceholder(oldVal)) {
                oldVal = this.instantiateClass(oldVal.__className__);
                this.env.set(name, oldVal);
            }
            if (value === vbaNothing && oldVal !== value) this.triggerTerminate(oldVal);

            const proc = this.env.getProcedure(name, 'set');
            if (proc) {
                this.callProcedure(name, [value], 'set');
                return;
            }

            // Auto-Instantiation: `Set x = Nothing` で auto-instance 変数なら placeholder に戻す。
            // 仕様: 再度参照したら新しいインスタンスが生成される。
            const className = this.autoInstanceVars.get(name.toLowerCase());
            if (className && value === vbaNothing) {
                this.env.set(name, createAutoInstancePlaceholder(className));
                return;
            }
            this.env.set(name, value);

            // Check for WithEvents binding
            if (this.env.isWithEvents(name) && value && value.__events__) {
                // Binding: look for VarName_EventName subs — first in class scope, then global
                const me = this.env.get('me');
                this.bindWithEventsHandlers(name, value, me?.__classDef__, me);
            }
        } else if (stmt.left.type === 'MemberExpression') {
            const member = stmt.left as MemberExpression;
            const obj = this.resolveAutoInstance(member.object, this.evaluateExpression(member.object));
            const propName = member.property.name.toLowerCase();
            if (obj && obj.__vbaClass__) {
                const classDef = obj.__classDef__ as ClassDeclaration;
                const instanceEnv = obj.__instanceEnv__ as Environment;
                const oldVal = instanceEnv.get(propName);
                if (value === vbaNothing && oldVal !== value) this.triggerTerminate(oldVal);
                const setter = classDef.procedures.find(
                    p => p.isProperty && p.propertyType === 'set' && p.name.name.toLowerCase() === propName
                );
                if (setter) {
                    this.callClassMethod(obj, setter, [value]);
                } else {
                    instanceEnv.set(propName, value);
                    // obj.Field = New X 形式（外部からの WithEvents フィールド代入）でも
                    // バインドが行われるようにする。bare identifier 代入のみを見ていたため、
                    // member access 経由の代入では handler が一切ワイヤリングされず
                    // イベントが静かに発火しなかった。
                    if (instanceEnv.isWithEvents(propName) && value && value.__events__) {
                        this.bindWithEventsHandlers(propName, value, classDef, obj);
                    }
                }
            } else if (obj && typeof obj === 'object') {
                const oldVal = obj[propName];
                if (value === vbaNothing && oldVal !== value) this.triggerTerminate(oldVal);
                obj[propName] = value;
            } else {
                if (obj === null || obj === undefined || obj === vbaNothing) {
                    this.throwVbaError(VbaErrorCode.OBJECT_VARIABLE_NOT_SET, 'Object variable or With block variable not set');
                } else {
                    this.throwVbaError(VbaErrorCode.OBJECT_REQUIRED, 'Object required');
                }
            }
        } else if (stmt.left.type === 'CallExpression') {
            // Set obj.Item(key) = obj2  or  Set obj(key) = obj2
            const call = stmt.left as CallExpression;
            if (call.callee.type === 'MemberExpression') {
                const memberCallee = call.callee as MemberExpression;
                const obj = this.resolveAutoInstance(memberCallee.object, this.evaluateExpression(memberCallee.object));
                const methodName = memberCallee.property.name.toLowerCase();
                if (obj && obj.__isVbaDict__) {
                    const key = String(this.evaluateExpression(call.args[0]));
                    const oldVal = obj.__map__.get(key);
                    if (value === vbaNothing && oldVal !== value) this.triggerTerminate(oldVal);
                    obj.__map__.set(key, value);
                } else if (obj && obj.__vbaClass__) {
                    const classDef = obj.__classDef__ as ClassDeclaration;
                    const setter = classDef.procedures.find(
                        p => p.isProperty && p.propertyType === 'set' && p.name.name.toLowerCase() === methodName
                    );
                    if (setter) {
                        const argsVals = call.args.map(a => this.evaluateExpression(a));
                        this.callClassMethod(obj, setter, [...argsVals, value]);
                    } else {
                        this.throwVbaError(VbaErrorCode.OBJECT_DOESNT_SUPPORT_PROPERTY, "Object doesn't support this property or method");
                    }
                } else if (obj && typeof obj === 'object') {
                    const key = String(this.evaluateExpression(call.args[0]));
                    obj[key] = value;
                } else {
                    this.throwVbaError(VbaErrorCode.INVALID_PROCEDURE_CALL, 'Invalid procedure call or argument');
                }
            } else if (call.callee.type === 'Identifier') {
                const name = (call.callee as Identifier).name;
                const target = this.env.get(name);
                if (target && target.__isVbaDict__) {
                    const key = String(this.evaluateExpression(call.args[0]));
                    const oldVal = target.__map__.get(key);
                    if (value === vbaNothing && oldVal !== value) this.triggerTerminate(oldVal);
                    target.__map__.set(key, value);
                } else if (target && typeof target === 'object') {
                    const key = String(this.evaluateExpression(call.args[0]));
                    target[key] = value;
                } else {
                    this.throwVbaError(VbaErrorCode.INVALID_PROCEDURE_CALL, 'Invalid procedure call or argument');
                }
            } else {
                this.throwVbaError(VbaErrorCode.INVALID_PROCEDURE_CALL, 'Invalid procedure call or argument');
            }
        } else {
            this.throwVbaError(VbaErrorCode.INVALID_PROCEDURE_CALL, 'Invalid procedure call or argument');
        }
    }

    private evaluateResumeStatement(stmt: ResumeStatement) {
        // Check if there's an active error
        if (this.errObj.number === 0) {
            this.throwVbaError(VbaErrorCode.RESUME_WITHOUT_ERROR, 'Resume without error');
        }

        const target = (stmt.target || '').toLowerCase().trim();
        if (target === '') {
            throw { type: 'Resume', mode: 'Current' };
        } else if (target === 'next') {
            throw { type: 'Resume', mode: 'Next' };
        } else {
            throw { type: 'Resume', mode: 'Label', label: stmt.target };
        }
    }

    private evaluateOnErrorStatement(stmt: OnErrorStatement) {
        const label = (stmt.label || '').toLowerCase().trim();
        if (label === '0') {
            this.errorHandlerLabel = null;
            this.errorHandlingMode = 'None';
        } else if (label === '-1') {
            this.isInErrorHandler = false;
        } else if (label === 'resume next') {
            this.errorHandlerLabel = null;
            this.errorHandlingMode = 'ResumeNext';
        } else {
            this.errorHandlerLabel = stmt.label;
            this.errorHandlingMode = 'GoTo';
        }
    }

    private evaluateOpenStatement(stmt: OpenStatement) {
        const vbaPath = String(this.evaluateExpression(stmt.path));
        const realPath = this.sandbox.toRealPath(vbaPath);
        const fileNum = Number(this.evaluateExpression(stmt.fileNumber));

        if (this.fileHandles.has(fileNum)) {
            this.throwVbaError(VbaErrorCode.FILE_ALREADY_OPEN, "File already open");
        }

        let flags = '';
        switch (stmt.mode) {
            case 'Input': flags = 'r'; break;
            case 'Output': flags = 'w'; break;
            case 'Append': flags = 'a'; break;
            case 'Random':
            case 'Binary':
                if (!this.fs.existsSync(realPath)) {
                    this.fs.writeFileSync(realPath, "");
                }
                flags = 'r+'; break;
        }

        try {
            // Ensure directory exists for write modes
            if (flags === 'w' || flags === 'a') {
                const dir = path.dirname(realPath);
                if (!this.fs.existsSync(dir)) {
                    this.fs.mkdirSync(dir, { recursive: true });
                }
            } else if (flags === 'r' && !this.fs.existsSync(realPath)) {
                this.throwVbaError(VbaErrorCode.FILE_NOT_FOUND, "File not found");
            }

            const fd = this.fs.openSync(realPath, flags);
            this.fileHandles.set(fileNum, {
                fd,
                mode: stmt.mode,
                path: realPath,
                pos: 0
            });
        } catch (e: any) {
            if (e.code === 'ENOENT') this.throwVbaError(VbaErrorCode.FILE_NOT_FOUND, "File not found");
            if (e.code === 'EACCES') this.throwVbaError(VbaErrorCode.PATH_FILE_ACCESS_ERROR, "Path/File access error");
            throw e;
        }
    }

    private evaluateCloseStatement(stmt: CloseStatement) {
        const nums = stmt.fileNumbers.length > 0
            ? stmt.fileNumbers.map(n => Number(this.evaluateExpression(n)))
            : Array.from(this.fileHandles.keys());

        for (const num of nums) {
            const handle = this.fileHandles.get(num);
            if (handle) {
                this.fs.closeSync(handle.fd);
                this.fileHandles.delete(num);
            }
        }
    }

    private evaluatePrintStatement(stmt: PrintStatement) {
        const fileNum = Number(this.evaluateExpression(stmt.fileNumber));
        const handle = this.fileHandles.get(fileNum);
        if (!handle) this.throwVbaError(VbaErrorCode.BAD_FILE_NAME_OR_NUMBER, "Bad file name or number");
        if (handle.mode === 'Input') this.throwVbaError(VbaErrorCode.BAD_FILE_MODE, "Bad file mode");

        let output = "";
        for (const expr of stmt.expressions) {
            if (expr === 'Comma') {
                const currentLen = output.length;
                const target = Math.ceil((currentLen + 1) / 14) * 14;
                output += " ".repeat(target - currentLen);
            } else if (expr === 'Semicolon') {
                // Continue
            } else if (typeof expr === 'object' && expr !== null && 'type' in expr) {
                 if (expr.type === 'Spc') {
                     const n = Number(this.evaluateExpression((expr as any).val));
                     output += " ".repeat(Math.max(0, n));
                 } else if (expr.type === 'Tab') {
                     const n = Number(this.evaluateExpression((expr as any).val));
                     output += " ".repeat(Math.max(0, n - output.length));
                 } else {
                     output += String(this.evaluateExpression(expr as any));
                 }
            } else {
                const val = this.evaluateExpression(expr as any);
                output += String(val === vbaNull ? "Null" : (val === vbaEmpty ? "" : val));
            }
        }

        const last = stmt.expressions[stmt.expressions.length - 1];
        if (last !== 'Semicolon' && last !== 'Comma') {
            output += "\r\n";
        }

        this.fs.writeSync(handle.fd, output);
        handle.pos! += output.length;
    }

    private evaluateLineInputStatement(stmt: LineInputStatement) {
        const fileNum = Number(this.evaluateExpression(stmt.fileNumber));
        const handle = this.fileHandles.get(fileNum);
        if (!handle) this.throwVbaError(VbaErrorCode.BAD_FILE_NAME_OR_NUMBER, "Bad file name or number");

        const buffer = new Uint8Array(1);
        let line = "";
        let bytesRead = 0;

        while (true) {
            bytesRead = this.fs.readSync(handle.fd, buffer, 0, 1, handle.pos ?? null);
            if (bytesRead === 0) break;
            const char = new TextDecoder().decode(buffer.subarray(0, 1));
            handle.pos!++;
            if (char === '\n') break;
            if (char !== '\r') line += char;
        }

        this.env.setLocally(stmt.variable.name, line);
    }

    private evaluatePutStatement(stmt: PutStatement) {
        const fileNum = Number(this.evaluateExpression(stmt.fileNumber));
        const handle = this.fileHandles.get(fileNum);
        if (!handle) this.throwVbaError(VbaErrorCode.BAD_FILE_NAME_OR_NUMBER, "Bad file name or number");

        const data = this.evaluateExpression(stmt.data);
        const s = String(data);
        const buffer = new TextEncoder().encode(s);

        let position: number | null = handle.pos ?? null;
        if (stmt.recordNumber) {
            position = (Number(this.evaluateExpression(stmt.recordNumber)) - 1);
        }

        this.fs.writeSync(handle.fd, buffer, 0, buffer.length, position);
        handle.pos! += buffer.length;
    }

    private vbaWildcardToRegex(pattern: string): RegExp {
        let regexStr = '';
        for (const ch of pattern) {
            if (ch === '*') regexStr += '.*';
            else if (ch === '?') regexStr += '.';
            else if (/[.+^${}()|[\]\\]/.test(ch)) regexStr += '\\' + ch;
            else regexStr += ch;
        }
        return new RegExp('^' + regexStr + '$', 'i');
    }

    private executeKill(vbaPath: string): void {
        if (vbaPath.includes('*') || vbaPath.includes('?')) {
            const norm = vbaPath.replace(/\\/g, '/');
            const lastSlash = norm.lastIndexOf('/');
            const dirPart = lastSlash === -1 ? '' : norm.substring(0, lastSlash);
            const pattern = lastSlash === -1 ? norm : norm.substring(lastSlash + 1);

            const realDir = this.sandbox.toRealPath(dirPart || '');
            let files: string[];
            try {
                files = this.fs.readdirSync(realDir);
            } catch {
                this.throwVbaError(VbaErrorCode.FILE_NOT_FOUND, 'File not found');
                return;
            }

            const regex = this.vbaWildcardToRegex(pattern);
            const matched = files.filter((f: string) => {
                if (!regex.test(f)) return false;
                try { return this.fs.statSync(realDir + '/' + f).isFile(); } catch { return false; }
            });

            if (matched.length === 0) {
                this.throwVbaError(VbaErrorCode.FILE_NOT_FOUND, 'File not found');
                return;
            }
            for (const file of matched) {
                this.fs.unlinkSync(realDir + '/' + file);
            }
        } else {
            const realPath = this.sandbox.toRealPath(vbaPath);
            try {
                this.fs.unlinkSync(realPath);
            } catch {
                this.throwVbaError(VbaErrorCode.FILE_NOT_FOUND, 'File not found');
            }
        }
    }

    private evaluateKillStatement(stmt: KillStatement) {
        const vbaPath = String(this.evaluateExpression(stmt.path));
        this.executeKill(vbaPath);
    }

    private evaluateWriteStatement(stmt: WriteStatement) {
        const fileNum = Number(this.evaluateExpression(stmt.fileNumber));
        const handle = this.fileHandles.get(fileNum);
        if (!handle) this.throwVbaError(VbaErrorCode.BAD_FILE_NAME_OR_NUMBER, `Bad file name or number: #${fileNum}`);

        const output = stmt.items.map(item => {
            const val = this.evaluateExpression(item);
            if (typeof val === 'string') return `"${val}"`;
            if (val instanceof VbaDate) return `#${val.toString()}#`;
            if (val === null) return "#NULL#";
            return String(val);
        }).join(",");

        const lineOutput = output + "\n";
        this.fs.writeSync(handle.fd, lineOutput);
        handle.pos! += lineOutput.length;
    }

    private evaluateInputStatement(stmt: InputStatement) {
        const fileNum = Number(this.evaluateExpression(stmt.fileNumber));
        const handle = this.fileHandles.get(fileNum);
        if (!handle) this.throwVbaError(VbaErrorCode.BAD_FILE_NAME_OR_NUMBER, `Bad file name or number: #${fileNum}`);

        // Simple line-based implementation for now.
        // Real VBA Input # parses delimiters even across lines.
        let content = "";
        const buf = new Uint8Array(1024);
        let bytesRead = 0;
        let readPos = handle.pos || 0;
        while ((bytesRead = this.fs.readSync(handle.fd, buf, 0, 1024, readPos)) > 0) {
            content += new TextDecoder().decode(buf.subarray(0, bytesRead));
            if (content.includes('\n')) break;
            readPos += bytesRead;
        }

        const lineEnd = content.indexOf('\n');
        const line = lineEnd === -1 ? content : content.slice(0, lineEnd);
        handle.pos = (handle.pos || 0) + line.length + (lineEnd === -1 ? 0 : 1);
        if (content[lineEnd - 1] === '\r') {
            // handle CRLF
        }

        const values = line.trim().split(",").map(v => v.trim().replace(/^"|"$/g, ''));
        for (let i = 0; i < stmt.variables.length; i++) {
            if (i < values.length) {
                this.evaluateAssignmentToVariable(stmt.variables[i], values[i]);
            }
        }
    }

    private evaluateGetStatement(stmt: GetStatement) {
        const fileNum = Number(this.evaluateExpression(stmt.fileNumber));
        const handle = this.fileHandles.get(fileNum);
        if (!handle) this.throwVbaError(VbaErrorCode.BAD_FILE_NAME_OR_NUMBER, `Bad file name or number: #${fileNum}`);

        // Basic implementation: read up to 1024 bytes or until EOF
        const buffer = new Uint8Array(1024);
        let position: number | null = handle.pos ?? null;
        if (stmt.recordNumber) {
            position = (Number(this.evaluateExpression(stmt.recordNumber)) - 1);
        }

        const bytesRead = this.fs.readSync(handle.fd, buffer, 0, buffer.length, position);
        const s = new TextDecoder().decode(buffer.subarray(0, bytesRead));
        this.evaluateAssignmentToVariable(stmt.variable, s);
        handle.pos! += bytesRead;
    }

    private evaluateSeekStatement(stmt: SeekStatement) {
        const fileNum = Number(this.evaluateExpression(stmt.fileNumber));
        const handle = this.fileHandles.get(fileNum);
        if (!handle) this.throwVbaError(VbaErrorCode.BAD_FILE_NAME_OR_NUMBER, `Bad file name or number: #${fileNum}`);

        const pos = Number(this.evaluateExpression(stmt.position));
        // Node doesn't have seekSync on FD directly without lseek,
        // but we can track it in our handle if we use it for subsequent read/write.
        handle.pos = Math.max(0, pos - 1);
    }

    private evaluateResetStatement(_stmt: ResetStatement) {
        for (const [_num, handle] of this.fileHandles) {
            this.fs.closeSync(handle.fd);
        }
        this.fileHandles.clear();
    }

    private evaluateAttributeStatement(stmt: AttributeStatement) {
        if (stmt.name.toLowerCase() === 'vb_name') {
            const val = String(this.evaluateExpression(stmt.value)).replace(/"/g, '');
            // Attribute VB_Name defines the module name for procedures in this source
            this.currentSourceModule = val;
            this.executingModuleName = val;
        }
    }

    private evaluateDeclareStatement(stmt: DeclareStatement) {
        const name = stmt.name.toLowerCase();
        this.env.set(name, (..._args: any[]) => {
            this.onPrint(`[DECLARE STUB] Calling ${stmt.isSub ? 'Sub' : 'Function'} ${stmt.name} from "${stmt.libName}" (Alias: ${stmt.aliasName || 'N/A'})`);
            return 0; // Dummy return
        });
    }

    private externalObjectFactories: Map<string, () => any> = new Map();
    /** §5.6.10 Tier 5: COM 型ライブラリの名前空間 sentinel（Scripting, ADODB 等）。
     * Tier 1-4 で解決できなかった修飾 LHS 識別子にのみ使用する。globalEnv には登録しない。 */
    private typeLibraryNamespaces: Map<string, VbaNamespaceRef> = new Map();
    /**
     * `Dim x As New ClassName` で宣言された変数の追跡。キーは変数名(小文字)、
     * 値はクラス名。Set x = Nothing 後の再インスタンス化判定で使う。
     */
    private autoInstanceVars: Map<string, string> = new Map();

    /**
     * CreateObject(progId) で返されるオブジェクトのファクトリを登録する。
     * 既存の組み込みスタブ（Scripting.Dictionary 等）よりも優先して使われる。
     *
     * 加えて、factory() が返すオブジェクトに `__progId__`（推奨）または `__className__`
     * が含まれていれば、その値でも別名登録する。これにより VBA の「参照設定」相当の構文
     * （`Dim app As Word.Application` / `Set app = New Word.Application`）からも
     * 同じ factory が呼ばれる。
     *
     * 主にテスト用途でモックを差し込むために使用する。
     */
    /**
     * COM オブジェクトのファクトリを登録する。
     *
     * factory() が返すオブジェクトの `__progId__` を主キーとして登録し、
     * ドット後ろの短縮クラス名（"Dictionary" 等）も自動的に別名登録する。
     * これにより `CreateObject("Scripting.Dictionary")` と `New Dictionary` の両方が動く。
     *
     * @param factory VbaComObject を返すファクトリ関数
     * @param extraProgIds 同じ factory を追加で登録する ProgID（例: "Microsoft.XMLHTTP"）
     */
    public registerComObject(factory: () => VbaComObject, ...extraProgIds: string[]): void {
        const register = (key: string) => {
            const k = key.toLowerCase();
            if (!this.externalObjectFactories.has(k)) {
                this.externalObjectFactories.set(k, factory);
            }
            // "ProjectName.ClassName" 形式ならプロジェクト名を Tier 5 に登録。
            // globalEnv には登録しない（Tier 3 と Tier 5 を混在させない）。
            // 値コンテキストでの検出は evaluateExpression の Identifier ケースで行う。
            const dot = k.indexOf('.');
            if (dot > 0) {
                const projectName = k.slice(0, dot);
                if (!this.typeLibraryNamespaces.has(projectName)) {
                    this.typeLibraryNamespaces.set(projectName, new VbaNamespaceRef(k.slice(0, dot), 'project'));
                }
            }
        };

        try {
            const sample = factory();
            const progId = sample.__progId__;
            // 主キー登録
            this.externalObjectFactories.set(progId.toLowerCase(), factory);
            // "X.Y" → "Y" の短縮形も登録
            const dot = progId.lastIndexOf('.');
            if (dot >= 0) register(progId.slice(dot + 1));
            // VbaNamespaceRef 登録
            register(progId);
        } catch { /* sample 取得時のエラーは無視 */ }

        // 追加 ProgID
        for (const extra of extraProgIds) {
            register(extra);
        }
    }

    private createExternalObject(progId: string): any {
        const id = progId.toLowerCase();
        const factory = this.externalObjectFactories.get(id);
        if (factory) return factory();
        this.throwVbaError(VbaErrorCode.ACTIVEX_CANT_CREATE_OBJECT, `ActiveX component can't create object: '${progId}'`);
    }

    /**
     * 組み込みの外部オブジェクト（Scripting.Dictionary 等）をファクトリ形式で
     * 登録する。これにより以下の両方の呼び出しから同じファクトリが使われる:
     *   - CreateObject("Scripting.Dictionary")
     *   - Dim d As New Dictionary / Set d = New Dictionary（参照設定相当）
     *
     * テスト用に registerComObject で同じ factory を再登録すれば
     * 上書き（モック差し替え）も可能。
     */
    private registerBuiltinExternalObjects(): void {
        // --- Scripting.Dictionary ---
        this.registerComObject( () => {
            const dict = new Map<any, any>();
            return {
                __isVbaDict__: true,
                __progId__: 'Scripting.Dictionary',
                __map__: dict,
                add: (k: any, v: any) => {
                    if (dict.has(k)) this.throwVbaError(VbaErrorCode.KEY_ALREADY_EXISTS, 'This key is already associated with an element of this collection');
                    dict.set(k, v);
                },
                exists: (k: any) => dict.has(k) ? vbaTrue : vbaFalse,
                remove: (k: any) => dict.delete(k),
                removeall: () => dict.clear(),
                count: () => dict.size,
                keys: () => Array.from(dict.keys()),
                items: () => Array.from(dict.values()),
                item: (k: any, v?: any) => {
                    if (v !== undefined) {
                        dict.set(k, v);
                    }
                    return dict.get(k);
                }
            };
        });

        // --- Collection (§6.1.3.1) ---
        this.registerComObject( () => {
            const items: any[] = [];
            const keys: (string | undefined)[] = [];

            const resolveIndex = (index: any): number => {
                if (typeof index === 'string') {
                    const i = keys.findIndex(k => k !== undefined && k.toLowerCase() === index.toLowerCase());
                    if (i === -1) this.throwVbaError(VbaErrorCode.INVALID_PROCEDURE_CALL, `Collection key not found: '${index}'`);
                    return i;
                }
                const i = Number(index) - 1;
                if (i < 0 || i >= items.length) this.throwVbaError(VbaErrorCode.SUBSCRIPT_OUT_OF_RANGE, 'Subscript out of range');
                return i;
            };

            return {
                __isVbaCollection__: true,
                __progId__: 'Collection',
                count: () => items.length,
                add: (item: any, key?: any, before?: any, after?: any) => {
                    if (key !== undefined && key !== null && typeof key === 'string') {
                        if (keys.some(k => k !== undefined && k.toLowerCase() === key.toLowerCase())) {
                            this.throwVbaError(VbaErrorCode.KEY_ALREADY_EXISTS, 'This key is already associated with an element of this collection');
                        }
                    }
                    const keyStr = (key !== undefined && key !== null && typeof key === 'string') ? key : undefined;
                    if (before !== undefined && before !== null) {
                        const pos = resolveIndex(before);
                        items.splice(pos, 0, item);
                        keys.splice(pos, 0, keyStr);
                    } else if (after !== undefined && after !== null) {
                        const pos = resolveIndex(after) + 1;
                        items.splice(pos, 0, item);
                        keys.splice(pos, 0, keyStr);
                    } else {
                        items.push(item);
                        keys.push(keyStr);
                    }
                },
                item: (index: any) => {
                    const i = resolveIndex(index);
                    return items[i];
                },
                remove: (index: any) => {
                    const i = resolveIndex(index);
                    items.splice(i, 1);
                    keys.splice(i, 1);
                },
                // For For Each support
                [Symbol.iterator]: function* () { yield* items; },
            };
        });

        // --- Scripting.FileSystemObject ---
        this.registerComObject( () => ({
            __isVbaFso__: true,
            __progId__: 'Scripting.FileSystemObject',
            fileexists: (p: string) => {
                try {
                    const full = this.sandbox.toRealPath(p);
                    return (this.fs.existsSync(full) && this.fs.statSync(full).isFile()) ? vbaTrue : vbaFalse;
                } catch { return vbaFalse; }
            },
            folderexists: (p: string) => {
                try {
                    const full = this.sandbox.toRealPath(p);
                    return (this.fs.existsSync(full) && this.fs.statSync(full).isDirectory()) ? vbaTrue : vbaFalse;
                } catch { return vbaFalse; }
            },
            createtextfile: (p: string, overwrite: boolean = true) => {
                const full = this.sandbox.toRealPath(p);
                if (!overwrite && this.fs.existsSync(full)) this.throwVbaError(VbaErrorCode.FILE_ALREADY_EXISTS, "File already exists");
                const fd = this.fs.openSync(full, 'w');
                return {
                    write: (s: string) => this.fs.writeSync(fd, s),
                    writeline: (s: string) => this.fs.writeSync(fd, s + "\r\n"),
                    close: () => this.fs.closeSync(fd)
                };
            },
            opentextfile: (p: string, iomode: number = 1) => {
                const full = this.sandbox.toRealPath(p);
                const mode = iomode === 1 ? 'r' : (iomode === 2 ? 'w' : 'a');
                const fd = this.fs.openSync(full, mode);
                let pos = 0;
                return {
                    readall: () => {
                        const content = this.fs.readFileSync(full, 'utf8');
                        return content;
                    },
                    readline: () => {
                        const content = this.fs.readFileSync(full, 'utf8');
                        const lines = content.slice(pos).split(/\r?\n/);
                        if (lines.length > 0) {
                            const line = lines[0];
                            pos += line.length + (content[pos + line.length] === '\r' ? 2 : 1);
                            return line;
                        }
                        return "";
                    },
                    write: (s: string) => this.fs.writeSync(fd, s),
                    close: () => this.fs.closeSync(fd)
                };
            },
            createfolder: (p: string) => {
                const full = this.sandbox.toRealPath(p);
                this.fs.mkdirSync(full, { recursive: true });
                return { path: p };
            },
            deletefile: (p: string) => {
                const full = this.sandbox.toRealPath(p);
                this.fs.unlinkSync(full);
            },
            deletefolder: (p: string) => {
                const full = this.sandbox.toRealPath(p);
                this.fs.rmSync?.(full, { recursive: true, force: true });
            },
            getfile: (p: string) => {
                const full = this.sandbox.toRealPath(p);
                const stats = this.fs.statSync(full);
                return {
                    path: p,
                    size: stats.size,
                    datecreated: new VbaDate(toVbaDate(stats.birthtime || stats.mtime)),
                    datelastaccessed: new VbaDate(toVbaDate(stats.mtime)),
                    datelastmodified: new VbaDate(toVbaDate(stats.mtime)),
                    attributes: stats.mode || 0
                };
            },
            getfolder: (p: string) => ({ path: p }),
            // VBA は Windows パス前提。path.win32 で `\` をセパレータとして処理する。
            // GetBaseName は VBA 仕様で「拡張子を除いたファイル名」。
            getbasename: (p: string) => {
                const base = path.win32.basename(p);
                const ext = path.win32.extname(base);
                return ext ? base.slice(0, -ext.length) : base;
            },
            getextensionname: (p: string) => path.win32.extname(p).replace('.', ''),
            getparentfoldername: (p: string) => path.win32.dirname(p),
            getabsolutepathname: (p: string) => p
        }));

        // --- MSXML2.XMLHTTP / Microsoft.XMLHTTP ---
        const xmlhttpFactory = () => {
            let responseText = "";
            let status = 0;
            return {
                __progId__: 'MSXML2.XMLHTTP',
                open: (_method: string, _url: string, _async: boolean = true) => { /* Mock */ },
                send: (_body?: any) => { /* Mock */ },
                setrequestheader: (_h: string, _v: string) => { /* Mock */ },
                getresponsetext: () => responseText,
                responsetext: responseText,
                getstatus: () => status,
                status: status,
                readystate: 4
            };
        };
        this.registerComObject(xmlhttpFactory, 'Microsoft.XMLHTTP');
        

        // --- ADODB.Stream ---
        this.registerComObject( () => {
            let content = "";
            let streamPos = 0;
            return {
                __progId__: 'ADODB.Stream',
                open: () => { streamPos = 0; },
                close: () => { },
                write: (data: any) => { content += String(data); },
                writetext: (text: string) => { content += text; },
                read: (len: number) => { const r = content.slice(streamPos, streamPos + len); streamPos += len; return r; },
                readtext: () => { const r = content.slice(streamPos); streamPos = content.length; return r; },
                savetofile: (p: string, _mode: number = 1) => {
                    const full = this.sandbox.toRealPath(p);
                    this.fs.writeFileSync(full, content);
                },
                loadfromfile: (p: string) => {
                    const full = this.sandbox.toRealPath(p);
                    content = this.fs.readFileSync(full, 'utf8');
                    streamPos = 0;
                },
                type: 2,
                charset: 'utf-8',
                position: streamPos,
                size: content.length
            };
        });
    }

    // Execute a sequence of statements starting from startIndex, with error handling support.
    // isTopLevel=true (default): full handling for procedure bodies (GoTo, GoSub, Return, Resume).
    // isTopLevel=false: only handle On Error Resume Next; all control flow re-throws to outer scope.
    private executeStatements(body: Statement[], startIndex: number, isTopLevel: boolean = true) {
        let i = startIndex;
        while (i < body.length) {
            const stmt = body[i];
            try {
                this.evaluateStatement(stmt);
                i++;
            } catch (e: any) {
                if (e && (e.type === 'Exit' || e.type === 'Terminate' || e.type === 'ProcedureReturn')) {
                    throw e;
                }

                // In nested block contexts, propagate all procedure-level control flow
                if (!isTopLevel && e && (e.type === 'GoTo' || e.type === 'GoSub' || e.type === 'Return' || e.type === 'Resume')) {
                    throw e;
                }

                if (e && e.type === 'GoTo') {
                    const labelName = e.label.toLowerCase();
                    const labelIndex = body.findIndex(s =>
                        s.type === 'LabelStatement' &&
                        (s as any).label.toLowerCase() === labelName
                    );
                    if (labelIndex >= 0) {
                        i = labelIndex;
                        continue;
                    }
                    this.throwVbaError(VbaErrorCode.SUB_OR_FUNCTION_NOT_DEFINED, `Sub or Function not defined: label '${e.label}'`);
                }

                if (e && e.type === 'GoSub') {
                    const labelName = e.label.toLowerCase();
                    const labelIndex = body.findIndex(s =>
                        s.type === 'LabelStatement' &&
                        (s as any).label.toLowerCase() === labelName
                    );
                    if (labelIndex >= 0) {
                        this.gosubStack.push(i);
                        i = labelIndex;
                        continue;
                    }
                    throw e;
                }

                if (e && e.type === 'Return') {
                    if (this.gosubStack.length === 0) {
                        this.throwVbaError(VbaErrorCode.RETURN_WITHOUT_GOSUB, 'Return without GoSub');
                    }
                    i = this.gosubStack.pop()! + 1;
                    continue;
                }

                if (e && e.type === 'Resume') {
                    this.isInErrorHandler = false;
                    if (e.mode === 'Current') {
                        i = this.lastErrorIndex !== null ? this.lastErrorIndex : i;
                    } else if (e.mode === 'Next') {
                        i = this.lastErrorIndex !== null ? this.lastErrorIndex + 1 : i + 1;
                    } else if (e.mode === 'Label') {
                        const labelName = e.label.toLowerCase();
                        const labelIndex = body.findIndex(s =>
                            s.type === 'LabelStatement' &&
                            (s as any).label.toLowerCase() === labelName
                        );
                        if (labelIndex >= 0) {
                            i = labelIndex;
                        } else {
                            this.throwVbaError(VbaErrorCode.SUB_OR_FUNCTION_NOT_DEFINED, `Sub or Function not defined: label '${e.label}'`);
                        }
                    }
                    this.errObj.clear();
                    continue;
                }

                // Handle VbaError or standard JS Error
                if (!this.isInErrorHandler) {
                    const errorNumber = e.type === 'VbaError' ? e.number : 1000;
                    // vbaBareMessage があれば "Run-time error 'N': ... (line X)" の枠組みを
                    // 含まない生のメッセージを使う（throwVbaError 経由のエラー）。
                    // Err.Raise（ErrObject.raise）は元々 message に生のメッセージを
                    // 直接渡しているため vbaBareMessage は持たない。
                    const errorMessage = e.vbaBareMessage ?? e.message ?? String(e);

                    this.errObj.number = errorNumber;
                    this.errObj.description = errorMessage;

                    if (this.errorHandlingMode === 'ResumeNext') {
                        this.lastErrorIndex = i;
                        i++;
                        continue;
                    }

                    if (this.errorHandlingMode === 'GoTo' && this.errorHandlerLabel) {
                        const labelName = this.errorHandlerLabel.toLowerCase();
                        const labelIndex = body.findIndex(s =>
                            s.type === 'LabelStatement' &&
                            (s as any).label.toLowerCase() === labelName
                        );
                        if (labelIndex >= 0) {
                            this.lastErrorIndex = i;
                            this.isInErrorHandler = true;
                            i = labelIndex;
                            continue;
                        }
                    }
                }

                // If we reach here, either we are already in an error handler,
                // or there's no handler configured. Bubble up.
                // Do NOT reset isInErrorHandler here: if we're inside a handler and a nested
                // block (If/For/While) re-throws, the outer executeStatements must still see
                // isInErrorHandler=true so it doesn't re-enter the same handler (infinite loop).
                throw e;
            }
        }
    }

    private evaluateEraseStatement(stmt: EraseStatement) {
        const varName = stmt.name.name;
        const arr = this.env.get(varName);
        if (Array.isArray(arr)) {
            if ((arr as any).vbaFixed) {
                // Fixed array: re-initialize elements
                const defaultValue = (arr as any).__vbaDefaultValue__ ?? 0;
                this.reinitializeArray(arr, defaultValue);
            } else {
                // Dynamic array: de-allocate
                const newArr: any[] = [];
                (newArr as any).vbaBase = (arr as any).vbaBase ?? this.arrayBase;
                (newArr as any).vbaFixed = false;
                this.env.set(varName, newArr);
            }
        }
    }

    private reinitializeArray(arr: any[], defaultValue: any) {
        for (let i = 0; i < arr.length; i++) {
            if (Array.isArray(arr[i]) && !(arr[i] as any).__vbaClass__) {
                this.reinitializeArray(arr[i], defaultValue);
            } else {
                arr[i] = defaultValue;
            }
        }
    }

    private fillArrayWithUdtInstances(arr: any[], typeName: string) {
        const dims = (arr as any).__vbaDimensions__;
        if (!dims) return;
        const fillDim = (a: any[], dimIdx: number) => {
            const { lower, upper } = dims[dimIdx];
            if (dimIdx < dims.length - 1) {
                for (let i = lower; i <= upper; i++) {
                    if (Array.isArray(a[i])) fillDim(a[i], dimIdx + 1);
                }
            } else {
                for (let i = lower; i <= upper; i++) {
                    a[i] = this.instantiateType(typeName);
                }
            }
        };
        fillDim(arr, 0);
    }

    /** Dim 文の配列境界式が定数式であることを検証する。変数参照を含む場合はコンパイルエラー。 */
    private validateConstantExpr(expr: Expression): void {
        switch (expr.type) {
            case 'NumberLiteral':
            case 'StringLiteral':
            case 'DateLiteral':
                return;
            case 'Identifier': {
                const name = (expr as Identifier).name;
                if (!this.env.isConstant(name.toLowerCase())) {
                    this.throwCompileError(VbaErrorCode.CONSTANT_EXPRESSION_REQUIRED, 'Constant expression required', expr.loc?.start.line);
                }
                return;
            }
            case 'UnaryExpression':
                this.validateConstantExpr((expr as UnaryExpression).argument);
                return;
            case 'BinaryExpression':
                this.validateConstantExpr((expr as BinaryExpression).left);
                this.validateConstantExpr((expr as BinaryExpression).right);
                return;
            case 'ParenthesizedExpression':
                this.validateConstantExpr((expr as ParenthesizedExpression).expression);
                return;
            default:
                this.throwCompileError(VbaErrorCode.CONSTANT_EXPRESSION_REQUIRED, 'Constant expression required', expr.loc?.start.line);
        }
    }

    private createMultiDimArray(bounds: ArrayBound[], initialValue: any): any[] {
        const dimensions: { lower: number, upper: number }[] = [];

        for (const bound of bounds) {
            const upper = Number(this.evaluateExpression(bound.upper));
            const lower = bound.lower ? Number(this.evaluateExpression(bound.lower)) : this.arrayBase;
            dimensions.push({ lower, upper });
        }

        // VBA index と JS index を一致させるため、length = upper + 1 の配列を作り、
        // [0]〜[lower-1] は undefined、[lower]〜[upper] を initialValue で埋める。
        const buildArray = (dimIdx: number): any[] => {
            const { lower, upper } = dimensions[dimIdx];
            if (upper < lower - 1) this.throwVbaError(VbaErrorCode.SUBSCRIPT_OUT_OF_RANGE, 'Subscript out of range');

            const totalSize = upper + 1;
            const arr = new Array(totalSize);
            if (dimIdx < dimensions.length - 1) {
                for (let i = lower; i <= upper; i++) {
                    arr[i] = buildArray(dimIdx + 1);
                }
            } else {
                for (let i = lower; i <= upper; i++) {
                    arr[i] = initialValue;
                }
            }
            return arr;
        };

        const result = buildArray(0);
        (result as any).__vbaDimensions__ = dimensions;
        (result as any).vbaBase = dimensions[0].lower; // 既存コードとの互換用
        (result as any).__vbaDefaultValue__ = initialValue;
        return result;
    }

    private copyPreservedData(oldArr: any, newArr: any, dimensions: { lower: number, upper: number }[], dimIdx: number = 0) {
        if (!Array.isArray(oldArr) || !Array.isArray(newArr)) return;

        const copyLen = Math.min(oldArr.length, newArr.length);

        if (dimIdx < dimensions.length - 1) {
            for (let i = 0; i < copyLen; i++) {
                this.copyPreservedData(oldArr[i], newArr[i], dimensions, dimIdx + 1);
            }
        } else {
            for (let i = 0; i < copyLen; i++) {
                newArr[i] = oldArr[i];
            }
        }
    }

    private evaluateReDimStatement(stmt: ReDimStatement) {
        const varName = stmt.name.name;
        const oldArr = this.env.get(varName);

        // UDT 配列の場合、Dim 時に保存した要素型名を引き継ぐ
        const elementTypeName: string | undefined =
            (Array.isArray(oldArr) ? (oldArr as any).__vbaElementTypeName__ : undefined) ??
            (stmt.objectType && this.env.getType(stmt.objectType) ? stmt.objectType : undefined);

        let defaultValue: any = 0;
        if (stmt.objectType) {
            const t = stmt.objectType.toLowerCase();
            if (t === 'string') defaultValue = '';
            else if (t === 'boolean') defaultValue = 0;
        } else if (Array.isArray(oldArr)) {
            defaultValue = (oldArr as any).__vbaDefaultValue__ ?? 0;
        }

        if (stmt.bounds.length > 0) {
            // Validate ReDim Preserve constraints for multi-dimensional arrays
            if (stmt.isPreserve && Array.isArray(oldArr) && (oldArr as any).__vbaDimensions__) {
                const oldDims = (oldArr as any).__vbaDimensions__ as { lower: number, upper: number }[];

                // Evaluate new bounds
                const newDims = stmt.bounds.map(bound => {
                    const lower = bound.lower ? this.evaluateExpression(bound.lower) : 0;
                    const upper = this.evaluateExpression(bound.upper);
                    return { lower, upper };
                });

                // Check constraint 1: Number of dimensions cannot change
                if (newDims.length !== oldDims.length) {
                    this.throwVbaError(VbaErrorCode.SUBSCRIPT_OUT_OF_RANGE, "Subscript out of range");
                }

                // Check constraint 2: Lower bound of any dimension cannot change
                for (let i = 0; i < newDims.length; i++) {
                    if (newDims[i].lower !== oldDims[i].lower) {
                        this.throwVbaError(VbaErrorCode.SUBSCRIPT_OUT_OF_RANGE, "Subscript out of range");
                    }
                }

                // Check constraint 3: Upper bound of any dimension other than the last cannot change
                for (let i = 0; i < newDims.length - 1; i++) {
                    if (newDims[i].upper !== oldDims[i].upper) {
                        this.throwVbaError(VbaErrorCode.SUBSCRIPT_OUT_OF_RANGE, "Subscript out of range");
                    }
                }
            }

            const arr = this.createMultiDimArray(stmt.bounds, defaultValue);
            (arr as any).vbaFixed = false;

            if (elementTypeName) {
                // UDT 配列: 要素ごとに独立したインスタンスを生成する
                // createMultiDimArray は全要素を同一参照で埋めるため、ここで上書きする
                (arr as any).__vbaElementTypeName__ = elementTypeName;
                if (!stmt.isPreserve) {
                    this.fillArrayWithUDT(arr, (arr as any).__vbaDimensions__, 0, elementTypeName);
                }
            }

            if (stmt.isPreserve && Array.isArray(oldArr)) {
                this.copyPreservedData(oldArr, arr, (arr as any).__vbaDimensions__);
            }

            this.env.set(varName, arr);
        }
    }

    private fillArrayWithUDT(arr: any[], dimensions: { lower: number, upper: number }[], dimIdx: number, typeName: string) {
        const { lower, upper } = dimensions[dimIdx];
        if (dimIdx < dimensions.length - 1) {
            for (let i = lower; i <= upper; i++) {
                this.fillArrayWithUDT(arr[i], dimensions, dimIdx + 1, typeName);
            }
        } else {
            for (let i = lower; i <= upper; i++) {
                arr[i] = this.instantiateType(typeName);
            }
        }
    }

    private evaluateExitStatement(stmt: ExitStatement) {
        throw { type: 'Exit', target: stmt.exitType };
    }

    private toDisplayString(val: any): string { return vbaToDisplayString(val); }

    private evaluateExpression(expr: Expression): any {
        switch (expr.type) {
            case 'MissingArgument':
                return vbaEmpty;
            case 'NumberLiteral': {
                const lit = expr as NumberLiteral;
                return this.applyLiteralTypeSuffix(lit.value, lit.typeSuffix);
            }
            case 'StringLiteral':
                return (expr as StringLiteral).value;
            case 'AddressOfExpression': {
                const ao = expr as AddressOfExpression;
                return ao.moduleName ? `${ao.moduleName}.${ao.procedureName.name}` : ao.procedureName.name;
            }
            case 'DateLiteral':
                return this.evaluateDateLiteral(expr as DateLiteral);
            case 'Identifier':
                const idName = (expr as Identifier).name;
                // §5.6.10 Tier 6: env に明示宣言がなく、手続きにも該当がない場合は
                // defaultBindingObject（MockApplication 等）のプロパティとして解決する。
                // env.get() より先に確認しないと暗黙初期化（vbaEmpty）で上書きされてしまう。
                if (!this.inConstEval && this.defaultBindingObject
                        && !this.env.hasVariable(idName) && !this.env.getProcedure(idName)) {
                    const tier6Key = this.resolveObjectMemberKey(this.defaultBindingObject, idName.toLowerCase());
                    if (tier6Key !== undefined) {
                        const tier6Val = this.defaultBindingObject[tier6Key];
                        if (typeof tier6Val === 'function') {
                            return (tier6Val as (...a: any[]) => any).apply(this.defaultBindingObject, []);
                        }
                        return tier6Val;
                    }
                }
                // §5.6.10 Tier 5: COM 型ライブラリ名前空間（Scripting 等）を値として使う場合はエラー。
                // env.get より先にチェックしないと暗黙初期化（vbaEmpty）で上書きされてしまう。
                // Tier 1-4 に同名の宣言がある場合は Tier 5 より優先（env.hasVariable が true）。
                if (!this.inConstEval && !this.env.hasVariable(idName)) {
                    const tier5 = this.typeLibraryNamespaces.get(idName.toLowerCase());
                    if (tier5) {
                        this.throwVbaError(VbaErrorCode.INVALID_PROCEDURE_CALL,
                            `Compile error: Expected variable or procedure, not project ('${tier5.name}')`);
                    }
                }
                // VbaNamespaceRef（プロジェクト名・モジュール名）を値として使う場合はコンパイルエラー。
                // Dim VBA As Long のように明示宣言済みの場合は env.set で上書きされているため
                // VbaNamespaceRef ではなく実際の変数値が返る。
                const v = this.inConstEval ? this.resolveConstIdent(idName) : this.env.get(idName);
                if (v instanceof VbaNamespaceRef) {
                    const label = v.kind === 'project' ? 'project' : 'module';
                    this.throwVbaError(VbaErrorCode.INVALID_PROCEDURE_CALL,
                        `Compile error: Expected variable or procedure, not ${label} ('${v.name}')`);
                }
                if (typeof v === 'function' && this.isAutoCallable(v)) {
                    return v();
                }
                const p = this.env.getProcedure(idName);
                if (p) {
                    // Inside function F, bare `F` refers to the return value, not a recursive call.
                    const isCurrentProcReturn = this.currentProcedureName &&
                        idName.toLowerCase() === this.currentProcedureName.toLowerCase();
                    if (!isCurrentProcReturn) {
                        // Only auto-call if it's a Function/Property and has 0 required arguments
                        const requiredCount = p.parameters.filter(param => !param.isOptional && !param.isParamArray).length;
                        if (requiredCount === 0) {
                            return this.callProcedure(idName, []);
                        }
                    }
                }
                return v;
            case 'CallExpression':
                return this.evaluateCallExpression(expr as CallExpression);
            case 'MemberExpression':
                return this.evaluateMemberExpression(expr as MemberExpression);
            case 'DictionaryAccessExpression':
                return this.evaluateDictionaryAccessExpression(expr as DictionaryAccessExpression);
            case 'TypeOfIsExpression':
                return this.evaluateTypeOfIsExpression(expr as TypeOfIsExpression);
            case 'UnaryExpression':
                return this.evaluateUnaryExpression(expr as UnaryExpression);
            case 'BinaryExpression':
                return this.evaluateBinaryExpression(expr as BinaryExpression);
            case 'ImplicitWithObjectExpression':
                return this.evaluateImplicitWithObjectExpression(expr as ImplicitWithObjectExpression);
            case 'ParenthesizedExpression':
                return this.evaluateExpression((expr as ParenthesizedExpression).expression);
            case 'NewExpression':
                return this.instantiateClass((expr as NewExpression).className);
            case 'NamedArgument':
                // Named arguments should only appear as call arguments, but evaluate the value if encountered
                return this.evaluateExpression((expr as NamedArgument).value);
            default:
                throw new Error(`Execution error: Unknown expression type ${expr.type}`);
        }
    }

    private evaluateDateLiteral(expr: DateLiteral): any {
        const parsed = this.parseDateLiteral(expr.value);
        if (!parsed) {
            this.throwVbaError(VbaErrorCode.TYPE_MISMATCH, `Type mismatch: invalid date literal #${expr.value}#`);
        }
        const d = new Date(parsed.year, parsed.month - 1, parsed.day, parsed.hour, parsed.minute, parsed.second);
        if (isNaN(d.getTime())) {
            this.throwVbaError(VbaErrorCode.TYPE_MISMATCH, `Type mismatch: invalid date literal #${expr.value}#`);
        }
        return new VbaDate(toVbaDate(d));
    }

    private parseDateLiteral(dateStr: string): { year: number, month: number, day: number, hour: number, minute: number, second: number } | null {
        // Month names mapping
        const monthMap: Record<string, number> = {
            'january': 1, 'february': 2, 'march': 3, 'april': 4, 'may': 5, 'june': 6,
            'july': 7, 'august': 8, 'september': 9, 'october': 10, 'november': 11, 'december': 12,
            'jan': 1, 'feb': 2, 'mar': 3, 'apr': 4, 'jun': 6, 'jul': 7,
            'aug': 8, 'sep': 9, 'oct': 10, 'nov': 11, 'dec': 12
        };

        // Parse date and time components
        const timeMatch = dateStr.match(/(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?\s*([ap]\.?m\.?)?/i);
        let dateComponent = dateStr;
        let hour = 0, minute = 0, second = 0;

        if (timeMatch) {
            hour = parseInt(timeMatch[1], 10);
            minute = parseInt(timeMatch[2], 10);
            second = timeMatch[3] ? parseInt(timeMatch[3], 10) : 0;
            const ampm = timeMatch[4];
            if (ampm && /^p/i.test(ampm) && hour < 12) hour += 12;
            if (ampm && /^a/i.test(ampm) && hour === 12) hour = 0;
            dateComponent = dateStr.replace(timeMatch[0], '').trim();
        }

        // Extract date components (up to 3 parts separated by /, -, or whitespace+comma)
        const parts = dateComponent.split(/[\/\-,\s]+/).filter(p => p.length > 0);
        if (parts.length === 0 || parts.length > 3) return null;

        // Parse each part (number or month name)
        const parsed: (number | string)[] = [];
        for (const part of parts) {
            if (/^\d+$/.test(part)) {
                parsed.push(parseInt(part, 10));
            } else {
                const monthNum = monthMap[part.toLowerCase()];
                if (monthNum) {
                    parsed.push(monthNum);
                } else {
                    return null;
                }
            }
        }

        let month: number, day: number, year: number;
        const currentYear = new Date().getFullYear();

        if (parsed.length === 2) {
            // Two components: try mm/dd or dd/mm
            const [L, M] = parsed;
            if (typeof L === 'number' && typeof M === 'number') {
                // Rule: If L is valid month and M is valid day, L=month, M=day, year=current
                if (this.isValidMonth(L) && this.isValidDay(L, M, currentYear)) {
                    month = L; day = M; year = currentYear;
                } else if (this.isValidMonth(M) && this.isValidDay(M, L, currentYear)) {
                    // Otherwise if M is valid month and L is valid day, M=month, L=day
                    month = M; day = L; year = currentYear;
                } else {
                    return null;
                }
            } else {
                // One part is a month name
                let monthNum: number, dayNum: number;
                if (typeof L === 'string') {
                    monthNum = L as any;
                    dayNum = M as number;
                } else {
                    monthNum = M as any;
                    dayNum = L as number;
                }
                if (this.isValidDay(monthNum, dayNum, currentYear)) {
                    month = monthNum;
                    day = dayNum;
                    year = currentYear;
                } else {
                    month = monthNum;
                    day = 1;
                    year = dayNum;
                }
            }
        } else if (parsed.length === 3) {
            // Three components: try mm/dd/yyyy, dd/mm/yyyy, etc.
            const [L, M, R] = parsed;
            if (typeof L === 'number' && typeof M === 'number' && typeof R === 'number') {
                // Rule: If L is valid month and M is valid day in that month with year R, use that
                if (this.isValidMonth(L) && this.isValidDay(L, M, R)) {
                    month = L; day = M; year = R;
                } else if (this.isValidMonth(M) && this.isValidDay(M, R, L)) {
                    // Otherwise if M is valid month and R is valid day in that month with year L
                    month = M; day = R; year = L;
                } else if (this.isValidMonth(M) && this.isValidDay(M, L, R)) {
                    // Otherwise if M is valid month and L is valid day in that month with year R
                    month = M; day = L; year = R;
                } else {
                    return null;
                }
            } else {
                // One part is a month name - find which one and apply rules
                let monthNum: number, n1: number, n2: number;
                if (typeof L === 'string') {
                    monthNum = L as any; n1 = M as number; n2 = R as number;
                } else if (typeof M === 'string') {
                    monthNum = M as any; n1 = L as number; n2 = R as number;
                } else {
                    monthNum = R as any; n1 = L as number; n2 = M as number;
                }
                // Try both possibilities
                if (this.isValidDay(monthNum, n1, n2)) {
                    month = monthNum; day = n1; year = n2;
                } else if (this.isValidDay(monthNum, n2, n1)) {
                    month = monthNum; day = n2; year = n1;
                } else {
                    return null;
                }
            }
        } else {
            return null;
        }

        // Handle two-digit years (sliding window: 00-99 -> 2000-2099 for now)
        if (year < 100) {
            year += 2000;
        }

        return { year, month, day, hour, minute, second };
    }

    private isValidMonth(m: number): boolean {
        return m >= 1 && m <= 12;
    }

    private isValidDay(month: number, day: number, year: number): boolean {
        if (day < 1) return false;
        const daysInMonth = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
        // Leap year check
        if (month === 2 && ((year % 4 === 0 && year % 100 !== 0) || year % 400 === 0)) {
            return day <= 29;
        }
        return day <= daysInMonth[month - 1];
    }

    /**
     * Infer the VBA type name for a number literal value.
     * - Integer if it's a whole number in -32768..32767
     * - Long if it's a whole number in Long range
     * - Double otherwise (or if it has a fractional part)
     */
    private applyLiteralTypeSuffix(val: number, suffix: NumberLiteral['typeSuffix']): any {
        switch (suffix) {
            case '%': {
                const n = this.vbaRound(val, 0);
                if (n < -32768 || n > 32767) this.throwVbaError(VbaErrorCode.OVERFLOW, 'Overflow');
                return n;
            }
            case '&': {
                const n = this.vbaRound(val, 0);
                if (n < -2147483648 || n > 2147483647) this.throwVbaError(VbaErrorCode.OVERFLOW, 'Overflow');
                return n;
            }
            case '!':
                // TypeName → 'Single' は AST の typeSuffix 経由で返る。
                // Math.fround は表示ラッパー（VbaSingle）が未実装のため適用しない。
                return val;
            case '#': return val;
            case '@': {
                const n = this.vbaRound(val, 4);
                if (n < -922337203685477.5808 || n > 922337203685477.5807) this.throwVbaError(VbaErrorCode.OVERFLOW, 'Overflow');
                return n;
            }
            case '^':
                // TypeName → 'LongLong' は AST の typeSuffix 経由で返る。
                // BigInt 変換は VbaSingle 同様にラッパー未実装のため行わない。
                return val;
            default: return val;
        }
    }

    private inferLiteralTypeName(lit: NumberLiteral): string {
        if (lit.typeSuffix) {
            const map: Record<string, string> = {
                '%': 'Integer', '&': 'Long', '!': 'Single',
                '#': 'Double', '@': 'Currency', '^': 'LongLong',
            };
            return map[lit.typeSuffix];
        }
        if (lit.isFloat) return 'Double';
        const val = lit.value;
        if (Number.isInteger(val)) {
            if (val >= -32768 && val <= 32767) return 'Integer';
            if (val >= -2147483648 && val <= 2147483647) return 'Long';
        }
        return 'Double';
    }

    private inferLiteralVarType(lit: NumberLiteral): number {
        if (lit.typeSuffix) {
            const map: Record<string, number> = {
                '%': 2, '&': 3, '!': 4, '#': 5, '@': 6, '^': 20,
            };
            return map[lit.typeSuffix];
        }
        if (lit.isFloat) return 5; // vbDouble
        const val = lit.value;
        if (Number.isInteger(val)) {
            if (val >= -32768 && val <= 32767) return 2; // vbInteger
            if (val >= -2147483648 && val <= 2147483647) return 3; // vbLong
        }
        return 5; // vbDouble
    }

    /**
     * Special evaluator for TypeName() and VarType() that inspects the AST
     * to resolve variable type metadata from the Environment.
     */
    // 組み込み関数の戻り型マップ（固定戻り型のもののみ）
    public static readonly BUILTIN_RETURN_TYPES: Record<string, VbaVarType> = {
        // 型変換関数
        'cbool': 'Boolean', 'cbyte': 'Byte', 'ccur': 'Currency',
        'cdate': 'Date', 'cdbl': 'Double', 'cint': 'Integer',
        'clng': 'Long', 'clnglng': 'LongLong', 'clngptr': 'LongPtr',
        'csng': 'Single', 'cstr': 'String',
        // Long を返す関数
        'ubound': 'Long', 'lbound': 'Long',
        'len': 'Long', 'lenb': 'Long',
        'instr': 'Long', 'instrb': 'Long', 'instrrev': 'Long',
        // Integer を返す関数
        'asc': 'Integer', 'ascb': 'Integer', 'ascw': 'Integer',
        'vartype': 'Integer',
        // Double を返す関数
        'sqr': 'Double', 'sin': 'Double', 'cos': 'Double',
        'tan': 'Double', 'atn': 'Double', 'exp': 'Double', 'log': 'Double',
        'val': 'Double',
        // Single を返す関数
        'rnd': 'Single', 'timer': 'Single',
        // Boolean を返す関数
        'isarray': 'Boolean', 'isdate': 'Boolean', 'isempty': 'Boolean',
        'iserror': 'Boolean', 'ismissing': 'Boolean', 'isnull': 'Boolean',
        'isnumeric': 'Boolean', 'isobject': 'Boolean',
        // Date を返す関数
        'now': 'Date', 'date': 'Date', 'time': 'Date',
        'dateserial': 'Date', 'timeserial': 'Date',
        'dateadd': 'Date',
        // String を返す関数
        'typename': 'String',
        'chr': 'String', 'chr$': 'String', 'chrb': 'String', 'chrw': 'String',
        'left': 'String', 'left$': 'String', 'right': 'String', 'right$': 'String',
        'mid': 'String', 'mid$': 'String',
        'ltrim': 'String', 'ltrim$': 'String', 'rtrim': 'String', 'rtrim$': 'String',
        'trim': 'String', 'trim$': 'String',
        'ucase': 'String', 'ucase$': 'String', 'lcase': 'String', 'lcase$': 'String',
        'space': 'String', 'space$': 'String',
        'string': 'String', 'string$': 'String',
        'hex': 'String', 'hex$': 'String', 'oct': 'String', 'oct$': 'String',
        'format': 'String', 'format$': 'String',
        'str': 'String', 'str$': 'String',
    };

    private evaluateTypeIntrinsic(funcName: 'typename' | 'vartype', argExpr: Expression): any {
        // VarType/TypeName のコードマップ（共通）
        const vtMap: Record<string, number> = {
            'Byte': 17, 'Integer': 2, 'Long': 3,
            'Single': 4, 'Double': 5, 'Currency': 6,
            'LongLong': 20, 'LongPtr': 20,
            'String': 8, 'Boolean': 11, 'Date': 7,
        };

        // 宣言型が確定しているケースを AST レベルで解決する（値の評価前）
        const declaredType = this.resolveDeclaredReturnType(argExpr);
        if (declaredType && declaredType !== 'Variant' && declaredType !== 'Object') {
            return funcName === 'typename' ? declaredType : (vtMap[declaredType] ?? 12);
        }

        // 値を評価して型を判定する
        const val = this.resolveAutoInstance(argExpr, this.evaluateExpression(argExpr));

        if (funcName === 'typename') {
            if (typeof val === 'number') {
                if (argExpr.type === 'NumberLiteral') return this.inferLiteralTypeName(argExpr as NumberLiteral);
                return 'Double'; // 型情報なしの式は Double
            }
            return this.env.get('typename')(val);
        } else {
            if (typeof val === 'number') {
                if (argExpr.type === 'NumberLiteral') return this.inferLiteralVarType(argExpr as NumberLiteral);
                return 5; // vbDouble
            }
            return this.env.get('vartype')(val);
        }
    }

    /** AST から式の宣言戻り型を解決する（変数・関数・組み込み変換関数） */
    private resolveDeclaredReturnType(expr: Expression): VbaVarType | undefined {
        if (expr.type === 'Identifier') {
            const typeInfo = this.env.getVariableType((expr as Identifier).name);
            return typeInfo?.vbaType;
        }
        if (expr.type === 'CallExpression') {
            const ce = expr as CallExpression;
            if (ce.callee.type === 'Identifier') {
                const nameLower = (ce.callee as Identifier).name.toLowerCase();
                // 組み込み変換関数
                const builtin = Evaluator.BUILTIN_RETURN_TYPES[nameLower];
                if (builtin) return builtin;
                // ユーザー定義関数の宣言型
                const proc = this.env.getProcedure(nameLower);
                if (proc?.returnType) {
                    const typeMap: Record<string, VbaVarType> = {
                        'byte': 'Byte', 'integer': 'Integer', 'long': 'Long',
                        'single': 'Single', 'double': 'Double', 'currency': 'Currency',
                        'longlong': 'LongLong', 'longptr': 'LongPtr',
                        'string': 'String', 'boolean': 'Boolean', 'date': 'Date',
                    };
                    return typeMap[proc.returnType.toLowerCase()];
                }
            }
        }
        return undefined;
    }

    private evaluateCallExpression(expr: CallExpression): any {
        if (expr.callee.type === 'Identifier') {
            const name = (expr.callee as Identifier).name;

            // Special handling: TypeName() and VarType() need AST-level access to check variable types
            const nameLower = name.toLowerCase();
            if ((nameLower === 'typename' || nameLower === 'vartype') && expr.args.length === 1) {
                return this.evaluateTypeIntrinsic(nameLower, expr.args[0]);
            }

            // B-2: When inside a class method (Me is in env), the class's own procedures
            // take priority over global procedures — matching VBA's name resolution order
            // (class scope → global scope).
            const me = this.env.getConst('me');
            if (me && me.__vbaClass__ && me.__classDef__) {
                const classProc = (me.__classDef__ as ClassDeclaration).procedures.find(
                    p => p.name.name.toLowerCase() === nameLower
                );
                if (classProc) {
                    this.checkNoGapOnRequiredParam(classProc.parameters, expr.args);
                    const argsVals = expr.args.map(a => this.resolveAutoInstance(a, this.evaluateExpression(a)));
                    this.vbaCallStack.push({ name: classProc.name.name, moduleName: me.__className__ ?? '', line: this.currentLine });
                    try {
                        return this.callClassMethod(me, classProc, argsVals);
                    } finally {
                        this.vbaCallStack.pop();
                    }
                }
            }

            const proc = this.env.getProcedure(name);

            if (proc) {
                // Cross-module Private access check
                if (
                    proc.scope === 'private' &&
                    proc.moduleName !== undefined &&
                    proc.moduleName !== '' &&
                    proc.moduleName !== this.executingModuleName
                ) {
                    this.throwVbaError(VbaErrorCode.INVALID_PROCEDURE_CALL,
                        `Cannot call Private procedure '${proc.name.name}' ` +
                        `from module '${this.executingModuleName || '(top-level)'}' ` +
                        `(defined in '${proc.moduleName}')`
                    );
                }

                this.vbaCallStack.push({ name: proc.name.name, moduleName: proc.moduleName ?? '', line: this.currentLine });
                try {

                this.precheckProc(proc);

                // Procedure call (Function/Sub): Tier 1, enclosing = Tier 2 (moduleEnv)
                const procModKey = (proc.moduleName ?? '').toLowerCase();
                const procParentEnv = this.moduleEnvs.get(procModKey) ?? this.env;
                const localEnv = new Environment(procParentEnv);

                // Map arguments to parameters
                const byRefArgs: { paramName: string, originalExpr: Expression }[] = [];
                let paramArrayParamName: string | null = null;
                let paramArrayByRefExprs: Expression[] = [];
                const namedArgs = new Map<string, any>();
                const namedArgExpressions = new Map<string, Expression>();
                const positionalArgs: any[] = [];
                const positionalArgExpressions: Expression[] = [];

                for (const argExpr of expr.args) {
                    if (argExpr.type === 'NamedArgument') {
                        const namedArg = argExpr as NamedArgument;
                        namedArgs.set(namedArg.name.toLowerCase(), this.evaluateExpression(namedArg.value));
                        namedArgExpressions.set(namedArg.name.toLowerCase(), namedArg.value);
                    } else {
                        positionalArgs.push(this.evaluateExpression(argExpr));
                        positionalArgExpressions.push(argExpr);
                    }
                }

                // Validate argument count
                {
                    const argBinderParams: ArgBinderParam[] = proc.parameters.map(p => ({
                        isOptional: !!p.isOptional || p.defaultValue != null,
                        isParamArray: !!p.isParamArray,
                    }));
                    this.checkArgCountGeneric(argBinderParams, positionalArgs.length + namedArgs.size);
                }
                this.checkNoGapOnRequiredParam(proc.parameters, positionalArgExpressions);

                for (let i = 0; i < proc.parameters.length; i++) {
                    const param = proc.parameters[i];
                    const paramNameLower = param.name.toLowerCase();

                    if (param.isParamArray) {
                        const remainingArgs = positionalArgs.slice(i);
                        (remainingArgs as any).vbaBase = 0;
                        localEnv.set(param.name, remainingArgs);
                        // Track for ByRef writeback (spec §5.3.1.5: param array elements behave as ByRef)
                        paramArrayParamName = param.name;
                        paramArrayByRefExprs = positionalArgExpressions.slice(i);
                        break;
                    }

                    let argVal: any;
                    const isMissingSlot = i < positionalArgExpressions.length &&
                        positionalArgExpressions[i].type === 'MissingArgument';
                    if (namedArgs.has(paramNameLower)) {
                        argVal = namedArgs.get(paramNameLower);
                    } else if (i < positionalArgs.length && !isMissingSlot) {
                        argVal = positionalArgs[i];
                    } else if (param.defaultValue) {
                        argVal = this.evaluateExpression(param.defaultValue);
                    } else {
                        // checkNoGapOnRequiredParam済みのため、未指定または省略スロットで
                        // ここに来る時点で param は必ず Optional（defaultValue なし）。
                        argVal = vbaMissing;
                    }
                    // Register parameter type metadata (but not for array parameters)
                    if (param.paramType && !param.isArray) {
                        const typeMap: Record<string, VbaVarType> = {
                            'byte': 'Byte', 'integer': 'Integer', 'long': 'Long',
                            'single': 'Single', 'double': 'Double', 'currency': 'Currency',
                            'string': 'String', 'boolean': 'Boolean', 'date': 'Date',
                        };
                        const mapped = typeMap[param.paramType.toLowerCase()];
                        if (mapped) {
                            localEnv.setVariableType(param.name, { vbaType: mapped });
                        }
                    }
                    localEnv.setLocally(param.name, argVal);

                    // ByRef handling
                    if (!param.isByVal) {
                        let originalExpr: Expression | undefined;
                        if (namedArgExpressions.has(paramNameLower)) {
                            originalExpr = namedArgExpressions.get(paramNameLower);
                        } else if (i < positionalArgExpressions.length) {
                            originalExpr = positionalArgExpressions[i];
                        }
                        if (originalExpr) {
                            byRefArgs.push({
                                paramName: param.name,
                                originalExpr: originalExpr
                            });
                        }
                    }
                }

                return this.execProcBody(proc, localEnv, {
                    byRefArgs,
                    paramArrayParamName,
                    paramArrayByRefExprs,
                    initialLastErrorIndex: -1,
                    acceptExitProperty: false,
                    returnOnProperty: false,
                    subReturnValue: undefined,
                });

                } finally {
                    this.vbaCallStack.pop();
                }
            } else {
                // Might be an array access, built-in function, or variable reference
                // Check explicit declaration BEFORE env.get() which implicitly initializes to 0
                const wasExplicitlyDeclared = this.env.hasVariable(name);

                // §5.6.10 Tier 6 (早期チェック): env に明示宣言がない名前は env.get() で
                // auto-initialize される前に defaultBindingObject を確認する。
                // これをしないと 2 回目の呼び出し時に wasExplicitlyDeclared = true になり
                // OBJECT_REQUIRED になってしまう。
                if (!wasExplicitlyDeclared && this.defaultBindingObject) {
                    const tier6Key = this.resolveObjectMemberKey(this.defaultBindingObject, name.toLowerCase());
                    if (tier6Key !== undefined) {
                        const tier6Member = this.defaultBindingObject[tier6Key];
                        if (typeof tier6Member === 'function') {
                            const argsVals = this.resolveCallArgs(tier6Member, expr.args, name);
                            return (tier6Member as (...a: any[]) => any).apply(this.defaultBindingObject, argsVals);
                        }
                        if (expr.args.length === 0) return tier6Member;
                    }
                }

                const variable = this.env.get(name);
                if (expr.args.length === 0 && typeof variable !== 'function' && !Array.isArray(variable) && !(variable && variable.__isVbaDict__)) {
                    // Undeclared identifier used as a procedure call: "Mainloo" when Sub is "MainLoop"
                    if (!wasExplicitlyDeclared) {
                        this.throwVbaError(VbaErrorCode.SUB_OR_FUNCTION_NOT_DEFINED, `Sub or Function not defined: '${name}'`);
                    }
                    return variable;
                }
                if (typeof variable === 'function') {
                    return variable(...this.resolveCallArgs(variable, expr.args, name));
                } else if (Array.isArray(variable)) {
                    if (expr.args.length === 0) this.throwVbaError(VbaErrorCode.SUBSCRIPT_OUT_OF_RANGE, 'Subscript out of range');
                    const dims = (variable as any).__vbaDimensions__ as { lower: number, upper: number }[] | undefined;
                    if (dims && expr.args.length !== dims.length) this.throwVbaError(VbaErrorCode.SUBSCRIPT_OUT_OF_RANGE, 'Subscript out of range');
                    // VBA index == JS index. Multi-dimensional: arr(i, j) -> arr[i][j]
                    let current = variable;
                    for (let i = 0; i < expr.args.length; i++) {
                        if (!Array.isArray(current)) this.throwVbaError(VbaErrorCode.SUBSCRIPT_OUT_OF_RANGE, 'Subscript out of range');
                        const idx = this.evaluateExpression(expr.args[i]) as number;
                        if (dims) {
                            const { lower, upper } = dims[i];
                            if (idx < lower || idx > upper) this.throwVbaError(VbaErrorCode.SUBSCRIPT_OUT_OF_RANGE, 'Subscript out of range');
                        }
                        current = current[idx];
                    }
                    if (current === undefined) return vbaEmpty;
                    return current;
                } else if (variable && variable.__isVbaDict__) {
                    // Dictionary read: dict("key")
                    if (expr.args.length === 0) this.throwVbaError(VbaErrorCode.ARGUMENT_NOT_OPTIONAL, 'Argument not optional');
                    const key = this.evaluateExpression(expr.args[0]);
                    return variable.__map__.get(key);
                } else if (variable && variable.__isVbaCollection__) {
                    // Collection read: col(index_or_key) -> col.Item(...)
                    const id = this.evaluateExpression(expr.args[0]);
                    return (variable as VbaCollection).item(id);
                } else if (variable && variable.__vbaClass__ && expr.args.length > 0) {
                    // Default property access: obj(args) -> obj.Item(args)
                    const classDef = variable.__classDef__ as ClassDeclaration;
                    // Look for Item property (or Value for single-arg no-index patterns)
                    let defaultProperty = classDef.procedures.find(
                        p => p.isProperty && p.propertyType === 'get' && p.name.name.toLowerCase() === 'item'
                    );
                    if (defaultProperty) {
                        this.checkNoGapOnRequiredParam(defaultProperty.parameters, expr.args);
                        const argsVals = expr.args.map(a => this.resolveAutoInstance(a, this.evaluateExpression(a)));
                        return this.callClassMethod(variable, defaultProperty, argsVals);
                    }
                    this.throwVbaError(VbaErrorCode.OBJECT_DOESNT_SUPPORT_PROPERTY, "Object doesn't support this property or method");
                } else if (variable instanceof VbaNamespaceRef) {
                    // §5.6.10: クラスモジュール名は型名前空間のみ。値名前空間では OBJECT_REQUIRED
                    // でなく SUB_OR_FUNCTION_NOT_DEFINED へフォールスルーする。
                    // defaultBindingObject がある場合は Tier 6 も試みる。
                    if (this.defaultBindingObject) {
                        const tier6Key = this.resolveObjectMemberKey(this.defaultBindingObject, name.toLowerCase());
                        if (tier6Key !== undefined) {
                            const tier6Member = this.defaultBindingObject[tier6Key];
                            if (typeof tier6Member === 'function') {
                                const argsVals = expr.args.map(a => this.resolveAutoInstance(a, this.evaluateExpression(a)));
                                return (tier6Member as (...a: any[]) => any).apply(this.defaultBindingObject, argsVals);
                            }
                            if (expr.args.length === 0) return tier6Member;
                        }
                    }
                } else if (expr.args.length > 0 && wasExplicitlyDeclared) {
                    // Variable is declared but not callable (e.g. Long used as function)
                    this.throwVbaError(VbaErrorCode.OBJECT_REQUIRED, 'Object required');
                }
                this.throwVbaError(VbaErrorCode.SUB_OR_FUNCTION_NOT_DEFINED, `Sub or Function not defined: '${name}'`);
            }
        } else if (expr.callee.type === 'MemberExpression' || expr.callee.type === 'ImplicitWithObjectExpression') {
            let obj: any;
            let methodNameOriginal: string;
            let callObjectName = '';

            if (expr.callee.type === 'MemberExpression') {
                const member = expr.callee as MemberExpression;

                if (member.object.type === 'Identifier') {
                    callObjectName = (member.object as Identifier).name;
                }

                // Check if this is a module-qualified procedure call (e.g., Module.Procedure)
                // member.object is Identifier -> it might be a module name
                if ((member.object as any).type === 'Identifier') {
                    const possibleModuleName = (member.object as any).name;
                    // env.getConst で auto-initialize せずに変数値を取得。
                    // VBA や未宣言モジュール名は undefined を返す（evaluateExpression だと vbaEmpty に初期化される）。
                    // Tier 1-4: env chain, Tier 5: typeLibraryNamespaces（COM 型ライブラリ名前空間）
                    const potentialObj = this.env.getConst(possibleModuleName)
                        ?? this.typeLibraryNamespaces.get(possibleModuleName.toLowerCase());

                    // If evaluating the object gives undefined/null/VbaNamespaceRef, it might be a module/project name
                    if (!potentialObj || potentialObj === vbaEmpty || potentialObj === vbaNull || potentialObj instanceof VbaNamespaceRef) {
                        // VBA 標準ライブラリ名前空間: VBA.InStr(...) は必ず組み込み関数を呼ぶ。
                        // ユーザーが同名の関数を定義していても組み込みが優先される（VBA 仕様）。
                        // variables のみを辿る getConst でユーザー定義（procedures）をスキップ。
                        if (possibleModuleName.toLowerCase() === 'vba' || (potentialObj instanceof VbaNamespaceRef && potentialObj.kind === 'project')) {
                            const builtin = this.env.getConst(member.property.name);
                            if (typeof builtin === 'function') {
                                return builtin(...this.resolveCallArgs(builtin, expr.args, member.property.name));
                            }
                            // 組み込みでない（想定外）場合は従来の経路でエラーを出す
                            return this.evaluateCallExpression({ ...expr, callee: member.property });
                        }

                        // ユーザー定義のモジュール修飾プロシージャ（Module1.Proc など）。
                        // 存在を事前確認してから呼ぶ（try/catch で呼び出し先のランタイムエラーを
                        // 握りつぶさないため）。見つからなければ下の member access へフォールスルー。
                        const qualifiedProc = this.env.getProcedureFromModule(member.property.name, possibleModuleName)
                            ?? this.env.getProcedureFromModule(member.property.name, possibleModuleName, 'get');
                        if (qualifiedProc) {
                            this.checkNoGapOnRequiredParam(qualifiedProc.parameters, expr.args);
                            const argsVals = expr.args.map(a => this.resolveAutoInstance(a, this.evaluateExpression(a)));
                            return this.callProcedure(member.property.name, argsVals, undefined, possibleModuleName);
                        }
                    }
                }

                obj = this.resolveAutoInstance(member.object, this.evaluateExpression(member.object));
                methodNameOriginal = member.property.name;
            } else {
                if (this.withObjectStack.length === 0) {
                    this.throwVbaError(VbaErrorCode.OBJECT_VARIABLE_NOT_SET, 'Object variable or With block variable not set');
                }
                obj = this.withObjectStack[this.withObjectStack.length - 1];
                methodNameOriginal = (expr.callee as ImplicitWithObjectExpression).property.name;
            }

            const methodNameLower = methodNameOriginal.toLowerCase();

            // Nothing / unset object check
            if (obj === null || obj === undefined || obj === vbaNothing) {
                const suffix = callObjectName ? `: '${callObjectName}' is not set` : '';
                this.throwVbaError(VbaErrorCode.OBJECT_VARIABLE_NOT_SET, `Object variable or With block variable not set${suffix}`);
            }

            // VBA class instance method call
            if (obj && obj.__vbaClass__) {
                const classDef = obj.__classDef__ as ClassDeclaration;
                const proc = classDef.procedures.find(p => p.name.name.toLowerCase() === methodNameLower);
                if (proc) {
                    this.checkNoGapOnRequiredParam(proc.parameters, expr.args);
                    const argsVals = expr.args.map(a => this.resolveAutoInstance(a, this.evaluateExpression(a)));
                    return this.callClassMethod(obj, proc, argsVals);
                }
                // Implements interface dispatch: obj.Speak -> obj.IAnimal_Speak
                const ifaceProc = this.findInterfaceDispatch(obj, methodNameOriginal);
                if (ifaceProc) {
                    this.checkNoGapOnRequiredParam(ifaceProc.parameters, expr.args);
                    const argsVals = expr.args.map(a => this.resolveAutoInstance(a, this.evaluateExpression(a)));
                    return this.callClassMethod(obj, ifaceProc, argsVals);
                }
                // プロシージャが見つからない場合、配列フィールドへの外部インデックスアクセスを試みる
                // 例: obj.Items(0)
                {
                    const instanceEnvForRead = obj.__instanceEnv__ as Environment;
                    const fieldArrRead = instanceEnvForRead.get(methodNameLower);
                    if (Array.isArray(fieldArrRead)) {
                        const dims = (fieldArrRead as any).__vbaDimensions__ as { lower: number, upper: number }[] | undefined;
                        let current: any = fieldArrRead;
                        for (let i = 0; i < expr.args.length - 1; i++) {
                            const d = this.evaluateExpression(expr.args[i]) as number;
                            if (dims) {
                                const { lower, upper } = dims[i];
                                if (d < lower || d > upper) this.throwVbaError(VbaErrorCode.SUBSCRIPT_OUT_OF_RANGE, 'Subscript out of range');
                            }
                            current = current[d];
                            if (!Array.isArray(current)) this.throwVbaError(VbaErrorCode.SUBSCRIPT_OUT_OF_RANGE, 'Subscript out of range');
                        }
                        const lastIdx = this.evaluateExpression(expr.args[expr.args.length - 1]) as number;
                        if (dims) {
                            const { lower, upper } = dims[expr.args.length - 1];
                            if (lastIdx < lower || lastIdx > upper) this.throwVbaError(VbaErrorCode.SUBSCRIPT_OUT_OF_RANGE, 'Subscript out of range');
                        }
                        return current[lastIdx];
                    }
                }
                this.throwVbaError(VbaErrorCode.OBJECT_DOESNT_SUPPORT_PROPERTY, `Object doesn't support this property or method: '${methodNameOriginal}'`);
            }

            if (obj) {
                // Try case-insensitive lookup first, then fallback to original casing
                let targetMethod = obj[methodNameLower];

                if (typeof targetMethod !== 'function') {
                    // Search object keys for case-insensitive match (for JS proxies/objects)
                    const keys = Object.keys(obj);
                    // If proxy, Object.keys might not work perfectly, so also check original
                    if (typeof obj[methodNameOriginal] === 'function') {
                        targetMethod = obj[methodNameOriginal];
                    } else {
                        const match = keys.find(k => k.toLowerCase() === methodNameLower);
                        if (match) targetMethod = obj[match];
                    }
                }

                if (typeof targetMethod === 'function') {
                    return targetMethod.apply(obj, this.resolveCallArgs(targetMethod, expr.args, methodNameOriginal));
                }
            }
            this.throwVbaError(VbaErrorCode.OBJECT_DOESNT_SUPPORT_PROPERTY, `Object doesn't support this property or method: '${methodNameOriginal}'`);
        }


        // Generic fallback for calling result of an expression: (expr)(args)
        // e.g. Array(1, 2)(0)
        const target = this.evaluateExpression(expr.callee);
        if (Array.isArray(target)) {
            if (expr.args.length === 0) this.throwVbaError(VbaErrorCode.SUBSCRIPT_OUT_OF_RANGE, 'Subscript out of range');
            let current = target;
            for (let i = 0; i < expr.args.length; i++) {
                if (!current) return vbaEmpty;
                const idx = this.evaluateExpression(expr.args[i]);
                current = current[idx];
            }
            return current === undefined ? vbaEmpty : current;
        } else if (target && target.__isVbaDict__) {
            if (expr.args.length === 0) this.throwVbaError(VbaErrorCode.ARGUMENT_NOT_OPTIONAL, 'Argument not optional');
            const key = this.evaluateExpression(expr.args[0]);
            return target.__map__.get(key);
        } else if (typeof target === 'function') {
            return target(...this.resolveCallArgs(target, expr.args, expr.callee.type));
        }

        this.throwVbaError(VbaErrorCode.OBJECT_REQUIRED, 'Object required');
    }

    private evaluateDictionaryAccessExpression(expr: DictionaryAccessExpression): any {
        const obj = this.evaluateExpression(expr.object);
        const property = expr.property.name;

        if (obj && obj.__isVbaDict__) {
            // VBA bang (!) access is essentially string key lookup: dict!Key -> dict("Key")
            return obj.__map__.get(property);
        }

        // Fallback or error (some objects might support ! besides Dictionary, but we only have Dictionary for now)
        this.throwVbaError(VbaErrorCode.OBJECT_DOESNT_SUPPORT_PROPERTY, "Object doesn't support this property or method");
    }

    private evaluateTypeOfIsExpression(expr: TypeOfIsExpression): any {
        const obj = this.evaluateExpression(expr.expression);
        const typeName = expr.typeName.toLowerCase();

        if (obj === null || obj === undefined || typeof obj !== 'object') return vbaFalse;

        // Check for built-in types
        if (typeName === 'object') return vbaTrue; // Everything that reaches here is an object
        if (typeName === 'dictionary' && obj.__isVbaDict__) return vbaTrue;
        if (typeName === 'collection' && obj.__isVbaCollection__) return vbaTrue;

        // User defined types or classes (if we store metadata)
        if (obj.__vbaTypeName__ && obj.__vbaTypeName__.toLowerCase() === typeName) return vbaTrue;

        return vbaFalse;
    }

    private evaluateUnaryExpression(expr: UnaryExpression): any {
        let argument = this.evaluateExpression(expr.argument);
        const op = expr.operator.toLowerCase();

        // VBA Null / Empty の伝播ルール
        // - 算術系の単項 (-, +): Null は Null、Empty は 0 として扱う
        // - Not: Null は Null
        if (argument === vbaNull) return vbaNull;
        if (argument === vbaEmpty && (op === '-' || op === '+' || op === 'not')) {
            argument = 0;
        }

        switch (op) {
            case 'not':
                // VBA Boolean は -1 / 0 の 2 値のみ。VbaBoolean インスタンスは
                // すべてシングルトン (vbaTrue / vbaFalse) であることが invariant。
                if (argument === vbaTrue) return vbaFalse;
                if (argument === vbaFalse) return vbaTrue;
                if (typeof argument === 'boolean') return argument ? vbaFalse : vbaTrue;
                return ~argument;
            case '-':
                return -argument;
            case '+':
                return +argument;
            default:
                throw new Error(`Execution error: Unknown unary operator ${expr.operator}`);
        }
    }

    /**
     * Auto-Instance placeholder を実際のインスタンスに解決する。
     * 解決後は `Identifier` の場合 env に書き戻して、次回のアクセスではインスタンスが
     * 直接使えるようにする（毎回 instantiate するコストを避ける）。
     * Placeholder でない値はそのまま返す。
     */
    private resolveAutoInstance(sourceExpr: Expression | null, value: any): any {
        if (!isAutoInstancePlaceholder(value)) return value;
        const instance = this.instantiateClass(value.__className__);
        // Identifier 経由のアクセスなら env に書き戻す
        if (sourceExpr && sourceExpr.type === 'Identifier') {
            this.env.set((sourceExpr as Identifier).name, instance);
        }
        return instance;
    }

    /**
     * プレーンな JS オブジェクト（VBA クラスではないモック等）に対し、VBA の大文字小文字
     * 無視ルールで実際のプロパティキーを解決する。getter/setter（アクセサ）や非列挙
     * プロパティも拾えるよう、プロトタイプチェーンを getOwnPropertyNames で辿る。
     * 各階層で厳密一致 → case-insensitive 一致の順に探し、own を prototype より優先する。
     * @param lowerName 小文字化済みのプロパティ名
     * @returns 実際のキー名。見つからなければ undefined。
     */
    private resolveObjectMemberKey(obj: any, lowerName: string): string | undefined {
        let cur = obj;
        while (cur && cur !== Object.prototype && cur !== Function.prototype) {
            const names = Object.getOwnPropertyNames(cur);
            if (names.includes(lowerName)) return lowerName;
            const ci = names.find(k => k.toLowerCase() === lowerName);
            if (ci) return ci;
            cur = Object.getPrototypeOf(cur);
        }
        return undefined;
    }

    private evaluateMemberExpression(expr: MemberExpression): any {
        const propName = expr.property.name.toLowerCase();

        // Check module-qualified variable/constant BEFORE evaluating expr.object.
        // This must come first to avoid Environment.get's implicit-zero initialization
        // turning an undeclared module name into a number.
        if (expr.object.type === 'Identifier') {
            const possibleModule = (expr.object as Identifier).name;
            const moduleKey = `${possibleModule.toLowerCase()}:${propName}`;
            // Constants are stored with module-qualified key (immutable → no sync issue)
            if (this.env.hasVariable(moduleKey)) {
                return this.env.get(moduleKey);
            }
            // VBA standard library module: VBA.vbNull, VBA.vbString, VBA.String$, etc.
            // Look up the property as a plain identifier (all vb* constants are in env).
            if (possibleModule.toLowerCase() === 'vba') {
                const val = this.env.get(propName);
                if (val !== null && val !== undefined) return val;
            }
            // Variables: look up by unqualified name via module registry
            const vars = this.moduleVarRegistry.get(possibleModule.toLowerCase());
            if (vars && vars.has(propName)) {
                return this.env.get(propName);
            }
        }

        const evaluated = this.evaluateExpression(expr.object);
        const obj = this.resolveAutoInstance(expr.object, evaluated);

        // Safety check: ensure obj is an object before trying member access
        if (obj === null || obj === undefined || obj === vbaNothing) {
            const objName = expr.object.type === 'Identifier'
                ? `: '${(expr.object as Identifier).name}' is not set`
                : '';
            this.throwVbaError(VbaErrorCode.OBJECT_VARIABLE_NOT_SET, `Object variable or With block variable not set${objName}`);
        } else if (typeof obj !== 'object' && typeof obj !== 'function') {
            this.throwVbaError(VbaErrorCode.OBJECT_REQUIRED, 'Object required');
        }

        // VBA class instance: look up field in instance environment or invoke Property Get
        if (obj && obj.__vbaClass__) {
            const classDef = obj.__classDef__ as ClassDeclaration;
            const instanceEnv = obj.__instanceEnv__ as Environment;

            // Check for Property Get
            const getter = classDef.procedures.find(
                p => p.isProperty && p.propertyType === 'get' && p.name.name.toLowerCase() === propName
            );
            if (getter) {
                return this.callClassMethod(obj, getter, []);
            }

            // No-arg Sub/Function access without parens (rare, treat as property-like call)
            const method = classDef.procedures.find(
                p => !p.isProperty && p.name.name.toLowerCase() === propName && p.parameters.length === 0 && p.isFunction
            );
            if (method) {
                return this.callClassMethod(obj, method, []);
            }

            // Implements interface dispatch: obj.Area -> obj.IShape_Area
            const ifaceProc = this.findInterfaceDispatch(obj, expr.property.name);
            if (ifaceProc) {
                return this.callClassMethod(obj, ifaceProc, []);
            }

            // Field access
            return instanceEnv.get(propName);
        }

        // case-insensitive にプロパティを解決（アクセサ・プロトタイプ・非列挙も対象）
        if (obj && (typeof obj === 'object' || typeof obj === 'function')) {
            const key = this.resolveObjectMemberKey(obj, propName);
            if (key !== undefined) {
                const val = obj[key];
                // Auto-call only zero-arg functions (VBA property/method without parens like col.Count, ws.Add)
                // Functions requiring args (like Worksheets(name)) are returned as references
                if (typeof val === 'function' && val.length === 0) {
                    return val.call(obj);
                }
                return val;
            }
        }
        this.throwVbaError(VbaErrorCode.OBJECT_DOESNT_SUPPORT_PROPERTY, `Object doesn't support this property or method: '${propName}'`);
    }

    /**
     * Implements インターフェースディスパッチ:
     * obj.Speak() で Speak が見つからない場合、obj の Implements IAnimal から
     * IAnimal に Speak があれば IAnimal_Speak にディスパッチする。
     */
    private findInterfaceDispatch(obj: any, methodNameOriginal: string): ProcedureDeclaration | null {
        if (!obj || !obj.__vbaClass__) return null;
        const classDef = obj.__classDef__ as ClassDeclaration;
        const methodNameLower = methodNameOriginal.toLowerCase();

        for (const stmt of classDef.body) {
            if (stmt.type !== 'ImplementsDirective') continue;
            const ifaceName = (stmt as ImplementsDirective).interfaceName;
            const ifaceClass = this.classDefinitions.get(ifaceName.toLowerCase());
            if (!ifaceClass) continue;

            // interface has the requested member?
            const inInterface = ifaceClass.procedures.some(
                p => p.name.name.toLowerCase() === methodNameLower
            ) || ifaceClass.fields.some(
                f => f.declarations.some(d => d.name.name.toLowerCase() === methodNameLower)
            );
            if (!inInterface) continue;

            // look for IfaceName_MethodName in concrete class
            const implName = `${ifaceName}_${methodNameOriginal}`.toLowerCase();
            const implProc = classDef.procedures.find(p => p.name.name.toLowerCase() === implName);
            if (implProc) return implProc;
        }
        return null;
    }

    private evaluateBinaryExpression(expr: BinaryExpression): any {
        let leftVal = this.evaluateExpression(expr.left);
        let rightVal = this.evaluateExpression(expr.right);

        // Normalize booleans to VBA integers (-1, 0)
        if (leftVal === true) leftVal = vbaTrue;
        if (leftVal === false) leftVal = vbaFalse;
        if (rightVal === true) rightVal = vbaTrue;
        if (rightVal === false) rightVal = vbaFalse;

        const op = expr.operator.toLowerCase();

        // ===== VBA Null / Empty の伝播ルール =====
        // 算術 / 比較 / 論理演算では Null が含まれていれば結果は Null。
        // 例外: 文字列連結 & では Null / Empty を "" として扱う。
        const arithmeticOps = new Set(['+', '-', '*', '/', '\\', 'mod', '^']);
        const comparisonOps = new Set(['=', '<>', '<', '>', '<=', '>=']);
        const logicalOps = new Set(['and', 'or', 'xor', 'eqv', 'imp']);

        if (op === '&') {
            // 文字列連結: Null も Empty も "" 扱い
            if (leftVal === vbaNull || leftVal === vbaEmpty) leftVal = '';
            if (rightVal === vbaNull || rightVal === vbaEmpty) rightVal = '';
        } else if (arithmeticOps.has(op) || comparisonOps.has(op) || logicalOps.has(op)) {
            // Null 伝播: どちらかが Null なら Null
            if (leftVal === vbaNull || rightVal === vbaNull) return vbaNull;
            // Empty を数値文脈では 0、文字列文脈では "" に正規化
            if (leftVal === vbaEmpty) {
                leftVal = typeof rightVal === 'string' ? '' : 0;
            }
            if (rightVal === vbaEmpty) {
                rightVal = typeof leftVal === 'string' ? '' : 0;
            }
        }

        // String → Number の暗黙変換
        // - `+` は両方 String なら連結、それ以外は数値加算 (VBA 仕様)
        // - 他の算術演算 (-, *, /, \, Mod, ^) は常に数値変換
        // - 変換できない文字列は Type mismatch (Error 13)
        const toVbaNumber = (v: any): number => {
            if (typeof v === 'number') return v;
            if (typeof v === 'bigint') return Number(v);
            if (v instanceof VbaBoolean) return v.value;
            if (v instanceof VbaDate) return v.value;
            if (typeof v === 'string') {
                const trimmed = v.trim();
                if (trimmed === '') return 0;
                const n = Number(trimmed);
                if (isNaN(n)) this.throwVbaError(VbaErrorCode.TYPE_MISMATCH, 'Type mismatch');
                return n;
            }
            this.throwVbaError(VbaErrorCode.TYPE_MISMATCH, 'Type mismatch');
        };

        // 比較演算子では VbaBoolean を数値として扱う（True=-1, False=0）
        if (comparisonOps.has(op)) {
            if (leftVal instanceof VbaBoolean) leftVal = leftVal.value;
            if (rightVal instanceof VbaBoolean) rightVal = rightVal.value;
        }

        const isPlusConcatenation = op === '+' && typeof leftVal === 'string' && typeof rightVal === 'string';
        const numericArithOps = new Set(['-', '*', '/', '\\', 'mod', '^']);
        if (numericArithOps.has(op) || (op === '+' && !isPlusConcatenation)) {
            // Date は VbaDate のままにして、後続の case で処理させる
            if (!(leftVal instanceof VbaDate)) leftVal = toVbaNumber(leftVal);
            if (!(rightVal instanceof VbaDate)) rightVal = toVbaNumber(rightVal);
        }

        switch (op) {
            case '+': {
                if (isPlusConcatenation) return leftVal + rightVal;
                const l = leftVal instanceof VbaDate ? leftVal.value : leftVal;
                const r = rightVal instanceof VbaDate ? rightVal.value : rightVal;
                const sum = l + r;
                return (leftVal instanceof VbaDate || rightVal instanceof VbaDate) ? new VbaDate(sum) : sum;
            }
            case '&': return this.toDisplayString(leftVal) + this.toDisplayString(rightVal);
            case '-': {
                const l = leftVal instanceof VbaDate ? leftVal.value : leftVal;
                const r = rightVal instanceof VbaDate ? rightVal.value : rightVal;
                const diff = l - r;
                if (leftVal instanceof VbaDate && rightVal instanceof VbaDate) return diff; // Date - Date = Number
                return (leftVal instanceof VbaDate) ? new VbaDate(diff) : diff;
            }
            case '*': return leftVal * rightVal;
            case '/':
                if (rightVal === 0) this.throwVbaError(VbaErrorCode.DIVISION_BY_ZERO, 'Division by zero');
                return leftVal / rightVal;
            case '\\':
                if (rightVal === 0) this.throwVbaError(VbaErrorCode.DIVISION_BY_ZERO, 'Division by zero');
                return Math.trunc(leftVal / rightVal);
            case 'mod':
                if (rightVal === 0) this.throwVbaError(VbaErrorCode.DIVISION_BY_ZERO, 'Division by zero');
                return leftVal % rightVal;
            case '^': return Math.pow(leftVal, rightVal);
            case '=':
                if (leftVal instanceof VbaErrorValue && rightVal instanceof VbaErrorValue) {
                    return leftVal.code === rightVal.code ? vbaTrue : vbaFalse;
                }
                if (leftVal instanceof VbaDate && rightVal instanceof VbaDate) {
                    return leftVal.value === rightVal.value ? vbaTrue : vbaFalse;
                }
                if (typeof leftVal === 'string' && typeof rightVal === 'string' && this.comparisonMode === 'Text') {
                    return leftVal.toLowerCase() === rightVal.toLowerCase() ? vbaTrue : vbaFalse;
                }
                return leftVal === rightVal ? vbaTrue : vbaFalse;
            case '<>':
                if (leftVal instanceof VbaErrorValue && rightVal instanceof VbaErrorValue) {
                    return leftVal.code !== rightVal.code ? vbaTrue : vbaFalse;
                }
                if (leftVal instanceof VbaDate && rightVal instanceof VbaDate) {
                    return leftVal.value !== rightVal.value ? vbaTrue : vbaFalse;
                }
                if (typeof leftVal === 'string' && typeof rightVal === 'string' && this.comparisonMode === 'Text') {
                    return leftVal.toLowerCase() !== rightVal.toLowerCase() ? vbaTrue : vbaFalse;
                }
                return leftVal !== rightVal ? vbaTrue : vbaFalse;
            case '<':
                if (typeof leftVal === 'string' && typeof rightVal === 'string' && this.comparisonMode === 'Text') {
                    return leftVal.toLowerCase() < rightVal.toLowerCase() ? vbaTrue : vbaFalse;
                }
                return leftVal < rightVal ? vbaTrue : vbaFalse;
            case '>':
                if (typeof leftVal === 'string' && typeof rightVal === 'string' && this.comparisonMode === 'Text') {
                    return leftVal.toLowerCase() > rightVal.toLowerCase() ? vbaTrue : vbaFalse;
                }
                return leftVal > rightVal ? vbaTrue : vbaFalse;
            case '<=':
                if (typeof leftVal === 'string' && typeof rightVal === 'string' && this.comparisonMode === 'Text') {
                    return leftVal.toLowerCase() <= rightVal.toLowerCase() ? vbaTrue : vbaFalse;
                }
                return leftVal <= rightVal ? vbaTrue : vbaFalse;
            case '>=':
                if (typeof leftVal === 'string' && typeof rightVal === 'string' && this.comparisonMode === 'Text') {
                    return leftVal.toLowerCase() >= rightVal.toLowerCase() ? vbaTrue : vbaFalse;
                }
                return leftVal >= rightVal ? vbaTrue : vbaFalse;
            case 'is': return leftVal === rightVal ? vbaTrue : vbaFalse;
            case 'like': return this.evaluateLike(leftVal, rightVal) ? vbaTrue : vbaFalse;
            case 'and':
                const andRes = leftVal & rightVal;
                return (leftVal instanceof VbaBoolean && rightVal instanceof VbaBoolean) ? (andRes === -1 ? vbaTrue : vbaFalse) : andRes;
            case 'or':
                const orRes = leftVal | rightVal;
                return (leftVal instanceof VbaBoolean && rightVal instanceof VbaBoolean) ? (orRes === -1 ? vbaTrue : vbaFalse) : orRes;
            case 'xor':
                const xorRes = leftVal ^ rightVal;
                return (leftVal instanceof VbaBoolean && rightVal instanceof VbaBoolean) ? (xorRes === -1 ? vbaTrue : vbaFalse) : xorRes;
            case 'eqv':
                const eqvRes = ~(leftVal ^ rightVal);
                return (leftVal instanceof VbaBoolean && rightVal instanceof VbaBoolean) ? (eqvRes === -1 ? vbaTrue : vbaFalse) : eqvRes;
            case 'imp':
                const impRes = (~leftVal) | rightVal;
                return (leftVal instanceof VbaBoolean && rightVal instanceof VbaBoolean) ? (impRes === -1 ? vbaTrue : vbaFalse) : impRes;
            default:
                throw new Error(`Execution error: Unknown operator ${expr.operator}`);
        }
    }

    private evaluateImplicitWithObjectExpression(expr: ImplicitWithObjectExpression): any {
        if (this.withObjectStack.length === 0) {
            this.throwVbaError(VbaErrorCode.OBJECT_VARIABLE_NOT_SET, 'Object variable or With block variable not set');
        }
        const obj = this.withObjectStack[this.withObjectStack.length - 1];
        const propName = expr.property.name.toLowerCase();
        if (obj && (typeof obj === 'object' || typeof obj === 'function')) {
            // resolveObjectMemberKey でプロトタイプチェーン（getter/メソッド含む）を辿る。
            // 旧実装は Object.keys() のみで own property しか探さなかったため、
            // TypeScript class の get accessor が With ブロック内で発見できないバグがあった。
            const key = this.resolveObjectMemberKey(obj, propName);
            if (key !== undefined) {
                const val = obj[key];
                if (typeof val === 'function' && val.length === 0) {
                    return val.call(obj);
                }
                if (typeof val === 'function') {
                    return val.bind(obj);
                }
                return val;
            }
        }
        if (obj === null || obj === undefined || obj === vbaNothing) {
            this.throwVbaError(VbaErrorCode.OBJECT_VARIABLE_NOT_SET, 'Object variable or With block variable not set');
        } else {
            this.throwVbaError(VbaErrorCode.OBJECT_REQUIRED, 'Object required');
        }
    }

    private evaluateLike(text: any, pattern: any): boolean {
        const textStr = String(text);
        const patternStr = String(pattern);

        // Convert VBA Like pattern to Regex
        // Special characters in VBA Like: *, ?, #, [
        // Regex special characters to escape: \, ^, $, ., |, (, ), [, ], {, }, +, * , ?
        let regexStr = '^';
        let i = 0;
        while (i < patternStr.length) {
            const char = patternStr[i];
            if (char === '*') {
                regexStr += '.*';
            } else if (char === '?') {
                regexStr += '.';
            } else if (char === '#') {
                regexStr += '\\d';
            } else if (char === '[') {
                regexStr += '[';
                i++;
                if (i < patternStr.length && patternStr[i] === '!') {
                    regexStr += '^';
                    i++;
                }
                while (i < patternStr.length && patternStr[i] !== ']') {
                    const charInList = patternStr[i];
                    // Escape regex special chars inside [] if they are not part of range
                    if ('\\^$[]{}|()+.'.includes(charInList) && charInList !== '-') {
                        regexStr += '\\' + charInList;
                    } else {
                        regexStr += charInList;
                    }
                    i++;
                }
                if (i < patternStr.length) {
                    regexStr += ']';
                }
            } else {
                // Escape regex special characters
                if ('\\^$[]{}|()+.'.includes(char)) {
                    regexStr += '\\' + char;
                } else {
                    regexStr += char;
                }
            }
            i++;
        }
        regexStr += '$';

        try {
            const flags = this.comparisonMode === 'Text' ? 'i' : '';
            const regex = new RegExp(regexStr, flags);
            return regex.test(textStr);
        } catch (e) {
            return false;
        }
    }

    private isTrue(val: any): boolean {
        if (val === vbaNull) this.throwVbaError(VbaErrorCode.INVALID_USE_OF_NULL, 'Invalid use of Null');
        if (val === undefined || val === null) return false;
        if (val instanceof VbaBoolean) return val.value !== 0;
        if (typeof val === 'number') return val !== 0;
        if (typeof val === 'boolean') return val;
        // 文字列は §5.6.9 + §6.1.2.3 の value coercion（CBool と同規則）で変換
        if (typeof val === 'string') return this.coerceToBoolean(val).value !== 0;
        return !!val;
    }

    private formatDate(d: Date, pattern: string): string {
        const pLower = pattern.toLowerCase();

        const pad2 = (n: number) => String(n).padStart(2, '0');
        const yyyy = String(d.getFullYear());
        const MM   = pad2(d.getMonth() + 1);
        const dd   = pad2(d.getDate());
        const HH   = pad2(d.getHours());
        const mm   = pad2(d.getMinutes());
        const ss   = pad2(d.getSeconds());

        const months      = ["January","February","March","April","May","June","July","August","September","October","November","December"];
        const monthsShort = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
        const days        = ["Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday"];
        const daysShort   = ["Sun","Mon","Tue","Wed","Thu","Fri","Sat"];

        const h12  = d.getHours() % 12 || 12;
        const ampm = d.getHours() >= 12 ? 'PM' : 'AM';

        // Named formats — fixed English format strings (VBA named formats are locale-dependent in real VBA,
        // but we use fixed format strings to ensure consistent output regardless of OS locale settings)
        if (pLower === 'general date') {
            const hasTime = d.getHours() !== 0 || d.getMinutes() !== 0 || d.getSeconds() !== 0;
            return hasTime ? `${yyyy}/${MM}/${dd} ${HH}:${mm}:${ss}` : `${yyyy}/${MM}/${dd}`;
        }
        if (pLower === 'long date')   return `${days[d.getDay()]}, ${months[d.getMonth()]} ${d.getDate()}, ${yyyy}`;
        if (pLower === 'medium date') return `${dd}-${monthsShort[d.getMonth()]}-${yyyy.slice(-2)}`;
        if (pLower === 'short date')  return `${yyyy}/${MM}/${dd}`;
        if (pLower === 'long time')   return `${HH}:${mm}:${ss}`;
        if (pLower === 'medium time') return `${h12}:${mm} ${ampm}`;
        if (pLower === 'short time')  return `${HH}:${mm}`;

        // VBA Format: mm/m is context-sensitive — minutes if immediately after h/hh, otherwise month
        // prevTokenWasHour tracks whether the previous format token was an hour (h/hh),
        // so the next mm/m is interpreted as minutes rather than month.
        const tokens = pattern.match(/yyyy|yy|mmmm|mmm|mm|m|dddd|ddd|dd|d|hh|h|nn|n|ss|s|AM\/PM|am\/pm|[^a-zA-Z]+|[a-zA-Z]/gi) || [];
        let prevTokenWasHour = false;
        return tokens.map(tok => {
            const tl = tok.toLowerCase();
            switch (tl) {
                case 'yyyy': prevTokenWasHour = false; return yyyy;
                case 'yy':   prevTokenWasHour = false; return yyyy.slice(-2);
                case 'mmmm': prevTokenWasHour = false; return months[d.getMonth()];
                case 'mmm':  prevTokenWasHour = false; return monthsShort[d.getMonth()];
                case 'mm': {
                    const isMinutes = prevTokenWasHour;
                    prevTokenWasHour = false;
                    return isMinutes ? mm : pad2(d.getMonth() + 1);
                }
                case 'm': {
                    const isMinutes = prevTokenWasHour;
                    prevTokenWasHour = false;
                    return isMinutes ? String(d.getMinutes()) : String(d.getMonth() + 1);
                }
                case 'dddd': prevTokenWasHour = false; return days[d.getDay()];
                case 'ddd':  prevTokenWasHour = false; return daysShort[d.getDay()];
                case 'dd':   prevTokenWasHour = false; return pad2(d.getDate());
                case 'd':    prevTokenWasHour = false; return String(d.getDate());
                case 'hh':   prevTokenWasHour = true;  return HH;
                case 'h':    prevTokenWasHour = true;  return String(d.getHours());
                case 'nn':   prevTokenWasHour = false; return mm;
                case 'n':    prevTokenWasHour = false; return String(d.getMinutes());
                case 'ss':   prevTokenWasHour = false; return ss;
                case 's':    prevTokenWasHour = false; return String(d.getSeconds());
                case 'am/pm': prevTokenWasHour = false; return tok === 'AM/PM' ? ampm : ampm.toLowerCase();
                default:     return tok;
            }
        }).join('');
    }

    private formatNumber(n: number, pattern: string): string {
        const pLower = pattern.toLowerCase();
        // locale-independent thousands separator helper
        const withThousands = (num: number, decimals: number): string => {
            const fixed = Math.abs(num).toFixed(decimals);
            const [intPart, decPart] = fixed.split('.');
            const grouped = intPart.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
            const sign = num < 0 ? '-' : '';
            return decPart !== undefined ? `${sign}${grouped}.${decPart}` : `${sign}${grouped}`;
        };

        // Handle named formats
        switch (pLower) {
            case 'general number': return String(n);
            case 'currency': return (n < 0 ? '-$' : '$') + withThousands(n, 2).replace(/^-/, '');
            case 'fixed':    return n.toFixed(2);
            case 'standard': return withThousands(n, 2);
            case 'percent':  return (n * 100).toFixed(2) + '%';
            case 'scientific': return n.toExponential(2);
            case 'true/false': return n !== 0 ? 'True' : 'False';
            case 'yes/no':   return n !== 0 ? 'Yes' : 'No';
            case 'on/off':   return n !== 0 ? 'On' : 'Off';
        }

        // Custom patterns: #, 0, ., ,, %
        let fmt = pattern;
        let isPercent = false;
        if (fmt.includes('%')) {
            isPercent = true;
            n *= 100;
            fmt = fmt.replace('%', '');
        }

        const hasThousands = fmt.includes(',');
        const parts = fmt.split('.');
        const decimalPart = parts.length > 1 ? parts[1] : '';

        // Determine decimal places
        const minDecimals = (decimalPart.match(/0/g) || []).length;
        const maxDecimals = decimalPart.length;

        const absFixed = Math.abs(n).toFixed(maxDecimals);
        const [intPart, decPart] = absFixed.split('.');
        const intFormatted = hasThousands
            ? intPart.replace(/\B(?=(\d{3})+(?!\d))/g, ',')
            : intPart;
        const decFormatted = decPart !== undefined
            ? decPart.replace(/0+$/, '').padEnd(minDecimals, '0') || (minDecimals > 0 ? '0'.repeat(minDecimals) : '')
            : '';
        let result = (n < 0 ? '-' : '') + intFormatted + (decFormatted ? '.' + decFormatted : '');

        if (isPercent) result += '%';
        return result;
    }
}
