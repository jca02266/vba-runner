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
    ConstDeclaratorItem,
    SetStatement,
    EraseStatement,
    ReDimStatement,
    ReDimDeclarator,
    ArrayBound,
    ExitStatement,
    OnErrorStatement,
    TypeDeclaration,
    TypeMember,
    SelectCaseStatement,
    ForEachStatement,
    WithStatement,
    ImplicitWithObjectExpression,
    ImplicitWithDictionaryAccessExpression,
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
    DebugPrintStatement,
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
    ByValArgument,
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
import { findClassProperty } from './property-resolution';
import { SandboxPath } from './sandbox';
import { FileSystem, MemoryFileSystem, VBA_FILE_ATTRIBUTE } from './filesystem';
import { checkOptionExplicit, UndefinedProcError } from './option-explicit-checker';
import * as path from 'path';
import iconv from 'iconv-lite';
import {
    VbaBoolean, VbaDate, VbaDecimal, VbaErrorValue, VbaNamespaceRef, VbaCurrency,
    vbaEmpty, vbaNull, vbaNothing, vbaMissing, vbaOmitted,
    vbaTrue, vbaFalse,
    toVbaDate,
    fromVbaDate,
    parseVbaDate,
    tryParseTimeFractionString,
    parseFixedPointString, bankersDivide, parseCurrencyString,
    createAutoInstancePlaceholder, isAutoInstancePlaceholder,
    isVbaObjectReferenceCompatible, isVbaBoundObject,
} from './vba-types';
import type { VbaVarType, VbaComObject } from './vba-types';
export {
    VbaBoolean, VbaDate, VbaErrorValue, VbaNamespaceRef, VbaCurrency,
    vbaEmpty, vbaNull, vbaNothing, vbaMissing,
    vbaTrue, vbaFalse,
    toVbaDate,
    createAutoInstancePlaceholder, isAutoInstancePlaceholder,
} from './vba-types';
export type { AutoInstancePlaceholder, VbaBooleanType, VbaNumericType, VbaVarType } from './vba-types';
import {
    vbaToNumber as _vbaToNumber,
    vbaToBoolean,
    vbaToDisplayString,
    vbaToString,
    vbaRound as _vbaRound,
} from './coerce';
import { VbaErrorCode, throwVbaError, VBA_ERROR_MESSAGES } from './vba-errors';
export { VbaErrorCode } from './vba-errors';

function vbaFlagIsTrue(value: any): boolean {
    return vbaToBoolean(value).value !== 0;
}

function encodeTextStream(text: string, unicode: boolean): Uint8Array {
    if (!unicode) return iconv.encode(text, 'cp932');
    const body = new Uint8Array(text.length * 2);
    let offset = 0;
    for (let i = 0; i < text.length; i++) {
        const code = text.charCodeAt(i);
        body[offset++] = code & 0xff;
        body[offset++] = code >>> 8;
    }
    const bytes = new Uint8Array(body.length + 2);
    bytes[0] = 0xff;
    bytes[1] = 0xfe;
    bytes.set(body, 2);
    return bytes;
}

function decodeTextStream(bytes: Uint8Array, unicodeHint: boolean): string {
    let start = 0;
    let unicode = unicodeHint;
    if (bytes.length >= 2 && bytes[0] === 0xff && bytes[1] === 0xfe) {
        unicode = true;
        start = 2;
    }
    if (bytes.length >= 2 && bytes[0] === 0xfe && bytes[1] === 0xff) {
        unicode = true;
        start = 2;
    }
    if (!unicode) return iconv.decode(Buffer.from(bytes), 'cp932');
    let text = '';
    for (let i = start; i + 1 < bytes.length; i += 2) {
        text += String.fromCharCode(bytes[i] | (bytes[i + 1] << 8));
    }
    return text;
}
import {
    type StdlibCtx, type BuiltinParamSpec, type BuiltinOverload,
    registerInformationFunctions,
    registerConversionFunctions,
    registerMathFunctions,
    registerStringFunctions,
    registerStdlibDateTimeFunctions,
    registerInteractionFunctions,
    registerFinancialFunctions,
    registerConstants,
    registerRegistryFunctions,
} from './builtins';
export type { BuiltinParamSpec, BuiltinOverload } from './builtins';

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
// 名前付き引数（Key:= / Before:= / After:=）を resolveCallArgs で解決できるように
// パラメーター仕様を付与する（Bug 33-B: 仕様がないと名前付き引数が位置引数に化けていた）
(VbaCollection.prototype.add as any).__vbaParamSpec__ = [
    { name: 'Item' },
    { name: 'Key', optional: true },
    { name: 'Before', optional: true },
    { name: 'After', optional: true },
];
(VbaCollection.prototype.item as any).__vbaParamSpec__ = [{ name: 'Index' }];
(VbaCollection.prototype.remove as any).__vbaParamSpec__ = [{ name: 'Index' }];

export class VbaErrObject {
    public number: number = 0;
    public source: string = "";
    public description: string = "";
    public helpfile: string = "";
    public helpcontext: number = 0;
    public lastdllerror: number = 0;

    /** Erl: エラー発生時点で最後に通過した数値行ラベル（行番号なしなら 0） */
    public erl: number = 0;

    public clear() {
        this.number = 0;
        this.source = "";
        this.description = "";
        this.helpfile = "";
        this.helpcontext = 0;
        this.lastdllerror = 0;
        this.erl = 0;
    }

    public raise(number: number, source?: any, description?: any, helpfile?: any, helpcontext?: any) {
        this.number = number;
        // VBA uses the current project name when Source is omitted.  Do not
        // retain a custom Source from an earlier error.
        if (source !== undefined && source !== vbaEmpty && source !== null) {
            this.source = String(source);
        } else {
            this.source = 'VBAProject';
        }
        if (description !== undefined && description !== vbaEmpty && description !== null) {
            this.description = String(description);
        } else {
            // §6.1.3.2 Raise: Description 未指定時は Number に対応する Error 関数の文字列、
            // 対応する VBA エラーがなければ "Application-defined or object-defined error"
            this.description = VBA_ERROR_MESSAGES[number] ?? 'Application-defined or object-defined error';
        }
        if (helpfile !== undefined && helpfile !== vbaEmpty && helpfile !== null) this.helpfile = String(helpfile);
        if (helpcontext !== undefined && helpcontext !== vbaEmpty && helpcontext !== null) this.helpcontext = Number(helpcontext);

        throw {
            type: 'VbaError',
            number: this.number,
            message: this.description,
            vbaSource: this.source,
        };
    }
}
export interface VbaTypeInfo {
    vbaType: VbaVarType;
    /** Fixed-length string: Dim s As String * N */
    fixedLength?: number;
    /** User-defined class required for an object reference. */
    objectTypeName?: string;
}

interface BinaryValueLayout {
    typeName: string;
    /** Element type when serializing a declared primitive array. */
    elementType?: string;
    fixedLength?: number;
    /** Compatibility path for an untyped target, where VBA's destination size
     * is unavailable and the runner must retain its historical whole-field read. */
    readToEnd?: boolean;
    /** UDT variable-length String members carry a 2-byte byte-length descriptor. */
    variableLengthDescriptor?: boolean;
}

export interface DebugHook {
    onBeforeStatement(
        line: number,
        callDepth: number,
        env: Environment,
        callStack: ReadonlyArray<{ name: string; moduleName: string; line: number }>
    ): void;
}

interface VbaLValueReference {
    get(): any;
    set(value: any): void;
}

interface ExecProcBodyOptions {
    byRefArgs: { paramName: string; reference: VbaLValueReference }[];
    paramArrayParamName: string | null;
    paramArrayReferences: (VbaLValueReference | null)[];
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
    private arrayTypeInfo: Map<string, {
        elementType?: string;
        elementTypeName?: string;
        elementObjectTypeName?: string;
        elementAutoNew?: boolean;
    }> = new Map();
    /** Dynamic numeric subtype for Variant variables (e.g. after `v = 42`, tracks "Integer") */
    private variantNumericSubtypes: Map<string, VbaVarType> = new Map();
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

    /** Class-field metadata must not fall through to a caller's same-named variable. */
    getOwnVariableType(name: string): VbaTypeInfo | undefined {
        return this.variableTypes.get(name.toLowerCase());
    }

    setArrayTypeInfo(name: string, value: any): void {
        if (!Array.isArray(value)) return;
        this.arrayTypeInfo.set(name.toLowerCase(), {
            elementType: (value as any).__vbaElementType__,
            elementTypeName: (value as any).__vbaElementTypeName__,
            elementObjectTypeName: (value as any).__vbaElementObjectTypeName__,
            elementAutoNew: (value as any).__vbaElementAutoNew__,
        });
    }

    getArrayTypeInfo(name: string): { elementType?: string; elementTypeName?: string; elementObjectTypeName?: string; elementAutoNew?: boolean } | undefined {
        const key = name.toLowerCase();
        if (this.arrayTypeInfo.has(key)) return this.arrayTypeInfo.get(key);
        return this.enclosing?.getArrayTypeInfo(name);
    }

    getOwnArrayTypeInfo(name: string): { elementType?: string; elementTypeName?: string; elementObjectTypeName?: string; elementAutoNew?: boolean } | undefined {
        return this.arrayTypeInfo.get(name.toLowerCase());
    }

    setVariantSubtype(name: string, subtype: VbaVarType): void {
        const key = name.toLowerCase();
        if (this.variables.has(key)) { this.variantNumericSubtypes.set(key, subtype); return; }
        let env: Environment | undefined = this.enclosing;
        while (env) {
            if (env.variables.has(key)) { env.variantNumericSubtypes.set(key, subtype); return; }
            env = env.enclosing;
        }
        this.variantNumericSubtypes.set(key, subtype);
    }

    getVariantSubtype(name: string): VbaVarType | undefined {
        const key = name.toLowerCase();
        if (this.variantNumericSubtypes.has(key)) return this.variantNumericSubtypes.get(key);
        let env: Environment | undefined = this.enclosing;
        while (env) {
            if (env.variantNumericSubtypes.has(key)) return env.variantNumericSubtypes.get(key);
            env = env.enclosing;
        }
        return undefined;
    }

    clearVariantSubtype(name: string): void {
        const key = name.toLowerCase();
        if (this.variantNumericSubtypes.has(key)) { this.variantNumericSubtypes.delete(key); return; }
        let env: Environment | undefined = this.enclosing;
        while (env) {
            if (env.variantNumericSubtypes.has(key)) { env.variantNumericSubtypes.delete(key); return; }
            env = env.enclosing;
        }
    }

    private coerceToType(key: string, value: any): any {
        const typeInfo = this.variableTypes.get(key);
        if (!typeInfo) return value;

        // Don't coerce special VBA values
        if (value === vbaEmpty || value === undefined || value === vbaNull || value === vbaNothing || value === vbaMissing) {
            return value;
        }

        if (typeInfo.objectTypeName) {
            const actualTypeName = (value as any)?.__className__ as string | undefined;
            const implementsRequiredType = (value as any)?.__classDef__?.body?.some((stmt: any) =>
                stmt.type === 'ImplementsDirective' &&
                stmt.interfaceName.toLowerCase() === typeInfo.objectTypeName!.toLowerCase());
            if (!actualTypeName || (
                actualTypeName.toLowerCase() !== typeInfo.objectTypeName.toLowerCase() &&
                !implementsRequiredType
            )) {
                throwVbaError(VbaErrorCode.TYPE_MISMATCH, 'Type mismatch');
            }
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
            case 'LongLong':
            case 'LongPtr': {
                let n: bigint;
                if (typeof value === 'bigint') {
                    n = value;
                } else if (typeof value === 'string' && /^[+-]?\d+$/.test(value.trim())) {
                    n = BigInt(value.trim());
                } else {
                    n = BigInt(Environment.vbaRoundStatic(Environment.toNumeric(value)));
                }
                if (n < -9223372036854775808n || n > 9223372036854775807n) {
                    Environment.throwOverflow();
                }
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
                if (value instanceof VbaCurrency) return value;
                if (value instanceof VbaBoolean) return new VbaCurrency(BigInt(value.value) * 10000n);
                if (typeof value === 'bigint') return new VbaCurrency(value * 10000n);
                if (typeof value === 'string') {
                    const trimmed = value.trim();
                    const decimalString = trimmed.replace(/,/g, '');
                    if (!/^\+?(\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?$/.test(decimalString) &&
                        !/^-?(\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?$/.test(decimalString)) {
                        Environment.throwOverflow();
                    }
                    return parseCurrencyString(trimmed);
                }
                return VbaCurrency.fromNumber(Environment.vbaRoundStatic(Environment.toNumeric(value), 4));
            }
            case 'String': {
                let s: string;
                if (typeof value === 'number' || value instanceof VbaBoolean || value instanceof VbaDate) {
                    s = String(value);
                } else {
                    s = typeof value === 'string' ? value : String(value);
                }
                if (typeInfo.fixedLength !== undefined) {
                    const n = typeInfo.fixedLength;
                    if (s.length > n) return s.slice(0, n);
                    if (s.length < n) return s + ' '.repeat(n - s.length);
                }
                return s;
            }
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

    hasOwnWithEvents(name: string): boolean {
        return this.withEventsVariables.has(name.toLowerCase());
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

    promoteProceduresFromModule(moduleName: string): void {
        const prefix = moduleName.toLowerCase() + ':';
        for (const [key, proc] of this.procedures.entries()) {
            if (key.startsWith(prefix)) {
                this.procedures.set(key.slice(prefix.length), proc);
            }
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

/** `BuiltinParamSpec` と最小限の形を共有する抽象（ProcedureDeclaration.parameters も満たす） */
interface ArgBinderParam {
    isOptional: boolean;
    isParamArray: boolean;
}

export class Evaluator {
    public env: Environment;          // Tier 3: Public cross-module names
    private builtinEnv!: Environment; // Tier 4: standard library
    private globalEnv!: Environment;  // Stable Tier 3 parent; not replaced during calls
    private moduleEnvs: Map<string, Environment> = new Map(); // Tier 2: per-module vars
    private fileHandles: Map<number, {
        fd: number,
        fileNumber: number,
        mode: 'Input' | 'Output' | 'Append' | 'Random' | 'Binary',
        path: string,
        access: 'Read' | 'Write' | 'Read Write',
        lock: 'Shared' | 'Lock Read' | 'Lock Write' | 'Lock Read Write',
        buffer?: Uint8Array,
        pos?: number,
        recordLen?: number,
        lastRecord?: number,
        nextRecord?: number,
        eof?: boolean,
        width?: number,
        lineColumn?: number,
        locks?: Array<{ start: number, end: number, key: string }>
    }> = new Map();
    /** Runtime Lock ranges shared by all handles in this evaluator. */
    private fileLocks: Map<string, Array<{ owner: number, start: number, end: number, key: string }>> = new Map();
    private sandbox: SandboxPath;
    public fs: FileSystem;
    private onPrint: PrintCallback;
    private currentProcBody: Statement[] | null = null;
    private currentProcedureName: string | null = null;
    private currentProcedureType: string | null = null;
    private currentProcedureReturnType: string | undefined;
    private currentProcedureReturnsArray = false;
    private currentSourceModule: string = '';
    private executingModuleName: string = '';
    // Maps module name (lower) -> set of variable/const names (lower) declared at module level
    private moduleVarRegistry: Map<string, Set<string>> = new Map();
    private withObjectStack: any[] = [];
    // Non-zero while executing a GoSub target in the current procedure.  A
    // Return propagates to that target executor; a new procedure resets this
    // value so a callee cannot consume its caller's GoSub continuation.
    private gosubTargetDepth = 0;
    private staticVarStore: Map<string, any> = new Map(); // persistent store for Static variables
    private currentProcIsStatic: boolean = false;
    private arrayBase: number = 0;
    /** §5.2.2 Def-Directive: letter (a-z) → VBA type name */
    private defTypeMap: Map<string, string> = new Map();
    /** VarPtr/StrPtr/ObjPtr: dummy address counter (non-zero unique integers) */
    private _ptrCounter: number = 0x10000;
    private staticVarsInCurrentProc: Set<string> = new Set();
    private errObj: VbaErrObject = new VbaErrObject();
    /** 最後に実行した数値行ラベル（Erl 用）。数値ラベルの LabelStatement 通過時に更新 */
    private lastLineNumberLabel: number = 0;
    private classDefinitions: Map<string, ClassDeclaration> = new Map();
    /** §5.6.10 Tier 6: 修飾なし識別子の最終フォールバック先オブジェクト（MockApplication 等） */
    private defaultBindingObject: any = null;
    private comparisonMode: 'Binary' | 'Text' = 'Binary';
    private readonly moduleComparisonModes = new Map<string, 'Binary' | 'Text'>();
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

    /** Keep Err object mutation in one path for Resume and runtime failures. */
    private clearErrState(): void {
        this.errObj.clear();
    }

    private recordErrState(error: any): void {
        const number = error?.type === 'VbaError' ? error.number : 1000;
        const description = error?.vbaBareMessage ?? error?.message ?? String(error);
        this.errObj.number = number;
        this.errObj.source = error?.vbaSource ?? (this.executingModuleName || 'VBAProject');
        this.errObj.description = description;
        this.errObj.erl = this.lastLineNumberLabel;
    }

    /** Apply VBA argument passing semantics at the single binding boundary. */
    private prepareArgumentValue(value: any, param: ProcedureDeclaration['parameters'][number], forcedByVal = false): any {
        return param.isByVal || forcedByVal ? this.deepCopyByValValue(value) : value;
    }

    /** Array-designated parameters require an array container, not a scalar. */
    private validateArrayParameterContainer(
        value: any,
        param: ProcedureDeclaration['parameters'][number],
    ): void {
        if (!param.isArray || value === vbaMissing) return;
        if (!Array.isArray(value)) {
            this.throwVbaError(VbaErrorCode.TYPE_MISMATCH, 'Type mismatch');
        }
        const expected = param.paramType?.toLowerCase();
        if (!expected || expected === 'variant') return;

        // Typed arrays and array-returning calls carry their element type as
        // runtime metadata. A container check alone is insufficient for
        // ByRef/Property Let array parameters: Integer() and Array()'s
        // Variant() must not enter a Long() binder.
        const actual = ((value as any).__vbaElementType__ ??
            (value as any).__vbaElementTypeName__ ??
            (value as any).__vbaElementObjectTypeName__ ??
            (value as any).__vbaArrayReturnType__)?.toLowerCase?.();
        if (actual && actual !== expected) {
            this.throwVbaError(VbaErrorCode.TYPE_MISMATCH, 'Type mismatch');
        }
    }

    /** Validate a Set assignment to a typed object-array element. */
    private validateObjectArrayElementValue(value: any, requiredType: string): void {
        if (requiredType.toLowerCase() === 'object') {
            if (!isVbaObjectReferenceCompatible(value)) {
                this.throwVbaError(VbaErrorCode.OBJECT_REQUIRED, 'Object required');
            }
            return;
        }
        const actualType = value?.__className__ as string | undefined;
        if (value !== vbaNothing &&
            (!actualType || actualType.toLowerCase() !== requiredType.toLowerCase())) {
            this.throwVbaError(VbaErrorCode.TYPE_MISMATCH, 'Type mismatch');
        }
    }

    private argumentReference(
        expr: Expression | undefined,
        param: ProcedureDeclaration['parameters'][number],
        forcedByVal = false,
    ): VbaLValueReference | null {
        return expr && !param.isByVal && !forcedByVal ? this.createLValueReference(expr) : null;
    }

    /** Bind positional values to a procedure frame for module and class calls. */
    private bindProcedureParameters(
        proc: ProcedureDeclaration,
        args: any[],
        localEnv: Environment,
        argSubtypes?: (VbaVarType | undefined)[],
        validateObjectArguments = false,
    ): string | null {
        for (let i = 0; i < proc.parameters.length; i++) {
            const param = proc.parameters[i];
            const paramName = param.name;
            if (param.isParamArray) {
                const remainingArgs = args.slice(i);
                (remainingArgs as any).vbaBase = 0;
                localEnv.setLocally(paramName, remainingArgs);
                return paramName;
            }

            let argValue: any;
            if (i < args.length && args[i] !== vbaOmitted) argValue = args[i];
            else if (param.defaultValue) argValue = this.evaluateExpression(param.defaultValue);
            else argValue = vbaMissing;

            this.validateArrayParameterContainer(argValue, param);

            if (param.paramType && !param.isArray) {
                const typeMap: Record<string, VbaVarType> = {
                    'byte': 'Byte', 'integer': 'Integer', 'long': 'Long',
                    'single': 'Single', 'double': 'Double', 'currency': 'Currency',
                    'string': 'String', 'boolean': 'Boolean', 'date': 'Date',
                };
                const mapped = typeMap[param.paramType.toLowerCase()];
                if (mapped) {
                    localEnv.setVariableType(paramName, {
                        vbaType: mapped,
                        fixedLength: mapped === 'String' ? param.fixedLength : undefined,
                    });
                } else if (param.paramType.toLowerCase() === 'object') {
                    localEnv.setVariableType(paramName, { vbaType: 'Object' });
                } else if (this.classDefinitions.has(param.paramType.toLowerCase())) {
                    localEnv.setVariableType(paramName, {
                        vbaType: 'Object',
                        objectTypeName: param.paramType,
                    });
                }
            }
            if (validateObjectArguments && param.paramType?.toLowerCase() === 'object' &&
                !param.isArray && argValue !== vbaMissing &&
                !isVbaObjectReferenceCompatible(argValue)) {
                this.throwVbaError(VbaErrorCode.OBJECT_REQUIRED, 'Object required');
            }
            argValue = this.prepareArgumentValue(argValue, param);
            localEnv.setLocally(paramName, argValue);
            if ((!param.paramType || param.paramType.toLowerCase() === 'variant')
                    && argSubtypes && i < argSubtypes.length && argSubtypes[i] && typeof argValue === 'number') {
                localEnv.setVariantSubtype(paramName, argSubtypes[i]!);
            }
            this.addRef(argValue);
        }
        return null;
    }

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
        this.globalEnv = this.env;
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
        const ctx = this.makeStdlibCtx();
        registerInformationFunctions(ctx);
        registerConversionFunctions(ctx);
        registerMathFunctions(ctx);
        registerStringFunctions(ctx);
        this.registerDateTimeFunctions();
        registerStdlibDateTimeFunctions(ctx);
        this.registerFileSystemFunctions();
        registerInteractionFunctions(ctx);
        registerFinancialFunctions(ctx);
        registerConstants(ctx);
        registerRegistryFunctions(ctx);
    }

    private makeStdlibCtx(): StdlibCtx {
        const self = this;
        return {
            reg: (n, f, p, v) => this.registerBuiltin(n, f, p, v),
            regOvl: (n, f, o, v) => this.registerOverloadedBuiltin(n, f, o, v),
            envSet: (n, v, vars) => this.envSet(n, v, vars),
            envSetConst: (n, v) => this.env.setConstant(n, v),
            envGet: (n) => this.env.get(n),
            toVbaNumber: (v) => this.toVbaNumber(v),
            throwError: (c, m) => this.throwVbaError(c, m),
            isTrue: (v) => this.isTrue(v),
            round: (n, d = 0) => this.vbaRound(n, d),
            callMethod: (o, p, a) => this.callClassMethod(o, p, a),
            get compMode() { return self.getComparisonMode(); },
            get arrayBase() { return self.arrayBase; },
            print: (m) => this.onPrint(m),
            errNum: () => this.errObj.number,
            getEnv: (k) => this.sandbox.getEnv(k),
            ptrNext: () => (this._ptrCounter += 4),
            createObj: (id) => this.createExternalObject(id),
            registry: this.virtualRegistry,
        };
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
        this.registerBuiltin('erl', () => this.errObj.erl, []);
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

    private registerFileSystemFunctions() {
        const freeFileFunc = (range?: any) => {
            if (range === vbaNull) {
                this.throwVbaError(VbaErrorCode.INVALID_USE_OF_NULL, 'Invalid use of Null');
            }
            const rangeNumber = range === undefined ? 0 : this.toVbaNumber(range);
            if (!Number.isInteger(rangeNumber) || (rangeNumber !== 0 && rangeNumber !== 1)) {
                this.throwVbaError(VbaErrorCode.INVALID_PROCEDURE_CALL, 'Invalid procedure call or argument');
            }
            const start = rangeNumber === 1 ? 256 : 1;
            const end = rangeNumber === 1 ? 511 : 255;
            for (let i = start; i <= end; i++) if (!this.fileHandles.has(i)) return i;
            this.throwVbaError(VbaErrorCode.TOO_MANY_FILES, "Too many files");
        };
        this.registerBuiltin('freefile', freeFileFunc, [{ name: 'RangeNumber', optional: true }]);
        this.registerBuiltin('eof', (fn: any) => {
            const h = this.fileHandles.get(this.toFileNumber(fn));
            if (!h) this.throwVbaError(VbaErrorCode.BAD_FILE_NAME_OR_NUMBER, "Bad file name or number");
            if (h.mode === 'Output' || h.mode === 'Append') return vbaTrue;
            if (h.mode === 'Random' || h.mode === 'Binary') return h.eof ? vbaTrue : vbaFalse;
            return h.pos! >= this.fs.statSync(h.path).size ? vbaTrue : vbaFalse;
        }, [{ name: 'FileNumber' }]);
        this.registerBuiltin('lof', (fn: any) => {
            const h = this.fileHandles.get(this.toFileNumber(fn));
            if (!h) this.throwVbaError(VbaErrorCode.BAD_FILE_NAME_OR_NUMBER, "Bad file name or number");
            return this.fs.statSync(h.path).size;
        }, [{ name: 'FileNumber' }]);
        this.registerBuiltin('loc', (fn: any) => {
            const h = this.fileHandles.get(this.toFileNumber(fn));
            if (!h) this.throwVbaError(VbaErrorCode.BAD_FILE_NAME_OR_NUMBER, "Bad file name or number");
            if (h.mode === 'Random') return h.lastRecord ?? 0;
            if (h.mode === 'Input' || h.mode === 'Output' || h.mode === 'Append') {
                return Math.floor((h.pos ?? 0) / 128);
            }
            return h.pos;
        }, [{ name: 'FileNumber' }]);
        this.registerBuiltin('seek', (fn: any) => {
            const h = this.fileHandles.get(this.toFileNumber(fn));
            if (!h) this.throwVbaError(VbaErrorCode.BAD_FILE_NAME_OR_NUMBER, "Bad file name or number");
            if (h.mode === 'Random') return h.nextRecord ?? 1;
            return (h.pos || 0) + 1;
        }, [{ name: 'FileNumber' }]);
        this.registerBuiltin('fileattr', (fn: any, info: any = 1) => {
            if (fn === vbaNull || info === vbaNull) {
                this.throwVbaError(VbaErrorCode.INVALID_USE_OF_NULL, 'Invalid use of Null');
            }
            const fileNumber = this.toVbaNumber(fn);
            const returnType = this.toVbaNumber(info);
            const handle = this.fileHandles.get(fileNumber);
            if (!handle) this.throwVbaError(VbaErrorCode.BAD_FILE_NAME_OR_NUMBER, 'Bad file name or number');
            if (returnType === 1) return handle.fd;
            if (returnType === 2) return fileNumber;
            this.throwVbaError(VbaErrorCode.INVALID_PROCEDURE_CALL, 'Invalid procedure call or argument');
        }, [{ name: 'FileNumber' }, { name: 'ReturnType', optional: true }]);
        this.registerBuiltin('chdrive', (drive: any) => {
            console.log(`[STUB] ChDrive "${vbaToString(drive ?? '')}"`);
        }, [{ name: 'Drive' }]);
        this.registerBuiltin('filedatetime', (path: any) => {
            const realPath = this.sandbox.toRealPath(vbaToString(path ?? ''));
            try {
                return new VbaDate(toVbaDate(this.fs.statSync(realPath).mtime));
            } catch {
                this.throwVbaError(VbaErrorCode.FILE_NOT_FOUND, 'File not found');
            }
        }, [{ name: 'PathName' }]);
        this.registerBuiltin('curdir', (drive?: any) => {
            if (drive === vbaNull) {
                this.throwVbaError(VbaErrorCode.INVALID_USE_OF_NULL, 'Invalid use of Null');
            }
            return this.sandbox.getCwd();
        }, [{ name: 'Drive', optional: true }], ['$']);
        this.registerBuiltin('dir', (pathName?: any, attributes?: any) => {
            if (pathName === vbaNull || attributes === vbaNull) {
                this.throwVbaError(VbaErrorCode.INVALID_USE_OF_NULL, 'Invalid use of Null');
            }
            const attrMask = attributes === undefined ? VBA_FILE_ATTRIBUTE.NORMAL : this.toVbaNumber(attributes);
            if (!Number.isInteger(attrMask) || attrMask < 0 || attrMask > 127) {
                this.throwVbaError(VbaErrorCode.INVALID_PROCEDURE_CALL, 'Invalid procedure call or argument');
            }
            if (pathName !== undefined && pathName !== null && pathName !== "") {
                try {
                    const realPath = this.sandbox.toRealPath(pathName);
                    const dirPath = path.dirname(realPath);
                    const filter = path.basename(realPath);
                    const files = this.fs.readdirSync(dirPath);
                    const includeDirectories = (attrMask & VBA_FILE_ATTRIBUTE.DIRECTORY) !== 0;
                    const requestedAttributes = attrMask & ~VBA_FILE_ATTRIBUTE.DIRECTORY;
                    if (filter === '' || filter === '*' || filter === '*.*') {
                        this.dirIterator = files.filter((entry: string) => {
                            const entryPath = path.join(dirPath, entry);
                            const stat = this.fs.statSync(entryPath);
                            if (stat.isDirectory() && !includeDirectories) return false;
                            if (requestedAttributes === 0) return true;
                            const actual = this.fs.getAttributes
                                ? this.fs.getAttributes(entryPath)
                                : (stat.isDirectory() ? VBA_FILE_ATTRIBUTE.DIRECTORY : VBA_FILE_ATTRIBUTE.NORMAL);
                            return (actual & requestedAttributes) === requestedAttributes;
                        });
                    } else {
                        const regex = this.vbaWildcardToRegex(filter);
                        this.dirIterator = files.filter((entry: string) => {
                            if (!regex.test(entry)) return false;
                            const entryPath = path.join(dirPath, entry);
                            const stat = this.fs.statSync(entryPath);
                            if (stat.isDirectory() && !includeDirectories) return false;
                            if (requestedAttributes === 0) return true;
                            const actual = this.fs.getAttributes
                                ? this.fs.getAttributes(entryPath)
                                : (stat.isDirectory() ? VBA_FILE_ATTRIBUTE.DIRECTORY : VBA_FILE_ATTRIBUTE.NORMAL);
                            return (actual & requestedAttributes) === requestedAttributes;
                        });
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
        this.registerBuiltin('filecopy', (src: any, dest: any) => {
            const srcPath = this.sandbox.toRealPath(vbaToString(src ?? ''));
            const destPath = this.sandbox.toRealPath(vbaToString(dest ?? ''));
            try {
                if (this.fs.existsSync(destPath)) {
                    this.throwVbaError(VbaErrorCode.FILE_ALREADY_EXISTS, 'File already exists');
                }
                this.fs.copyFileSync?.(srcPath, destPath);
            } catch {
                // Preserve the VBA error raised for an existing destination.
                if (this.fs.existsSync(destPath)) {
                    this.throwVbaError(VbaErrorCode.FILE_ALREADY_EXISTS, 'File already exists');
                }
                this.throwVbaError(VbaErrorCode.FILE_NOT_FOUND, 'File not found');
            }
        }, [
            { name: 'Source' }, { name: 'Destination' },
        ]);
        this.registerBuiltin('kill', (p: any) => this.executeKill(vbaToString(p ?? '')), [{ name: 'PathName' }]);
        this.registerBuiltin('mkdir', (p: any) => {
            const realPath = this.sandbox.toRealPath(vbaToString(p ?? ''));
            this.validateSingleDirectoryCreation(realPath, VbaErrorCode.PATH_FILE_ACCESS_ERROR,
                'Path/File access error', true);
            this.fs.mkdirSync(realPath);
        }, [{ name: 'Path' }]);
        this.registerBuiltin('rmdir', (p: any) => {
            if (p === vbaNull) {
                this.throwVbaError(VbaErrorCode.INVALID_USE_OF_NULL, 'Invalid use of Null');
            }
            const realPath = this.sandbox.toRealPath(vbaToString(p ?? ''));
            let stat: ReturnType<FileSystem['statSync']>;
            try {
                stat = this.fs.statSync(realPath);
            } catch {
                this.throwVbaError(VbaErrorCode.PATH_NOT_FOUND, 'Path not found');
            }
            if (!stat!.isDirectory()) {
                this.throwVbaError(VbaErrorCode.PATH_FILE_ACCESS_ERROR, 'Path/File access error');
            }
            if (this.fs.readdirSync(realPath).length > 0) {
                this.throwVbaError(VbaErrorCode.PATH_FILE_ACCESS_ERROR, 'Path/File access error');
            }
            try {
                this.fs.rmdirSync?.(realPath);
            } catch {
                this.throwVbaError(VbaErrorCode.PATH_FILE_ACCESS_ERROR, 'Path/File access error');
            }
        }, [{ name: 'Path' }]);
        this.registerBuiltin('chdir', (p: any) => {
            if (p === vbaNull) {
                this.throwVbaError(VbaErrorCode.INVALID_USE_OF_NULL, 'Invalid use of Null');
            }
            try {
                this.sandbox.setCwd(vbaToString(p ?? ''));
            } catch {
                this.throwVbaError(VbaErrorCode.PATH_NOT_FOUND, 'Path not found');
            }
        }, [{ name: 'Path' }]);
        this.registerBuiltin('filelen', (p: any) => {
            const realPath = this.sandbox.toRealPath(vbaToString(p ?? ''));
            try {
                return this.fs.statSync(realPath).size;
            } catch {
                this.throwVbaError(VbaErrorCode.FILE_NOT_FOUND, 'File not found');
            }
        }, [{ name: 'PathName' }]);
        this.registerBuiltin('getattr', (p: any) => {
            const realPath = this.sandbox.toRealPath(vbaToString(p ?? ''));
            try {
                if (this.fs.getAttributes) return this.fs.getAttributes(realPath);
                const stat = this.fs.statSync(realPath);
                let attributes = stat.isDirectory() ? VBA_FILE_ATTRIBUTE.DIRECTORY : VBA_FILE_ATTRIBUTE.NORMAL;
                if (stat.mode !== undefined && (stat.mode & 0o222) === 0) {
                    attributes |= VBA_FILE_ATTRIBUTE.READ_ONLY;
                }
                return attributes;
            } catch {
                this.throwVbaError(VbaErrorCode.FILE_NOT_FOUND, 'File not found');
            }
        }, [{ name: 'PathName' }]);
        this.registerBuiltin('setattr', (p: any, attr: any) => {
            const realPath = this.sandbox.toRealPath(vbaToString(p ?? ''));
            const attributes = this.toVbaNumber(attr);
            try {
                if (!this.fs.setAttributes) {
                    this.throwVbaError(VbaErrorCode.INVALID_PROCEDURE_CALL, 'File attributes are not supported');
                }
                if ((attributes & VBA_FILE_ATTRIBUTE.DIRECTORY) !== 0) {
                    this.throwVbaError(VbaErrorCode.INVALID_PROCEDURE_CALL, 'Directory attribute cannot be set');
                }
                this.fs.setAttributes(realPath, attributes);
            } catch (e: any) {
                if (e?.type === 'VbaError') throw e;
                if (e?.message?.startsWith('Invalid file attributes')) {
                    this.throwVbaError(VbaErrorCode.INVALID_PROCEDURE_CALL, e.message);
                }
                this.throwVbaError(VbaErrorCode.FILE_NOT_FOUND, 'File not found');
            }
        }, [{ name: 'PathName' }, { name: 'Attributes' }]);
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

    private addRef(obj: any): void {
        if (obj && obj.__vbaClass__) {
            obj.__refCount__ = (obj.__refCount__ ?? 0) + 1;
        }
    }

    // Decrement refcount; call triggerTerminate when count reaches 0.
    private releaseRef(obj: any): void {
        if (obj && obj.__vbaClass__) {
            obj.__refCount__ = (obj.__refCount__ ?? 0) - 1;
            if (obj.__refCount__ <= 0) {
                this.triggerTerminate(obj);
            }
        }
    }

    private instantiateType(typeName: string): any {
        const typeMembers = this.env.getType(typeName);
        if (!typeMembers) return 0;

        const instance: any = {
            __vbaTypeName__: typeName,
            __vbaUdt__: true
        };
        for (const member of typeMembers) {
            const mt = member.memberType;
            const mtLower = mt.toLowerCase();
            const memberKey = member.name.toLowerCase();

            if (member.isArray) {
                if (member.arrayBounds && member.arrayBounds.length > 0) {
                    const dims = member.arrayBounds;
                    const buildArr = (dimIdx: number): any[] => {
                        const { lower, upper } = dims[dimIdx];
                        const a = new Array(upper + 1);
                        if (dimIdx < dims.length - 1) {
                            for (let i = lower; i <= upper; i++) a[i] = buildArr(dimIdx + 1);
                        } else if (this.env.getType(mt)) {
                            for (let i = lower; i <= upper; i++) a[i] = this.instantiateType(mt);
                        } else {
                            const def = this.makeUdtMemberDefault(mt, mtLower, member.fixedLength);
                            for (let i = lower; i <= upper; i++) a[i] = def;
                        }
                        return a;
                    };
                    const arr = buildArr(0);
                    (arr as any).__vbaDimensions__ = dims.map(b => ({ lower: b.lower, upper: b.upper }));
                    (arr as any).vbaFixed = true;
                    (arr as any).vbaBase = dims[0].lower;
                    if (this.env.getType(mt)) (arr as any).__vbaElementTypeName__ = mt;
                    if (mtLower === 'string' && member.fixedLength !== undefined) {
                        (arr as any).__vbaElementFixedLength__ = member.fixedLength;
                    }
                    instance[memberKey] = arr;
                } else {
                    const arr: any[] = [];
                    (arr as any).vbaBase = this.arrayBase;
                    (arr as any).vbaFixed = false;
                    if (this.env.getType(mt)) (arr as any).__vbaElementTypeName__ = mt;
                    instance[memberKey] = arr;
                }
            } else if (mtLower === 'string') {
                instance[memberKey] = member.fixedLength !== undefined
                    ? '\0'.repeat(member.fixedLength)
                    : '';
                if (member.fixedLength !== undefined) {
                    // Tag instance so UDT member assignment can apply fixed-length coercion
                    if (!instance.__fixedLengths__) instance.__fixedLengths__ = {};
                    instance.__fixedLengths__[memberKey] = member.fixedLength;
                }
            } else if (mtLower === 'boolean') {
                instance[memberKey] = 0; // vbaFalse
            } else if (this.env.getType(mt)) {
                instance[memberKey] = this.instantiateType(mt);
            } else {
                instance[memberKey] = 0;
            }
        }
        return instance;
    }

    private makeUdtMemberDefault(mt: string, mtLower: string, fixedLength?: number): any {
        if (mtLower === 'string') return fixedLength !== undefined ? '\0'.repeat(fixedLength) : '';
        if (mtLower === 'boolean') return 0;
        if (this.env.getType(mt)) return this.instantiateType(mt);
        return 0;
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

    public callProcedure(name: string, args: any[], type?: 'get' | 'let' | 'set', moduleName?: string, argSubtypes?: (VbaVarType | undefined)[]): any {
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

        this.bindProcedureParameters(proc, args, localEnv, argSubtypes);

        const result = this.execProcBody(proc, localEnv, {
            byRefArgs: [],
            paramArrayParamName: null,
            paramArrayReferences: [],
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
            this.moduleEnvs.set(key, new Environment(this.globalEnv));
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
            const findings = this.collectPrecheckFindings(proc);
            if (findings.subAsValue) {
                this.throwCompileError(VbaErrorCode.TYPE_MISMATCH,
                    `Function or variable expected: '${findings.subAsValue.name}'`,
                    findings.subAsValue.line, proc.moduleName ?? undefined);
            }
            if (findings.undefinedCalls.length > 0) {
                const error = findings.undefinedCalls[0];
                this.throwCompileError(VbaErrorCode.SUB_OR_FUNCTION_NOT_DEFINED,
                    `Sub or Function not defined: '${error.name}'`,
                    error.line || undefined, proc.moduleName ?? undefined);
            }
            for (const bound of findings.arrayBounds) {
                this.validateConstantExpr(bound.upper);
                if (bound.lower) this.validateConstantExpr(bound.lower);
            }
            if (findings.duplicate) {
                this.throwCompileError(VbaErrorCode.INVALID_PROCEDURE_CALL,
                    `${findings.duplicate.kind} '${findings.duplicate.name}' is already declared in this scope (duplicate declaration)`,
                    findings.duplicate.line, proc.moduleName ?? undefined);
            }
            for (const jump of findings.jumps) {
                if (!findings.labels.has(jump.label.toLowerCase())) {
                    this.throwCompileError(VbaErrorCode.SUB_OR_FUNCTION_NOT_DEFINED,
                        `Label not defined: '${jump.label}'`, jump.line || undefined,
                        proc.moduleName ?? undefined);
                }
            }
            if (findings.argumentError) {
                this.throwCompileError(findings.argumentError.code,
                    findings.argumentError.message, findings.argumentError.line,
                    proc.moduleName ?? undefined);
            }
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

    /**
     * Collect all procedure-level static findings in one depth-first AST walk.
     * The findings are reported by precheckProc in the historical checker order
     * so combining the traversal does not change which error wins.
     */
    private collectPrecheckFindings(proc: ProcedureDeclaration): {
        subAsValue?: { name: string; line?: number };
        undefinedCalls: UndefinedProcError[];
        arrayBounds: ArrayBound[];
        duplicate?: { kind: 'Variable' | 'Constant'; name: string; line?: number };
        labels: Set<string>;
        jumps: Array<{ label: string; line: number }>;
        argumentError?: { code: number; message: string; line?: number };
    } {
        const findings = {
            undefinedCalls: [] as UndefinedProcError[],
            arrayBounds: [] as ArrayBound[],
            labels: new Set<string>(),
            jumps: [] as Array<{ label: string; line: number }>,
        } as {
            subAsValue?: { name: string; line?: number };
            undefinedCalls: UndefinedProcError[];
            arrayBounds: ArrayBound[];
            duplicate?: { kind: 'Variable' | 'Constant'; name: string; line?: number };
            labels: Set<string>;
            jumps: Array<{ label: string; line: number }>;
            argumentError?: { code: number; message: string; line?: number };
        };

        const knownNames = this.env.collectAllNames();
        for (const mEnv of this.moduleEnvs.values()) {
            for (const name of mEnv.collectAllNames()) knownNames.add(name);
        }
        if (this.defaultBindingObject) {
            let cur = this.defaultBindingObject;
            while (cur && cur !== Object.prototype && cur !== Function.prototype) {
                for (const name of Object.getOwnPropertyNames(cur)) knownNames.add(name.toLowerCase());
                cur = Object.getPrototypeOf(cur);
            }
        }

        const declared = new Set<string>([proc.name.name.toLowerCase()]);
        for (const param of proc.parameters) declared.add(param.name.toLowerCase());
        let inResumeNext = false;
        const seen = new Set<string>(declared);
        if (proc.isFunction || proc.isProperty) seen.add(proc.name.name.toLowerCase());

        const visitExpression = (expr: Expression, assignmentTarget = false): void => {
            if (expr.type === 'CallExpression') {
                const call = expr as CallExpression;
                if (!(assignmentTarget && call.callee.type === 'Identifier')) {
                    if (call.callee.type === 'Identifier') {
                        const id = call.callee as Identifier;
                        const lower = id.name.toLowerCase();
                        if (!inResumeNext && !id.foreign && !declared.has(lower) && !knownNames.has(lower)) {
                            findings.undefinedCalls.push({ name: id.name, line: id.loc?.start.line ?? 0 });
                        }
                    }
                    if (!findings.argumentError && call.callee.type === 'Identifier') {
                        const target = this.env.getProcedure((call.callee as Identifier).name);
                        if (target && !target.parameters.some(p => p.isParamArray)) {
                            const min = target.parameters.filter(p => !p.isOptional && p.defaultValue == null).length;
                            if (call.args.length > target.parameters.length) {
                                findings.argumentError = {
                                    code: VbaErrorCode.WRONG_NUMBER_OF_ARGUMENTS,
                                    message: 'Wrong number of arguments or invalid property assignment',
                                    line: call.loc?.start.line ?? call.callee.loc?.start.line,
                                };
                            } else if (call.args.length < min) {
                                findings.argumentError = {
                                    code: VbaErrorCode.ARGUMENT_NOT_OPTIONAL,
                                    message: 'Argument not optional',
                                    line: call.loc?.start.line ?? call.callee.loc?.start.line,
                                };
                            }
                        }
                    }
                }
            }

            if (!inResumeNext || expr.type !== 'CallExpression') {
                switch (expr.type) {
                    case 'CallExpression': {
                        const call = expr as CallExpression;
                        visitExpression(call.callee);
                        for (const arg of call.args) visitExpression(arg);
                        break;
                    }
                    case 'NamedArgument':
                        visitExpression((expr as NamedArgument).value); break;
                    case 'ByValArgument':
                        visitExpression((expr as ByValArgument).value); break;
                    case 'MemberExpression':
                    case 'DictionaryAccessExpression':
                        visitExpression((expr as MemberExpression).object, assignmentTarget); break;
                    case 'TypeOfIsExpression':
                        visitExpression((expr as TypeOfIsExpression).expression); break;
                    case 'BinaryExpression': {
                        const b = expr as BinaryExpression;
                        visitExpression(b.left); visitExpression(b.right); break;
                    }
                    case 'UnaryExpression':
                        visitExpression((expr as UnaryExpression).argument); break;
                    case 'ParenthesizedExpression':
                        visitExpression((expr as ParenthesizedExpression).expression); break;
                }
            }
        };

        const visitStatement = (stmt: Statement): void => {
            switch (stmt.type) {
                case 'VariableDeclaration': {
                    const decl = stmt as VariableDeclaration;
                    for (const d of decl.declarations) {
                        const key = d.name.name.toLowerCase();
                        if (d.isArray && d.arrayBounds) findings.arrayBounds.push(...d.arrayBounds);
                        if (!findings.duplicate && seen.has(key)) {
                            findings.duplicate = { kind: 'Variable', name: d.name.name, line: d.name.loc?.start.line };
                        }
                        seen.add(key); declared.add(key);
                    }
                    break;
                }
                case 'ConstDeclaration': {
                    const decl = stmt as ConstDeclaration;
                    for (const d of decl.declarations) {
                        const key = d.name.name.toLowerCase();
                        if (!findings.duplicate && seen.has(key)) {
                            findings.duplicate = { kind: 'Constant', name: d.name.name, line: d.name.loc?.start.line };
                        }
                        seen.add(key); declared.add(key);
                    }
                    break;
                }
                case 'OnErrorStatement':
                    inResumeNext = (stmt as any).label === 'Resume Next'; break;
                case 'LabelStatement':
                    findings.labels.add(((stmt as any).label as string).toLowerCase()); break;
                case 'GoToStatement':
                case 'GoSubStatement':
                    findings.jumps.push({ label: (stmt as any).label, line: stmt.loc?.start.line ?? 0 }); break;
                case 'OnGoToSubStatement':
                    for (const label of (stmt as OnGoToSubStatement).labels) {
                        findings.jumps.push({ label, line: stmt.loc?.start.line ?? 0 });
                    }
                    break;
                case 'AssignmentStatement': {
                    const a = stmt as AssignmentStatement;
                    const name = this.subNameInValueExpr(a.right);
                    if (!findings.subAsValue && name) {
                        const target = this.env.getProcedure(name);
                        if (target && !target.isFunction && !target.isProperty) {
                            findings.subAsValue = { name, line: a.right.loc?.start.line };
                        }
                    }
                    visitExpression(a.left, true); visitExpression(a.right); break;
                }
                case 'SetStatement': {
                    const s = stmt as SetStatement;
                    visitExpression(s.left, true); visitExpression(s.right); break;
                }
                case 'CallStatement': visitExpression((stmt as CallStatement).expression); break;
                case 'IfStatement': {
                    const s = stmt as IfStatement;
                    visitExpression(s.condition); this.walkPrecheckStatements(s.consequent, visitStatement);
                    if (Array.isArray(s.alternate)) this.walkPrecheckStatements(s.alternate, visitStatement);
                    else if (s.alternate) visitStatement(s.alternate);
                    break;
                }
                case 'ForStatement': {
                    const s = stmt as ForStatement;
                    visitExpression(s.start); visitExpression(s.end); if (s.step) visitExpression(s.step);
                    this.walkPrecheckStatements(s.body, visitStatement); break;
                }
                case 'ForEachStatement': {
                    const s = stmt as ForEachStatement;
                    visitExpression(s.collection); this.walkPrecheckStatements(s.body, visitStatement); break;
                }
                case 'DoWhileStatement': {
                    const s = stmt as DoWhileStatement;
                    if (s.condition) visitExpression(s.condition);
                    this.walkPrecheckStatements(s.body, visitStatement); break;
                }
                case 'WhileStatement': {
                    const s = stmt as WhileStatement;
                    visitExpression(s.condition); this.walkPrecheckStatements(s.body, visitStatement); break;
                }
                case 'WithStatement': {
                    const s = stmt as WithStatement;
                    visitExpression(s.expression); this.walkPrecheckStatements(s.body, visitStatement); break;
                }
                case 'SelectCaseStatement': {
                    const s = stmt as SelectCaseStatement;
                    visitExpression(s.expression);
                    for (const c of s.cases) {
                        for (const range of c.ranges) {
                            if (range.kind === 'to') { visitExpression(range.start); visitExpression(range.end); }
                            else visitExpression(range.value);
                        }
                        this.walkPrecheckStatements(c.body, visitStatement);
                    }
                    if (s.elseBody) this.walkPrecheckStatements(s.elseBody, visitStatement);
                    break;
                }
                case 'ReDimStatement': {
                    const s = stmt as ReDimStatement;
                    for (const d of s.declarations) for (const b of d.bounds) {
                        if (b.lower) visitExpression(b.lower); visitExpression(b.upper);
                    }
                    break;
                }
                case 'LSetStatement': case 'RSetStatement': {
                    const s = stmt as LSetStatement | RSetStatement;
                    visitExpression(s.left); visitExpression(s.right); break;
                }
                case 'MidStatement': {
                    const s = stmt as MidStatement;
                    visitExpression(s.target); visitExpression(s.start);
                    if (s.length) visitExpression(s.length); visitExpression(s.value); break;
                }
            }
        };

        this.walkPrecheckStatements(proc.body, visitStatement);
        return findings;
    }

    private walkPrecheckStatements(stmts: Statement[], visitor: (stmt: Statement) => void): void {
        for (const stmt of stmts) visitor(stmt);
    }

    /** Reject case-insensitive name collisions among declarations in one module/class scope. */
    private checkDuplicateDeclarationsInScope(stmts: Statement[], scopeName: string): void {
        const seen = new Map<string, { kind: string; line: number }>();
        const add = (name: string, kind: string, line: number): void => {
            const key = name.toLowerCase();
            const previous = seen.get(key);
            if (!previous) {
                seen.set(key, { kind, line });
                return;
            }
            // VBA permits the three Property procedures (Get/Let/Set) to share
            // one name.  They are distinct accessors, not duplicate declarations.
            const propertyOverload = previous.kind.startsWith('procedure:') &&
                kind.startsWith('procedure:') && previous.kind !== kind &&
                previous.kind !== 'procedure:sub/function' && kind !== 'procedure:sub/function';
            if (propertyOverload) return;
            // VBA keeps type and value names in separate namespaces.  A class
            // declaration may therefore share its name with a Function/Sub;
            // calls resolve the procedure while As/New resolve the class.
            const typeValueNamespacePair =
                (previous.kind === 'class' && kind === 'procedure:sub/function') ||
                (previous.kind === 'procedure:sub/function' && kind === 'class');
            if (typeValueNamespacePair) return;
            this.throwCompileError(
                VbaErrorCode.INVALID_PROCEDURE_CALL,
                `Duplicate declaration '${name}' in scope '${scopeName}'`,
                line,
            );
        };

        for (const stmt of stmts) {
            switch (stmt.type) {
                case 'VariableDeclaration':
                    for (const d of (stmt as VariableDeclaration).declarations) {
                        add(d.name.name, 'variable', d.name.loc?.start.line ?? stmt.loc?.start.line ?? 0);
                    }
                    break;
                case 'ConstDeclaration':
                    for (const d of (stmt as ConstDeclaration).declarations) {
                        add(d.name.name, 'constant', d.name.loc?.start.line ?? stmt.loc?.start.line ?? 0);
                    }
                    break;
                case 'ProcedureDeclaration': {
                    const proc = stmt as ProcedureDeclaration;
                    const kind = proc.isProperty && proc.propertyType
                        ? `procedure:${proc.propertyType}`
                        : 'procedure:sub/function';
                    add(proc.name.name, kind, proc.name.loc?.start.line ?? stmt.loc?.start.line ?? 0);
                    break;
                }
                case 'TypeDeclaration':
                    add((stmt as TypeDeclaration).name, 'type', stmt.loc?.start.line ?? 0);
                    break;
                case 'EnumDeclaration': {
                    const enumDecl = stmt as EnumDeclaration;
                    add(enumDecl.name.name, 'enum', enumDecl.name.loc?.start.line ?? stmt.loc?.start.line ?? 0);
                    break;
                }
                case 'ClassDeclaration': {
                    const cls = stmt as ClassDeclaration;
                    add(cls.name, 'class', stmt.loc?.start.line ?? 0);
                    this.checkDuplicateDeclarationsInScope(cls.body, cls.name);
                    break;
                }
            }
        }
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
        if (proc.isFunction || (proc.isProperty && proc.propertyType === 'get')) {
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
                else if (this.classDefinitions.has(proc.returnType.toLowerCase())) {
                    localEnv.setVariableType(proc.name.name, {
                        vbaType: 'Object',
                        objectTypeName: proc.returnType,
                    });
                }
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
        const previousProcedureReturnType = this.currentProcedureReturnType;
        const previousProcedureReturnsArray = this.currentProcedureReturnsArray;
        const previousExecutingModule = this.executingModuleName;
        const previousProcIsStatic = this.currentProcIsStatic;
        const previousStaticVars = this.staticVarsInCurrentProc;
        const previousGoSubTargetDepth = this.gosubTargetDepth;

        this.env = localEnv;
        this.errorHandlerLabel = null;
        this.errorHandlingMode = 'None';
        this.isInErrorHandler = false;
        this.lastErrorIndex = opts.initialLastErrorIndex;
        this.currentProcBody = proc.body;
        this.currentProcedureName = proc.name.name;
        this.currentProcedureType = proc.propertyType || (proc.isFunction ? 'function' : 'sub');
        this.currentProcedureReturnType = proc.returnType;
        this.currentProcedureReturnsArray = proc.returnsArray ?? false;
        this.executingModuleName = proc.moduleName ?? '';
        this.currentProcIsStatic = proc.isStatic ?? false;
        this.staticVarsInCurrentProc = new Set();
        // GoSub/Return continuations belong to one procedure.  Do not allow a
        // Return in a called procedure to consume the caller's continuation.
        this.gosubTargetDepth = 0;

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
                    ref.reference.set(updatedVal);
                } catch {
                    // r-value: silently ignored
                }
            }
            if (opts.paramArrayParamName !== null) {
                const updatedArray = localEnv.get(opts.paramArrayParamName) as any[];
                if (Array.isArray(updatedArray)) {
                    for (let j = 0; j < opts.paramArrayReferences.length; j++) {
                        try {
                            opts.paramArrayReferences[j]?.set(updatedArray[j]);
                        } catch { }
                    }
                }
            }

            for (const varName of this.staticVarsInCurrentProc) {
                const key = `${procNameKey}:${varName}`;
                this.staticVarStore.set(key, localEnv.get(varName));
            }

            // Scope exit: release references held by local variables (including params).
            // The retVar object (retCapture) is released like any local, but Terminate is
            // suppressed for it — even when another local var points to the same object.
            {
                const retVarName = (proc.isFunction || proc.isProperty) ? proc.name.name.toLowerCase() : null;
                const retCapture = retVarName !== null ? localEnv.getLocalVariables().get(retVarName) : undefined;
                for (const [name, val] of localEnv.getLocalVariables()) {
                    if (this.staticVarsInCurrentProc.has(name)) continue;
                    if (val && val.__vbaClass__) {
                        val.__refCount__ = (val.__refCount__ ?? 0) - 1;
                        if (val.__refCount__ <= 0 && val !== retCapture) {
                            this.triggerTerminate(val);
                        }
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
            this.currentProcedureReturnType = previousProcedureReturnType;
            this.currentProcedureReturnsArray = previousProcedureReturnsArray;
            this.executingModuleName = previousExecutingModule;
            this.currentProcIsStatic = previousProcIsStatic;
            this.staticVarsInCurrentProc = previousStaticVars;
            this.gosubTargetDepth = previousGoSubTargetDepth;
        }

        if (proc.isFunction || (opts.returnOnProperty && proc.isProperty)) {
            const result = localEnv.get(proc.name.name);
            if (proc.returnsArray && Array.isArray(result)) (result as any).__vbaArrayReturn__ = true;
            return result;
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
                case 'Identifier': {
                    // VBA keyword constants are not callable procedures
                    const lower = (node as Identifier).name.toLowerCase();
                    if (lower === 'true' || lower === 'false' || lower === 'null' || lower === 'empty' || lower === 'nothing') return false;
                    return true;
                }
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

    /**
     * Returns true if `expr` can appear as the left-hand side of a VBA assignment (`=`).
     * Valid LValues: Identifier (not constant keyword), MemberExpression, CallExpression (subscript/property).
     * Literals, parenthesized expressions, and constant keywords are never assignable.
     */
    private isAssignableTarget(expr: Expression): boolean {
        switch (expr.type) {
            case 'Identifier': {
                const lower = (expr as Identifier).name.toLowerCase();
                return !['true', 'false', 'null', 'empty', 'nothing'].includes(lower);
            }
            case 'MemberExpression':
            case 'CallExpression':
            case 'DictionaryAccessExpression':
                return true;
            default:
                return false;
        }
    }

    /** Convert an assignable VBA expression into the common ByRef write target. */
    private createLValueReference(expr: Expression): VbaLValueReference | null {
        if (!this.isAssignableTarget(expr)) return null;
        return {
            get: () => this.resolveAutoInstance(expr, this.evaluateExpression(expr)),
            // ByRef array parameters may be rebound by ReDim and must write
            // the replacement array back to the caller.
            set: (value: any) => this.evaluateAssignmentToVariable(expr, value, undefined, true),
        };
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
                    // `=` is ambiguous when left side is a valid assignment target:
                    // identifier (not constant), member access, or subscript/call expression.
                    // String/number literals on the left are always comparisons.
                    ((expr as BinaryExpression).operator === '=' && this.isAssignableTarget((expr as BinaryExpression).left)) ||
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
            try {
                this.executeStatements(program.body, 0);
            } catch (e: any) {
                if (!(e && e.type === 'Exit' && (e.target === 'Sub' || e.target === 'Function' || e.target === 'Property'))) {
                    throw e;
                }
            }
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
            case 'DebugPrintStatement':
                this.evaluateDebugPrintStatement(stmt as DebugPrintStatement);
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
            case 'LabelStatement': {
                // 数値行ラベル（レガシー行番号）は Erl のために記録する
                const labelText = (stmt as any).label as string;
                if (/^\d+$/.test(labelText)) this.lastLineNumberLabel = Number(labelText);
                break;
            }
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

        // VBA rejects Null in any numeric For control expression with
        // "Invalid use of Null" (Error 94).  Check before assigning the
        // control variable; Environment.set otherwise obscures the VBA
        // error as an internal type-assignment failure.
        if (startValue === vbaNull || endValue === vbaNull || stepValue === vbaNull) {
            this.throwVbaError(VbaErrorCode.INVALID_USE_OF_NULL, 'Invalid use of Null');
        }

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
        const start = Number((arr as any).vbaBase ?? 0);
        for (let index = start; index < arr.length; index++) {
            const item = arr[index];
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

    /** Convert a file-number argument while preserving VBA's Null error. */
    private toFileNumber(val: any): number {
        if (val === vbaNull) {
            this.throwVbaError(VbaErrorCode.INVALID_USE_OF_NULL, 'Invalid use of Null');
        }
        return this.toVbaNumber(val);
    }

    /** Evaluate a statement's file-number expression with the same Null rule. */
    private evaluateFileNumber(expr: Expression): number {
        return this.toFileNumber(this.evaluateExpression(expr));
    }

    /**
     * Enforce VBA's single-level directory creation contract before calling a
     * backend whose mkdir semantics may differ (for example recursive memory
     * or Node implementations).
     */
    private validateSingleDirectoryCreation(
        realPath: string,
        existingError: number,
        existingMessage: string,
        allowSandboxRootParent: boolean,
    ): void {
        if (this.fs.existsSync(realPath)) {
            this.throwVbaError(existingError, existingMessage);
        }
        const parentPath = path.dirname(realPath);
        if (this.fs.existsSync(parentPath)) return;
        if (allowSandboxRootParent && parentPath === this.sandbox.getRoot()) {
            this.fs.mkdirSync(parentPath, { recursive: true });
            return;
        }
        this.throwVbaError(VbaErrorCode.PATH_NOT_FOUND, 'Path not found');
    }

    private toVbaString(val: any): string {
        try {
            return vbaToString(val);
        } catch (e: any) {
            if (e?.type === 'VbaError' && !(e instanceof Error)) this.throwVbaError(e.number, e.message);
            throw e;
        }
    }

    private vbaRound(val: number, decimals: number = 0): number { return _vbaRound(val, decimals); }

    private throwVbaError(number: number, message: string, overrideLine?: number, overrideModule?: string, source?: string): never {
        const line = overrideLine ?? (this.currentLine || undefined);
        const mod = overrideModule ?? (this.executingModuleName || this.currentSourceModule || null);
        const msg = line !== undefined ? `Run-time error '${number}': ${message} (line ${line})` : `Run-time error '${number}': ${message}`;
        const err: any = new Error(msg);
        err.type = 'VbaError';
        err.number = number;
        err.vbaLine = line;
        err.vbaModule = mod;
        if (source !== undefined) err.vbaSource = source;
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
        // 単一行 If（If x Then a / If x Then a Else b）の節は If 文と同じ行にある。
        // この場合、節内のエラーは If 文の粒度で扱う（Resume が If 全体を再実行できるように）
        const isInline = (clause: Statement[]) =>
            clause.length > 0 && (clause[0] as any).line === (stmt as any).line;
        if (this.isTrue(conditionVal)) {
            this.executeStatements(stmt.consequent, 0, false, isInline(stmt.consequent));
        } else if (stmt.alternate) {
            if (Array.isArray(stmt.alternate)) {
                const alt = stmt.alternate as Statement[];
                this.executeStatements(alt, 0, false, isInline(alt));
            } else {
                this.evaluateIfStatement(stmt.alternate as IfStatement);
            }
        }
    }

    /** Compare VBA values using the same coercion rules as Select Case. */
    private compareVbaValues(rawLeft: any, rawRight: any): number | undefined {
        let left = rawLeft === vbaEmpty ? (typeof rawRight === 'string' ? '' : 0) : rawLeft;
        let right = rawRight === vbaEmpty ? (typeof left === 'string' ? '' : 0) : rawRight;
        if (left instanceof VbaDate && right instanceof VbaDate) {
            return left.value === right.value ? 0 : (left.value < right.value ? -1 : 1);
        }
        if (left instanceof VbaDecimal || right instanceof VbaDecimal ||
            left instanceof VbaCurrency || right instanceof VbaCurrency) {
            const asDecimal = (value: any): VbaDecimal => {
                if (value instanceof VbaDecimal) return value;
                if (value instanceof VbaCurrency) return new VbaDecimal(value.internal, 4);
                if (typeof value === 'bigint') return VbaDecimal.fromString(value.toString());
                if (typeof value === 'string') return VbaDecimal.fromString(value.trim());
                return VbaDecimal.fromNumber(this.toVbaNumber(value));
            };
            const l = asDecimal(left), r = asDecimal(right);
            const scale = Math.max(l.scale, r.scale);
            const lm = l.mantissa * (10n ** BigInt(scale - l.scale));
            const rm = r.mantissa * (10n ** BigInt(scale - r.scale));
            return lm === rm ? 0 : (lm < rm ? -1 : 1);
        }
        if (typeof left === 'string' && typeof right === 'string' && this.getComparisonMode() === 'Text') {
            const l = left.toLowerCase();
            const r = right.toLowerCase();
            return l === r ? 0 : (l < r ? -1 : 1);
        }
        if (typeof left === 'bigint' || typeof right === 'bigint') {
            const exact = (value: any): bigint => {
                if (typeof value === 'bigint') return value;
                if (typeof value === 'string' && /^[+-]?\d+$/.test(value.trim())) return BigInt(value.trim());
                return BigInt(this.toVbaNumber(value));
            };
            const l = exact(left), r = exact(right);
            return l === r ? 0 : (l < r ? -1 : 1);
        }
        if (typeof left === 'string' && (typeof right === 'number' || right instanceof VbaBoolean || right instanceof VbaDate)) {
            return this.compareVbaValues(this.toVbaNumber(left), right);
        }
        if (typeof right === 'string' && (typeof left === 'number' || left instanceof VbaBoolean || left instanceof VbaDate)) {
            return this.compareVbaValues(left, this.toVbaNumber(right));
        }
        if (typeof left === 'number' && typeof right === 'number') {
            return left === right ? 0 : (left < right ? -1 : 1);
        }
        return undefined;
    }

    private evaluateSelectCaseStatement(stmt: SelectCaseStatement) {
        const selectVal = this.evaluateExpression(stmt.expression);
        const isVbaNull = (value: any): boolean => value === vbaNull;
        const normalizeEmpty = (left: any, right: any): [any, any] => {
            if (left === vbaEmpty) left = typeof right === 'string' ? '' : 0;
            if (right === vbaEmpty) right = typeof left === 'string' ? '' : 0;
            return [left, right];
        };
        const selectCaseCompare = (left: any, right: any): number | undefined =>
            this.compareVbaValues(left, right);
        const selectCaseEquals = (left: any, right: any): boolean => {
            // Empty is context-sensitive in VBA: it is 0 in a numeric Case
            // and an empty string in a string Case.  Keep this normalization
            // local to Select Case so DateSerial and ordinary expressions keep
            // their existing coercion rules.
            [left, right] = normalizeEmpty(left, right);
            const compared = selectCaseCompare(left, right);
            if (compared !== undefined) return compared === 0;
            if (left instanceof VbaDate && right instanceof VbaDate) return left.value === right.value;
            if (typeof left === 'bigint' || typeof right === 'bigint') {
                const bigintValue = typeof left === 'bigint' ? left : right;
                const otherValue = typeof left === 'bigint' ? right : left;
                if (typeof otherValue === 'bigint') return bigintValue === otherValue;
                if (typeof otherValue === 'number') {
                    if (Number.isSafeInteger(otherValue)) return bigintValue === BigInt(otherValue);
                    return Number(bigintValue) === otherValue;
                }
            }
            if (left instanceof VbaDecimal || right instanceof VbaDecimal) {
                try {
                    const asDecimal = (value: any): VbaDecimal =>
                        value instanceof VbaDecimal ? value : VbaDecimal.fromNumber(this.toVbaNumber(value));
                    const leftDecimal = asDecimal(left);
                    const rightDecimal = asDecimal(right);
                    const scale = Math.max(leftDecimal.scale, rightDecimal.scale);
                    const leftMantissa = leftDecimal.mantissa * (10n ** BigInt(scale - leftDecimal.scale));
                    const rightMantissa = rightDecimal.mantissa * (10n ** BigInt(scale - rightDecimal.scale));
                    return leftMantissa === rightMantissa;
                } catch { return false; }
            }
            if (left instanceof VbaCurrency || right instanceof VbaCurrency) {
                try { return this.toVbaNumber(left) === this.toVbaNumber(right); } catch { return false; }
            }
            return left === right;
        };

        for (const caseClause of stmt.cases) {
            let matched = false;
            for (const range of caseClause.ranges) {
                if (range.kind === 'expression') {
                    const value = this.evaluateExpression(range.value);
                    matched = !isVbaNull(selectVal) && !isVbaNull(value) && selectCaseEquals(selectVal, value);
                } else if (range.kind === 'to') {
                    const start = this.evaluateExpression(range.start);
                    const end = this.evaluateExpression(range.end);
                    if (isVbaNull(selectVal) || isVbaNull(start) || isVbaNull(end)) {
                        matched = false;
                    } else {
                        const lower = selectCaseCompare(selectVal, start);
                        const upper = selectCaseCompare(selectVal, end);
                        matched = lower !== undefined && upper !== undefined
                            ? lower >= 0 && upper <= 0
                            : selectVal >= start && selectVal <= end;
                    }
                } else {
                    // comparison: selectVal <op> value
                    const val = this.evaluateExpression(range.value);
                    if (isVbaNull(selectVal) || isVbaNull(val)) {
                        matched = false;
                        continue;
                    }
                    switch (range.operator) {
                        case '=':  matched = selectCaseEquals(selectVal, val); break;
                        case '<>': matched = !selectCaseEquals(selectVal, val); break;
                        case '<':  { const cmp = selectCaseCompare(selectVal, val); matched = cmp !== undefined ? cmp < 0 : selectVal < val; break; }
                        case '>':  { const cmp = selectCaseCompare(selectVal, val); matched = cmp !== undefined ? cmp > 0 : selectVal > val; break; }
                        case '<=': { const cmp = selectCaseCompare(selectVal, val); matched = cmp !== undefined ? cmp <= 0 : selectVal <= val; break; }
                        case '>=': { const cmp = selectCaseCompare(selectVal, val); matched = cmp !== undefined ? cmp >= 0 : selectVal >= val; break; }
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

    /** Shared VBA array arity and bound checks for indexed reads and writes. */
    private validateArrayArity(dims: { lower: number, upper: number }[] | undefined, count: number): void {
        if (dims && count !== dims.length) {
            this.throwVbaError(VbaErrorCode.SUBSCRIPT_OUT_OF_RANGE, 'Subscript out of range');
        }
    }

    private validateArrayIndex(
        dims: { lower: number, upper: number }[] | undefined,
        dimension: number,
        index: number,
    ): void {
        if (!Number.isFinite(index) || !Number.isInteger(index)) {
            this.throwVbaError(VbaErrorCode.SUBSCRIPT_OUT_OF_RANGE, 'Subscript out of range');
        }
        const bounds = dims?.[dimension];
        if (bounds && (index < bounds.lower || index > bounds.upper)) {
            this.throwVbaError(VbaErrorCode.SUBSCRIPT_OUT_OF_RANGE, 'Subscript out of range');
        }
    }

    /** Evaluate and validate one VBA array subscript through the shared path. */
    private evaluateArrayIndex(expression: Expression, dims: { lower: number, upper: number }[] | undefined, dimension: number): number {
        const index = Number(this.evaluateExpression(expression));
        this.validateArrayIndex(dims, dimension, index);
        return index;
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
        const obj = this.resolveAutoInstance(stmt.expression, this.evaluateExpression(stmt.expression));
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

    /** Execute a GoSub target in the owning procedure and stop at Return. */
    private executeGoSubLabel(label: string): void {
        const body = this.currentProcBody;
        if (!body) {
            this.throwVbaError(VbaErrorCode.SUB_OR_FUNCTION_NOT_DEFINED,
                `Sub or Function not defined: label '${label}'`);
        }
        const labelIndex = this.findLabelInBody(body, label);
        if (labelIndex < 0) {
            this.throwVbaError(VbaErrorCode.SUB_OR_FUNCTION_NOT_DEFINED,
                `Sub or Function not defined: label '${label}'`);
        }
        this.gosubTargetDepth++;
        try {
            this.executeStatements(body, labelIndex + 1, true);
        } catch (e: any) {
            if (e && e.type === 'Return') return;
            throw e;
        } finally {
            this.gosubTargetDepth--;
        }
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
        // UDT 間コピー（LSet b = a）: 実 VBA はメモリコピー。ここでは同一レイアウト
        // （フィールド数が一致）の場合に位置ベースでフィールド値をコピーする（Bug 32-E）
        if (stmt.left.type === 'Identifier') {
            const name = (stmt.left as Identifier).name;
            const target = this.env.get(name);
            if (target && typeof target === 'object' && (target as any).__vbaUdt__ === true) {
                const source = this.evaluateExpression(stmt.right);
                if (!(source && typeof source === 'object' && (source as any).__vbaUdt__ === true)) {
                    this.throwVbaError(VbaErrorCode.TYPE_MISMATCH, 'Type mismatch: LSet target is a user-defined type but source is not');
                }
                const fieldNames = (obj: any) => Object.keys(obj).filter(k => !k.startsWith('__'));
                const tFields = fieldNames(target);
                const sFields = fieldNames(source);
                if (tFields.length !== sFields.length) {
                    this.throwVbaError(VbaErrorCode.INVALID_PROCEDURE_CALL,
                        'LSet between user-defined types with different layouts is not supported');
                }
                for (let i = 0; i < tFields.length; i++) {
                    (target as any)[tFields[i]] = (source as any)[sFields[i]];
                }
                return;
            }
        }
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
            this.invokeBuiltin(errObj.raise.bind(errObj), [errNum]);
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
                let args = this.evaluateExpressions(stmt.args);
                for (const handler of handlers) {
                    const updatedArgs = handler(...args);
                    if ((handler as any).__vbaEventByRef__ && Array.isArray(updatedArgs)) {
                        args = updatedArgs;
                    }
                }
                const eventDecl = (me.__classDef__ as ClassDeclaration | undefined)?.body
                    .find((s: any) => s.type === 'EventDeclaration' &&
                        s.name.name.toLowerCase() === eventName) as EventDeclaration | undefined;
                if (eventDecl) {
                    for (let i = 0; i < eventDecl.parameters.length && i < stmt.args.length; i++) {
                        if (!eventDecl.parameters[i].isByVal) {
                            try {
                                this.evaluateAssignmentToVariable(stmt.args[i], args[i]);
                            } catch {
                                // Invalid ByRef targets are diagnosed by the normal
                                // call/precheck path; keep event dispatch compatible.
                            }
                        }
                    }
                }
            }
        }
    }

    private evaluateImplementsDirective(_stmt: ImplementsDirective) {
        // No-op for now. Used for interface compliance metadata.
    }

    private evaluateAppActivateStatement(stmt: AppActivateStatement) {
        this.evaluateInteractionStatement('AppActivate', stmt.title, stmt.wait);
    }

    private evaluateSendKeysStatement(stmt: SendKeysStatement) {
        this.evaluateInteractionStatement('SendKeys', stmt.keys, stmt.wait);
    }

    private evaluateInteractionStatement(name: string, required: Expression, wait?: Expression) {
        const args = wait === undefined ? [required] : [required, wait];
        this.evaluateCallExpression({
            type: 'CallExpression',
            callee: { type: 'Identifier', name } as Identifier,
            args,
        });
    }

    private requireFileMode(
        handle: { mode: OpenStatement['mode'] },
        allowed: OpenStatement['mode'][],
    ): void {
        if (!allowed.includes(handle.mode)) {
            this.throwVbaError(VbaErrorCode.BAD_FILE_MODE, "Bad file mode");
        }
    }

    private evaluateLockStatement(stmt: LockStatement) {
        const fileNum = this.evaluateFileNumber(stmt.fileNumber);
        const handle = this.fileHandles.get(fileNum);
        if (!handle) this.throwVbaError(VbaErrorCode.BAD_FILE_NAME_OR_NUMBER, "Bad file name or number");
        const range = this.evaluateLockRange(handle.mode, stmt.recordRange);
        const locks = handle.locks ?? (handle.locks = []);
        if (locks.some(lock => lock.start <= range.end && range.start <= lock.end)) {
            this.throwVbaError(VbaErrorCode.PATH_FILE_ACCESS_ERROR, "Path/File access error");
        }
        const pathLocks = this.fileLocks.get(handle.path) ?? [];
        if (pathLocks.some(lock => lock.owner !== fileNum
            && lock.start <= range.end && range.start <= lock.end)) {
            this.throwVbaError(VbaErrorCode.PERMISSION_DENIED, "Permission denied");
        }
        locks.push(range);
        pathLocks.push({ owner: fileNum, ...range });
        this.fileLocks.set(handle.path, pathLocks);
    }

    private evaluateUnlockStatement(stmt: UnlockStatement) {
        const fileNum = this.evaluateFileNumber(stmt.fileNumber);
        const handle = this.fileHandles.get(fileNum);
        if (!handle) this.throwVbaError(VbaErrorCode.BAD_FILE_NAME_OR_NUMBER, "Bad file name or number");
        const range = this.evaluateLockRange(handle.mode, stmt.recordRange);
        const locks = handle.locks ?? [];
        const index = locks.findIndex(lock => lock.key === range.key);
        if (index < 0) this.throwVbaError(VbaErrorCode.PATH_FILE_ACCESS_ERROR, "Path/File access error");
        locks.splice(index, 1);
        const pathLocks = this.fileLocks.get(handle.path) ?? [];
        const globalIndex = pathLocks.findIndex(lock => lock.owner === fileNum && lock.key === range.key);
        if (globalIndex >= 0) pathLocks.splice(globalIndex, 1);
        if (pathLocks.length > 0) this.fileLocks.set(handle.path, pathLocks);
        else this.fileLocks.delete(handle.path);
    }

    private evaluateLockRange(
        mode: OpenStatement['mode'],
        range?: { start?: Expression, end?: Expression },
    ): { start: number, end: number, key: string } {
        const requested = range
            ? {
                start: range.start ? Number(this.evaluateExpression(range.start)) : undefined,
                end: range.end ? Number(this.evaluateExpression(range.end)) : undefined,
            }
            : undefined;
        const key = requested
            ? `${requested.start ?? ''}:${requested.end ?? ''}`
            : '';
        // Sequential files do not have record ranges. VBA treats any supplied
        // range as a request to lock or unlock the whole file, but validates
        // an explicitly supplied range before applying that normalization.
        if (mode === 'Input' || mode === 'Output' || mode === 'Append') {
            if (requested) {
                const start = requested.start;
                const end = requested.end;
                if ((start !== undefined && (!Number.isFinite(start) || !Number.isInteger(start) || start < 0))
                    || (end !== undefined && (!Number.isFinite(end) || !Number.isInteger(end) || end < 1))
                    || (start !== undefined && end !== undefined && end < start)) {
                    this.throwVbaError(VbaErrorCode.PERMISSION_DENIED, "Permission denied");
                }
            }
            return { start: 0, end: Infinity, key };
        }
        if (!requested) return { start: 0, end: Infinity, key };
        const start = requested.start ?? 1;
        const end = requested.end ?? start;
        if (!Number.isInteger(start) || !Number.isInteger(end) || start < 1 || end < start) {
            this.throwVbaError(VbaErrorCode.INVALID_PROCEDURE_CALL, "Invalid procedure call or argument");
        }
        return { start, end, key };
    }

    private evaluateWidthStatement(stmt: WidthStatement) {
        const fileNum = this.evaluateFileNumber(stmt.fileNumber);
        const width = Number(this.evaluateExpression(stmt.width));
        const handle = this.fileHandles.get(fileNum);
        if (!handle) this.throwVbaError(VbaErrorCode.BAD_FILE_NAME_OR_NUMBER, "Bad file name or number");
        if (!Number.isInteger(width) || width < 0 || width > 255) {
            this.throwVbaError(VbaErrorCode.INVALID_PROCEDURE_CALL, "Invalid procedure call or argument");
        }
        handle.width = width;
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

        this.evaluateAssignmentToVariable(stmt.left, val, stmt.right);

        // Track numeric subtype for Variant variables so TypeName/VarType reflect the assigned type
        if (typeof val === 'number') {
            if (stmt.left.type === 'Identifier') {
                const name = (stmt.left as Identifier).name;
                const typeInfo = this.env.getVariableType(name);
                if (!typeInfo || typeInfo.vbaType === 'Variant') {
                    const subtype = this.resolveNumericSubtype(stmt.right);
                    if (subtype) {
                        this.env.setVariantSubtype(name, subtype);
                    } else {
                        this.env.clearVariantSubtype(name);
                    }
                }
            } else if (stmt.left.type === 'CallExpression') {
                // Bug CG: Track numeric subtype for Variant array element assignments (arr(i) = 42)
                // so that TypeName/VarType on the element returns the correct subtype (not 'Double').
                // Index args are re-evaluated here; this is safe since index expressions in VBA
                // are almost universally side-effect-free (literals, simple variables).
                const call = stmt.left as CallExpression;
                if (call.callee.type === 'Identifier') {
                    const arrName = (call.callee as Identifier).name;
                    const arr = this.env.get(arrName);
                    if (Array.isArray(arr)) {
                        const elemType = (arr as any).__vbaElementType__;
                        if (!elemType || elemType === 'variant') {
                            const idxs = this.evaluateIndexExpressions(call.args);
                            const key = idxs.join(',');
                            const subtype = this.resolveNumericSubtype(stmt.right);
                            if (!(arr as any).__vbaSubtypes__) (arr as any).__vbaSubtypes__ = Object.create(null);
                            if (subtype) (arr as any).__vbaSubtypes__[key] = subtype;
                            else delete (arr as any).__vbaSubtypes__[key];
                        }
                    }
                }
            }
        }
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
            case 'Decimal':
                return this.env.get('cdec')(val);
            case 'String':
                return this.env.get('cstr')(val);
            case 'Date':
                return this.env.get('cdate')(val);
            default:
                return val;
        }
    }

    private coerceToBoolean(val: any): VbaBoolean { return vbaToBoolean(val); }

    private rejectTypedArrayWholeAssignment(existing: any, value: any, sourceExpr?: Expression, allowArrayRebind = false) {
        const isArrayReturn = Array.isArray(value) && (value as any).__vbaArrayReturn__ === true;
        if (isArrayReturn) delete (value as any).__vbaArrayReturn__;
        if (!Array.isArray(existing) || !Array.isArray(value) || allowArrayRebind) return;
        const typed = (existing as any).__vbaElementType__ ||
            (existing as any).__vbaElementTypeName__ ||
            (existing as any).__vbaElementObjectTypeName__;
        const returnType = (value as any).__vbaArrayReturnType__ as string | undefined;
        // A call expression alone does not prove that the value is a legal
        // typed-array return. Only an evaluated procedure/host getter that
        // explicitly carries the array-return marker may bypass this guard.
        void sourceExpr;
        if (typed && (!isArrayReturn || (returnType && returnType.toLowerCase() !== String(typed).toLowerCase()))) {
            this.throwVbaError(VbaErrorCode.TYPE_MISMATCH, "Can't assign to an array");
        }
    }

    /** Class fields must not resolve through the caller's environment chain. */
    private getClassField(instanceEnv: Environment, name: string): any {
        return instanceEnv.hasOwnVariable(name) ? instanceEnv.get(name) : undefined;
    }

    private evaluateAssignmentToVariable(left: Expression, val: any, sourceExpr?: Expression, allowArrayRebind = false) {
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
                    const valueSetter = findClassProperty(classDef.procedures, 'value', 'let');
                    if (valueSetter) {
                        // Call the Value property setter with the assigned value
                        this.callClassMethodWithExpressions(variable, valueSetter, [sourceExpr ?? null], [val]);
                        return;
                    }
                    if (findClassProperty(classDef.procedures, 'value', 'set')) {
                        this.throwVbaError(VbaErrorCode.OBJECT_DOESNT_SUPPORT_PROPERTY, "Object doesn't support this property or method");
                    }
                }
            }

            const oldVal = this.env.get(name);
            if (val === vbaNothing && oldVal !== val) this.releaseRef(oldVal);

            // VBA coerces a String assigned to a dynamic Byte() into its UTF-16LE
            // code units. This is distinct from Binary Put/Get, which uses the
            // active ANSI code page for strings.
            if (Array.isArray(oldVal) && (oldVal as any).__vbaElementType__ === 'byte' && typeof val === 'string') {
                const bytes: any[] = [];
                for (let i = 0; i < val.length; i++) {
                    const codeUnit = val.charCodeAt(i);
                    bytes.push(codeUnit & 0xff, codeUnit >>> 8);
                }
                (bytes as any).vbaBase = 0;
                (bytes as any).vbaFixed = true;
                (bytes as any).__vbaElementType__ = 'byte';
                this.env.set(name, bytes);
                return;
            }

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

            // VBA does not allow assigning one whole typed array to another
            // (`target = source`).  Element assignment and Variant targets
            // are separate valid paths; reject only a declared typed-array
            // destination here before the normal Let-coercion can alias the
            // JavaScript array object.
            this.rejectTypedArrayWholeAssignment(oldVal, val, sourceExpr, allowArrayRebind);

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
            // UDTs are value types in VBA.  Assignment must not make the
            // destination share nested UDTs or arrays with the source.
            if (val && typeof val === 'object' && val.__vbaUdt__ === true) {
                val = this.deepCopyByValValue(val);
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
                        const d = this.evaluateArrayIndex(call.args[i], dims, i);
                        if (!current[d]) {
                            current[d] = [];
                        }
                        current = current[d];
                    }
                    const lastIdx = this.evaluateArrayIndex(call.args[call.args.length - 1], dims, call.args.length - 1);
                    // 型付き配列の要素代入: __vbaElementType__ に従って強制変換
                    const elemType = (target as any).__vbaElementType__;
                    if ((target as any).__vbaElementObjectTypeName__) {
                        this.throwVbaError(VbaErrorCode.OBJECT_REQUIRED, 'Object required');
                    }
                    if (elemType && elemType !== 'variant') {
                        val = this.coerceToDeclaredType(val,
                            elemType.charAt(0).toUpperCase() + elemType.slice(1));
                    }
                    const fixedLength = (target as any).__vbaElementFixedLength__ as number | undefined;
                    if (elemType === 'string' && fixedLength !== undefined && typeof val === 'string') {
                        val = val.length > fixedLength
                            ? val.slice(0, fixedLength)
                            : val + ' '.repeat(fixedLength - val.length);
                    }
                    current[lastIdx] = val;
                } else if (target && target.__isVbaDict__) {
                    // Treat as Dictionary assignment dict("key") = val
                    const key = target.__resolveKey__(this.evaluateExpression(call.args[0]));
                    target.__map__.set(key, val);
                } else if (target && target.__vbaClass__) {
                    // Default property assignment: obj(args) = val -> obj.Item(args) = val
                    const classDef = target.__classDef__ as ClassDeclaration;
                    const setter = findClassProperty(classDef.procedures, 'item', 'let');
                    if (setter) {
                        const argsVals = this.evaluateExpressions(call.args);
                        this.callClassMethodWithExpressions(target, setter, [...call.args, sourceExpr ?? null], [...argsVals, val]);
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
                    const key = obj.__resolveKey__(this.evaluateExpression(call.args[0]));
                    obj.__map__.set(key, val);
                } else if (obj && obj.__vbaClass__) {
                    const classDef = obj.__classDef__ as ClassDeclaration;
                    const instanceEnv = obj.__instanceEnv__ as Environment;
                    const fieldArr = this.getClassField(instanceEnv, methodName);
                    if (Array.isArray(fieldArr)) {
                        const requiredType = (fieldArr as any).__vbaElementObjectTypeName__ as string | undefined;
                        const actualType = (val as any)?.__className__ as string | undefined;
                        if (requiredType && val !== vbaNothing &&
                            (!actualType || actualType.toLowerCase() !== requiredType.toLowerCase())) {
                            this.throwVbaError(VbaErrorCode.TYPE_MISMATCH, 'Type mismatch');
                        }
                        const index = this.evaluateArrayIndex(call.args[0],
                            (fieldArr as any).__vbaDimensions__ as { lower: number, upper: number }[] | undefined, 0);
                        fieldArr[index] = val;
                        return;
                    }
                    const setter = findClassProperty(classDef.procedures, methodName, 'let');
                    if (setter) {
                        const argsVals = this.evaluateExpressions(call.args);
                        this.callClassMethodWithExpressions(obj, setter, [...call.args, sourceExpr ?? null], [...argsVals, val]);
                    } else {
                        if (findClassProperty(classDef.procedures, methodName, 'set')) {
                            this.throwVbaError(VbaErrorCode.OBJECT_DOESNT_SUPPORT_PROPERTY, "Object doesn't support this property or method");
                        }
                        // Property Let/Set がなければ、配列フィールドへの外部インデックス代入を試みる
                        // 例: obj.Items(0) = val
                        const instanceEnvForArr = obj.__instanceEnv__ as Environment;
                        const fieldArr = this.getClassField(instanceEnvForArr, methodName);
                        if (Array.isArray(fieldArr)) {
                            const dims = (fieldArr as any).__vbaDimensions__ as { lower: number, upper: number }[] | undefined;
                            if (dims && call.args.length !== dims.length) this.throwVbaError(VbaErrorCode.SUBSCRIPT_OUT_OF_RANGE, 'Subscript out of range');
                            let current = fieldArr;
                            for (let i = 0; i < call.args.length - 1; i++) {
                                const d = this.evaluateArrayIndex(call.args[i], dims, i);
                                if (!current[d]) current[d] = [];
                                current = current[d];
                            }
                            const lastIdx = this.evaluateArrayIndex(call.args[call.args.length - 1], dims, call.args.length - 1);
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
                    // UDT array member subscript assignment: udt.Items(i) = val
                    if (Array.isArray(method) && call.args.length > 0) {
                        const dims = (method as any).__vbaDimensions__ as { lower: number, upper: number }[] | undefined;
                        let current: any = method;
                        for (let i = 0; i < call.args.length - 1; i++) {
                            const d = this.evaluateArrayIndex(call.args[i], dims, i);
                            current = current[d];
                            if (!Array.isArray(current)) this.throwVbaError(VbaErrorCode.SUBSCRIPT_OUT_OF_RANGE, 'Subscript out of range');
                        }
                        const lastIdx = this.evaluateArrayIndex(call.args[call.args.length - 1], dims, call.args.length - 1);
                        const fixedLength = (method as any).__vbaElementFixedLength__ as number | undefined;
                        if (fixedLength !== undefined && typeof val === 'string') {
                            val = val.length > fixedLength
                                ? val.slice(0, fixedLength)
                                : val + ' '.repeat(fixedLength - val.length);
                        }
                        current[lastIdx] = val;
                        return;
                    }
                    if (typeof method === 'function') {
                        const argsVals = this.evaluateExpressions(call.args);
                        const result = method.call(obj, ...argsVals);
                        if (result && result.__vbaDefault__ === true) {
                            const valueKey = this.resolveObjectMemberKey(result, 'value') ?? 'Value';
                            result[valueKey] = val;
                            return;
                        }
                    }
                    const key = obj.__resolveKey__(this.evaluateExpression(call.args[0]));
                    obj[key] = val;
                } else {
                    this.throwVbaError(VbaErrorCode.INVALID_PROCEDURE_CALL, 'Invalid procedure call or argument');
                }
            } else if (call.callee.type === 'ImplicitWithObjectExpression') {
                // Bug CA: `.Data(0,0) = val` inside With block
                // Same logic as MemberExpression case but obj comes from withObjectStack
                if (this.withObjectStack.length === 0) {
                    this.throwVbaError(VbaErrorCode.OBJECT_VARIABLE_NOT_SET, 'Object variable or With block variable not set');
                }
                const implicitObj = this.withObjectStack[this.withObjectStack.length - 1];
                const implicitProp = (call.callee as ImplicitWithObjectExpression).property.name.toLowerCase();
                if (implicitObj && implicitObj.__vbaClass__) {
                    const implicitClassDef = implicitObj.__classDef__ as ClassDeclaration;
                    const implicitSetter = findClassProperty(implicitClassDef.procedures, implicitProp, 'let');
                    if (implicitSetter) {
                        const argsVals = this.evaluateExpressions(call.args);
                        this.callClassMethodWithExpressions(implicitObj, implicitSetter,
                            [...call.args, sourceExpr ?? null], [...argsVals, val]);
                    } else {
                        if (findClassProperty(implicitClassDef.procedures, implicitProp, 'set')) {
                            this.throwVbaError(VbaErrorCode.OBJECT_DOESNT_SUPPORT_PROPERTY, "Object doesn't support this property or method");
                        }
                        const implicitInstanceEnv = implicitObj.__instanceEnv__ as Environment;
                        const implicitFieldArr = implicitInstanceEnv.get(implicitProp);
                        if (Array.isArray(implicitFieldArr)) {
                            const dims = (implicitFieldArr as any).__vbaDimensions__ as { lower: number, upper: number }[] | undefined;
                            if (dims && call.args.length !== dims.length) this.throwVbaError(VbaErrorCode.SUBSCRIPT_OUT_OF_RANGE, 'Subscript out of range');
                            let current: any = implicitFieldArr;
                            for (let i = 0; i < call.args.length - 1; i++) {
                                const d = this.evaluateArrayIndex(call.args[i], dims, i);
                                if (!current[d]) current[d] = [];
                                current = current[d];
                            }
                            const lastIdx = this.evaluateArrayIndex(call.args[call.args.length - 1], dims, call.args.length - 1);
                            current[lastIdx] = val;
                        } else {
                            this.throwVbaError(VbaErrorCode.OBJECT_DOESNT_SUPPORT_PROPERTY, "Object doesn't support this property or method");
                        }
                    }
                } else if (implicitObj && typeof implicitObj === 'object') {
                    const resolvedKey = this.resolveObjectMemberKey(implicitObj, implicitProp) ?? implicitProp;
                    const fieldArr = implicitObj[resolvedKey];
                    if (Array.isArray(fieldArr)) {
                        const dims = (fieldArr as any).__vbaDimensions__ as { lower: number, upper: number }[] | undefined;
                        const lastIdx = this.evaluateArrayIndex(call.args[call.args.length - 1], dims, call.args.length - 1);
                        fieldArr[lastIdx] = val;
                    } else {
                        this.throwVbaError(VbaErrorCode.OBJECT_DOESNT_SUPPORT_PROPERTY, "Object doesn't support this property or method");
                    }
                }
            } else if (call.callee.type === 'CallExpression') {
                // outer("sub")("x") = val  →  evaluate outer("sub") to get inner dict, then assign
                const innerObj = this.evaluateExpression(call.callee);
                if (innerObj && innerObj.__isVbaDict__) {
                    const key = innerObj.__resolveKey__(this.evaluateExpression(call.args[0]));
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
        } else if (left.type === 'ImplicitWithDictionaryAccessExpression') {
            if (this.withObjectStack.length === 0) {
                this.throwVbaError(VbaErrorCode.OBJECT_VARIABLE_NOT_SET, 'Object variable or With block variable not set');
            }
            const obj = this.withObjectStack[this.withObjectStack.length - 1];
            if (obj && obj.__isVbaDict__) {
                obj.__map__.set((left as ImplicitWithDictionaryAccessExpression).property.name, val);
            } else {
                this.throwVbaError(VbaErrorCode.OBJECT_DOESNT_SUPPORT_PROPERTY, "Object doesn't support this property or method");
            }
        } else if (left.type === 'MemberExpression') {
            const member = left as MemberExpression;
            let memberObject: any;
            if (member.object.type === 'CallExpression') {
                const call = member.object as CallExpression;
                const isCurrentArrayReturn = call.callee.type === 'Identifier' && this.currentProcedureName &&
                    (call.callee as Identifier).name.toLowerCase() === this.currentProcedureName.toLowerCase();
                if (isCurrentArrayReturn) {
                    memberObject = this.env.get((call.callee as Identifier).name);
                    for (const indexExpr of call.args) {
                        memberObject = memberObject?.[this.evaluateExpression(indexExpr) as number];
                    }
                } else {
                    memberObject = this.evaluateExpression(member.object);
                }
            } else {
                memberObject = this.evaluateExpression(member.object);
            }
            const obj = this.resolveAutoInstance(member.object, memberObject);
            const propName = member.property.name.toLowerCase();
            if (obj && obj.__vbaClass__) {
                const classDef = obj.__classDef__ as ClassDeclaration;
                const instanceEnv = obj.__instanceEnv__ as Environment;
                this.rejectTypedArrayWholeAssignment(this.getClassField(instanceEnv, propName), val, sourceExpr);
                const setter = findClassProperty(classDef.procedures, propName, 'let');
                if (setter) {
                    this.callClassMethodWithExpressions(obj, setter, [sourceExpr ?? null], [val]);
                } else {
                    if (findClassProperty(classDef.procedures, propName, 'set')) {
                        this.throwVbaError(VbaErrorCode.OBJECT_DOESNT_SUPPORT_PROPERTY, "Object doesn't support this property or method");
                    }
                    // Bug CC: enforce fixed-length string truncation/padding for class fields
                    const fl = obj.__fixedLengths__?.[propName];
                    if (fl !== undefined && typeof val === 'string') {
                        if (val.length > fl) val = val.slice(0, fl);
                        else if (val.length < fl) val = val + ' '.repeat(fl - val.length);
                    }
                    instanceEnv.setLocally(propName, val);
                }
            } else if (obj && typeof obj === 'object') {
                // VBA は大文字小文字不問: アクセサ(setter)・プロトタイプも辿って実キーを解決する
                const resolvedProp = this.resolveObjectMemberKey(obj, propName) ?? propName;
                this.rejectTypedArrayWholeAssignment(obj[resolvedProp], val, sourceExpr);
                // UDT fixed-length string member coercion
                const fl = obj.__fixedLengths__?.[propName];
                if (fl !== undefined && typeof val === 'string') {
                    if (val.length > fl) val = val.slice(0, fl);
                    else if (val.length < fl) val = val + ' '.repeat(fl - val.length);
                }
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
            if (obj && obj.__vbaClass__) {
                const classDef = obj.__classDef__ as ClassDeclaration;
                const instanceEnv = obj.__instanceEnv__ as Environment;
                this.rejectTypedArrayWholeAssignment(this.getClassField(instanceEnv, propName), val, sourceExpr);
                const setter = findClassProperty(classDef.procedures, propName, 'let');
                if (setter) {
                    this.callClassMethodWithExpressions(obj, setter, [sourceExpr ?? null], [val]);
                } else {
                    if (findClassProperty(classDef.procedures, propName, 'set')) {
                        this.throwVbaError(VbaErrorCode.OBJECT_DOESNT_SUPPORT_PROPERTY, "Object doesn't support this property or method");
                    }
                    // Bug CC: enforce fixed-length string for With-block class field assignment
                    const fl = obj.__fixedLengths__?.[propName];
                    if (fl !== undefined && typeof val === 'string') {
                        if (val.length > fl) val = val.slice(0, fl);
                        else if (val.length < fl) val = val + ' '.repeat(fl - val.length);
                    }
                    instanceEnv.setLocally(propName, val);
                }
            } else if (obj && typeof obj === 'object') {
                this.rejectTypedArrayWholeAssignment(obj[propName], val, sourceExpr);
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
                    this.env.setVariableType(varName, { vbaType: mapped, fixedLength: decl.fixedLength });
                } else if (this.classDefinitions.has(effectiveType.toLowerCase())) {
                    this.env.setVariableType(varName, {
                        vbaType: 'Object',
                        objectTypeName: effectiveType,
                    });
                } else {
                    // Bug CB: Enum-typed variable (e.g. `Dim c As Color`) — map to 'Long'
                    // so TypeName/VarType reflects the underlying numeric type instead of "Double"
                    const enumObj = this.env.getConst(effectiveType);
                    // Bug CI: Exclude VbaNamespaceRef (module/class names stored in env) to prevent
                    // treating class types like MyClass as enums when they share an env key.
                    if (enumObj && typeof enumObj === 'object' && !enumObj.__vbaClass__ && !enumObj.__vbaTypeName__
                            && !(enumObj instanceof VbaNamespaceRef)) {
                        this.env.setVariableType(varName, { vbaType: 'Long' });
                    }
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
                if (['integer', 'long', 'single', 'double', 'currency', 'byte'].includes(t)) {
                    initialValue = 0;
                } else if (t === 'longlong' || t === 'longptr') {
                    initialValue = 0n;
                } else if (t === 'string') {
                    // Fixed-length strings initialize with null characters (VBA spec)
                    initialValue = decl.fixedLength !== undefined ? '\0'.repeat(decl.fixedLength) : '';
                } else if (t === 'boolean') {
                    initialValue = 0; // vbaFalse
                } else if (
                    this.classDefinitions.has(t) || this.externalObjectFactories.has(t) ||
                    t === 'object' || t === 'collection'
                ) {
                    initialValue = vbaNothing;
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
                // 型付き配列: 要素代入時の coerceToDeclaredType に使う型名を保持
                if (effectiveType) {
                    (initialValue as any).__vbaElementType__ = effectiveType.toLowerCase();
                    if (effectiveType.toLowerCase() === 'string' && decl.fixedLength !== undefined) {
                        (initialValue as any).__vbaElementFixedLength__ = decl.fixedLength;
                    }
                }
                // UDT 型配列: ReDim 時に要素を初期化できるよう型名を保持する
                if (decl.objectType && this.env.getType(decl.objectType)) {
                    (initialValue as any).__vbaElementTypeName__ = decl.objectType;
                    // B-4: For fixed-size UDT arrays, initialize each element with a proper UDT instance
                    if ((initialValue as any).vbaFixed) {
                        this.fillArrayWithUdtInstances(initialValue, decl.objectType);
                    }
                }
                if (decl.objectType && (
                    this.classDefinitions.has(decl.objectType.toLowerCase()) ||
                    this.externalObjectFactories.has(decl.objectType.toLowerCase()) ||
                    ['object', 'collection'].includes(decl.objectType.toLowerCase())
                )) {
                    (initialValue as any).__vbaElementObjectTypeName__ = decl.objectType;
                    // `Dim items(0 To n) As New ClassName` gives each array
                    // element the same lazy-instantiation semantics as a
                    // scalar `Dim item As New ClassName`.
                    if (decl.isNew && this.classDefinitions.has(decl.objectType.toLowerCase())) {
                        (initialValue as any).__vbaElementAutoNew__ = true;
                        if ((initialValue as any).vbaFixed) {
                            this.fillArrayWithAutoInstancePlaceholders(initialValue, decl.objectType);
                        }
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
        const moduleName = this.currentSourceModule || this.executingModuleName;
        if (moduleName) this.moduleComparisonModes.set(moduleName.toLowerCase(), stmt.mode);
    }

    private getComparisonMode(): 'Binary' | 'Text' {
        const moduleName = this.executingModuleName || this.currentSourceModule;
        return moduleName
            ? (this.moduleComparisonModes.get(moduleName.toLowerCase()) ?? this.comparisonMode)
            : this.comparisonMode;
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
        for (const bodyStmt of stmt.body) {
            if (bodyStmt.type === 'OptionCompareStatement') {
                this.moduleComparisonModes.set(stmt.name.toLowerCase(), (bodyStmt as OptionCompareStatement).mode);
            }
        }
        for (const proc of stmt.procedures) proc.moduleName = stmt.name;
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
        // Keep module/global visibility without capturing the transient caller
        // procedure environment as a class-field parent.
        const moduleName = (this.executingModuleName || this.currentSourceModule || '').toLowerCase();
        const parentEnv = this.moduleEnvs.get(moduleName) ?? this.globalEnv;
        const instanceEnv = new Environment(parentEnv);

        // Bug CC: Track fixed-length string fields for enforcement on assignment
        const classFixedLengths: Record<string, number> = {};

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
                if (mt === 'string') {
                    if (decl.fixedLength !== undefined) {
                        defaultVal = ' '.repeat(decl.fixedLength);
                        classFixedLengths[decl.name.name.toLowerCase()] = decl.fixedLength;
                    } else {
                        defaultVal = '';
                    }
                }
                else if (['integer', 'long', 'single', 'double', 'currency', 'byte', 'longlong', 'longptr'].includes(mt)) defaultVal = 0;
                else if (mt === 'boolean') defaultVal = 0; // vbaFalse
                else if (decl.objectType && this.env.getType(decl.objectType)) {
                    // UDT (Type ... End Type) フィールド: 既定値では Empty のままになり、
                    // Class_Initialize 等でのメンバー代入が Error 91 になっていた。
                    // Dim 変数の UDT 初期化と同じ instantiateType() を使う。
                    defaultVal = this.instantiateType(decl.objectType);
                } else if (decl.isNew && decl.objectType && (
                    this.classDefinitions.has(mt) || this.externalObjectFactories.has(mt) ||
                    mt === 'collection'
                )) {
                    // Bug BY: `Public P As New ClassName` — As New フィールドをすぐにインスタンス化する
                    if (mt === 'collection') {
                        defaultVal = new VbaCollection();
                    } else if (this.externalObjectFactories.has(mt)) {
                        defaultVal = this.externalObjectFactories.get(mt)!();
                    } else {
                        const nestedDef = this.classDefinitions.get(mt);
                        if (nestedDef) defaultVal = this.createInstanceFromDef(nestedDef);
                    }
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
                    if (['byte', 'integer', 'long', 'single', 'double', 'currency', 'longlong', 'longptr', 'string', 'boolean', 'date'].includes(mt)) {
                        (defaultVal as any).__vbaElementType__ = mt;
                    }
                    if (decl.objectType && this.env.getType(decl.objectType)) {
                        (defaultVal as any).__vbaElementTypeName__ = decl.objectType;
                        if ((defaultVal as any).vbaFixed) {
                            this.fillArrayWithUdtInstances(defaultVal, decl.objectType);
                        }
                    }
                    if (decl.objectType && (
                        this.classDefinitions.has(mt) || this.externalObjectFactories.has(mt) ||
                        ['object', 'collection'].includes(mt)
                    )) {
                        (defaultVal as any).__vbaElementObjectTypeName__ = decl.objectType;
                        if (decl.isNew && this.classDefinitions.has(mt)) {
                            (defaultVal as any).__vbaElementAutoNew__ = true;
                            if ((defaultVal as any).vbaFixed) {
                                this.fillArrayWithAutoInstancePlaceholders(defaultVal, decl.objectType);
                            }
                        }
                    }
                }
                instanceEnv.setLocally(decl.name.name, defaultVal);
                if (decl.objectType && this.classDefinitions.has(mt) && !decl.isArray) {
                    instanceEnv.setVariableType(decl.name.name, {
                        vbaType: 'Object',
                        objectTypeName: decl.objectType,
                    });
                }
                if (decl.isWithEvents) instanceEnv.setWithEvents(decl.name.name);
            }
        }

        const instance: any = {
            __vbaClass__: true,
            __refCount__: 0,
            __className__: classDef.name,
            __vbaTypeName__: classDef.name,
            __classDef__: classDef,
            __instanceEnv__: instanceEnv,
            __events__: new Map<string, ((...args: any[]) => void)[]>(),
        };
        // Bug CC: attach fixed-length info so assignment can truncate/pad
        if (Object.keys(classFixedLengths).length > 0) {
            instance.__fixedLengths__ = classFixedLengths;
        }

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
                    for (const decl of (stmt as ConstDeclaration).declarations) {
                        const constVal = this.evaluateConstValue(decl.value);
                        instanceEnv.setConstant(decl.name.name, constVal);
                    }
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
     * 指定モジュールの手続きをモジュール修飾なし（グローバル）に昇格させる（後勝ち）。
     * VBA モックで複数ファイルが同名手続きを定義した場合に最後ロードされたものを有効にする。
     */
    public promoteProcedures(moduleName: string): void {
        this.env.promoteProceduresFromModule(moduleName);
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
        const split = this.splitArgumentExpressions(argExprs);
        for (const entry of split.ordered) {
            const argExpr = entry.expression;
            if (entry.name !== undefined) {
                namedArgs.set(entry.name, this.evaluateExpression(argExpr));
                continue;
            }
            if (argExpr.type === 'MissingArgument') {
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
        // 仕様に存在しない名前付き引数は Error 448（Named argument not found）
        for (const nameLower of namedArgs.keys()) {
            if (!spec.some(p => p.name.toLowerCase() === nameLower)) {
                this.throwVbaError(448, `Named argument not found: '${nameLower}'`);
            }
        }
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
                result.push(this.coerceBoundArgument(namedArgs.get(nameLower), p));
            } else if (i < positionalArgs.length) {
                result.push(this.coerceBoundArgument(positionalArgs[i], p));
            } else {
                result.push(undefined); // JS 側のデフォルト引数構文に委ねる
            }
        }
        return result;
    }

    /** Apply only explicitly requested host/COM parameter coercions. */
    private coerceBoundArgument(value: any, spec: BuiltinParamSpec): any {
        if (value === undefined || !spec.coerce) return value;
        if (spec.coerce === 'boolean') return vbaToBoolean(value);
        if (spec.coerce === 'string') return vbaToString(value);
        if (spec.coerce === 'long') {
            // Null/Missing remain visible to the builtin so it can apply its
            // function-specific propagation or error policy. Empty follows
            // VBA numeric coercion and becomes zero.
            if (value === vbaNull || value === vbaMissing) return value;
            return this.vbaRound(this.toVbaNumber(value));
        }
        return value;
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

        const knownNames = new Set(overloads.flatMap(o => o.params.map(p => p.name.toLowerCase())));
        for (const namedKey of namedArgs.keys()) {
            if (!knownNames.has(namedKey)) {
                this.throwVbaError(448, `Named argument not found: '${namedKey}'`);
            }
        }

        if (namedArgs.size === 0) {
            if (!arities.includes(totalCount)) {
                if (totalCount < Math.min(...arities)) {
                    this.throwVbaError(VbaErrorCode.ARGUMENT_NOT_OPTIONAL, 'Argument not optional');
                }
                this.throwVbaError(VbaErrorCode.WRONG_NUMBER_OF_ARGUMENTS, 'Wrong number of arguments or invalid property assignment');
            }
            const selected = overloads.find(o => o.params.length === totalCount)!;
            return positionalArgs.map((value, index) =>
                this.coerceBoundArgument(value, selected.params[index]));
        }

        const candidates = overloads.filter(o => {
            if (positionalArgs.length > o.params.length) return false;
            const paramNames = o.params.map(p => p.name.toLowerCase());
            for (const namedKey of namedArgs.keys()) {
                if (!paramNames.includes(namedKey)) return false;
            }
            // 位置引数が占める先頭スロットの名前が、名前付き引数と重複していないこと
            for (let i = 0; i < positionalArgs.length; i++) {
                if (namedArgs.has(paramNames[i])) return false;
            }
            // 残り（位置引数で埋まらないスロット）は名前付き引数、または
            // Optional の省略で埋まること。先頭 Optional を飛ばした
            // `InStr(String1:=..., String2:=..., Compare:=...)` を許可する。
            for (let i = positionalArgs.length; i < o.params.length; i++) {
                if (!namedArgs.has(paramNames[i]) && !o.params[i].optional) return false;
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
            // 複数の形が同じ名前集合に適用できる場合は、省略スロットが
            // 最も少ない（より具体的な）形を選ぶ。InStr の2引数形を
            // Start省略の3引数形より優先するための決定規則。
            candidates.sort((a, b) =>
                (a.params.length - totalCount) - (b.params.length - totalCount));
            if (candidates[0].params.length !== candidates[1].params.length) {
                candidates.splice(1);
            }
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
                result.push(this.coerceBoundArgument(positionalArgs[i], overload.params[i]));
            } else {
                const value = namedArgs.get(overload.params[i].name.toLowerCase());
                result.push(this.coerceBoundArgument(value, overload.params[i]));
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
    /** 組み込み関数・メソッドを呼び出し、素の VbaError オブジェクトを枠組み付き Error に包み直す */
    private invokeBuiltin(fn: Function, args: any[], thisArg: any = null): any {
        try {
            return fn.apply(thisArg, args);
        } catch (e: any) {
            if (e?.type === 'VbaError' && !(e instanceof Error)) this.throwVbaError(e.number, e.message, undefined, undefined, e.vbaSource);
            throw e;
        }
    }

    private resolveCallArgs(fn: Function, argExprs: Expression[], nameForError: string): any[] {
        const overloads = (fn as any).__vbaOverloads__ as BuiltinOverload[] | undefined;
        if (overloads) {
            return this.bindOverloadedCallArguments(nameForError, overloads, argExprs);
        }
        const spec = (fn as any).__vbaParamSpec__ as BuiltinParamSpec[] | undefined;
        if (spec) {
            return this.bindCallArguments(spec, argExprs);
        }
        // 仕様のない関数に名前付き引数は解決できない。黙って位置引数化すると
        // 呼び出しの意味が変わるため Error 448 で明示的に失敗させる（Bug 33-B）
        const named = argExprs.find(a => a.type === 'NamedArgument');
        if (named) {
            this.throwVbaError(448, `Named argument not found: '${(named as NamedArgument).name}'`);
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

    /**
     * Call a class procedure from an expression and preserve ByRef arguments.
     * Class calls historically evaluated arguments to values and discarded the
     * callee's updates, unlike module procedure calls which write back through
     * their original expressions.
     */
    private callClassMethodWithExpressions(instance: any, proc: ProcedureDeclaration, argExprs: (Expression | null)[], evaluatedArgs?: any[]): any {
        const aligned = this.alignProcedureCallExpressions(proc, argExprs, evaluatedArgs);
        const { args, references } = aligned;
        const byRefValues: any[] = [];
        const result = this.callClassMethod(instance, proc, args, byRefValues);
        for (let i = 0; i < proc.parameters.length && i < args.length; i++) {
            const param = proc.parameters[i];
            if (param.isParamArray) {
                // bindProcedureParameters packs all arguments at and after the
                // ParamArray declaration into one array.  Keep each original
                // expression as a separate ByRef write-back target instead of
                // attempting to write the packed array to the parameter slot.
                const packed = byRefValues[i];
                if (Array.isArray(packed)) {
                    for (let j = i; j < args.length; j++) {
                        if (references[j]) {
                            try {
                                references[j]!.set(packed[j - i]);
                            } catch {
                                // Non-assignable expressions are temporaries.
                            }
                        }
                    }
                }
                break;
            }
            if (!param.isByVal && references[i]) {
                try {
                    references[i]!.set(byRefValues[i]);
                } catch {
                    // Non-assignable expressions are passed as temporary values.
                }
            }
        }
        return result;
    }

    /** Normalize positional and named call arguments to declaration order. */
    private splitArgumentExpressions(argExprs: (Expression | null)[]): {
        named: Map<string, Expression>;
        positional: Expression[];
        ordered: Array<{ name?: string; expression: Expression }>;
    } {
        const named = new Map<string, Expression>();
        const positional: Expression[] = [];
        const ordered: Array<{ name?: string; expression: Expression }> = [];
        let namedSeen = false;
        for (const argExpr of argExprs) {
            if (!argExpr) continue;
            if (argExpr.type === 'NamedArgument') {
                const namedArg = argExpr as NamedArgument;
                const nameLower = namedArg.name.toLowerCase();
                if (named.has(nameLower)) {
                    this.throwVbaError(
                        VbaErrorCode.WRONG_NUMBER_OF_ARGUMENTS,
                        `Duplicate named argument: '${namedArg.name}'`,
                    );
                }
                namedSeen = true;
                named.set(nameLower, namedArg.value);
                ordered.push({ name: nameLower, expression: namedArg.value });
            } else {
                if (namedSeen) {
                    this.throwVbaError(
                        VbaErrorCode.WRONG_NUMBER_OF_ARGUMENTS,
                        'Positional argument cannot follow a named argument',
                    );
                }
                positional.push(argExpr);
                ordered.push({ expression: argExpr });
            }
        }
        return { named, positional, ordered };
    }

    /** Evaluate positional argument expressions in their original order. */
    private evaluateCallArgumentValues(argExprs: Expression[]): any[] {
        return argExprs.map(argExpr =>
            this.resolveAutoInstance(argExpr, this.evaluateExpression(argExpr)));
    }

    /** Evaluate an expression list left-to-right without additional coercion. */
    private evaluateExpressions(expressions: Expression[]): any[] {
        return expressions.map(expression => this.evaluateExpression(expression));
    }

    /** Evaluate index expressions without changing their existing numeric coercion. */
    private evaluateIndexExpressions(expressions: Expression[]): number[] {
        return expressions.map(expression => this.evaluateExpression(expression) as number);
    }

    /** Read a multi-dimensional VBA array while preserving caller-specific policies. */
    private readArrayAtIndexes(
        array: any[],
        indexes: number[],
        dims: { lower: number; upper: number }[] | undefined,
        undefinedAsEmpty: boolean,
        validateArity: boolean,
    ): any {
        if (validateArity) this.validateArrayArity(dims, indexes.length);
        let current: any = array;
        for (let i = 0; i < indexes.length; i++) {
            if (!Array.isArray(current)) {
                this.throwVbaError(VbaErrorCode.SUBSCRIPT_OUT_OF_RANGE, 'Subscript out of range');
            }
            const index = indexes[i];
            this.validateArrayIndex(dims, i, index);
            current = current[index];
        }
        // `As New` object arrays store a separate lazy placeholder for each
        // element. Materialize the selected element once and retain it in the
        // array, otherwise `items(0).Value` would lose prior assignments.
        if (isAutoInstancePlaceholder(current) && (array as any).__vbaElementAutoNew__) {
            const instance = this.instantiateClass(current.__className__);
            let owner: any = array;
            for (const index of indexes.slice(0, -1)) owner = owner[index];
            owner[indexes[indexes.length - 1]] = instance;
            current = instance;
        }
        return current === undefined && undefinedAsEmpty ? vbaEmpty : current;
    }

    private alignProcedureCallExpressions(
        proc: ProcedureDeclaration,
        argExprs: (Expression | null)[],
        evaluatedArgs?: any[],
    ): { args: any[]; references: Array<VbaLValueReference | null>; expressions: Array<Expression | null> } {
        const split = this.splitArgumentExpressions(argExprs);
        const paramArrayIndex = proc.parameters.findIndex(p => p.isParamArray);

        // NamedArgument itself is a wrapper, not an assignable expression. Keep
        // its value node so ByRef writes target the original caller variable.
        // ParamArray is different from ordinary parameters: every positional
        // expression from its slot onward must remain an individual argument
        // so bindProcedureParameters can perform the canonical packing.
        const supplied: Array<{ expr: Expression | null; value: any } | undefined> =
            new Array(paramArrayIndex >= 0 ? paramArrayIndex : proc.parameters.length);
        const paramArraySupplied: Array<{ expr: Expression; value: any }> = [];
        let nextPositional = 0;
        for (let sourceIndex = 0; sourceIndex < split.ordered.length; sourceIndex++) {
            const entry = split.ordered[sourceIndex];
            let paramIndex: number;
            if (entry.name !== undefined) {
                paramIndex = proc.parameters.findIndex(p => p.name.toLowerCase() === entry.name);
                if (paramIndex < 0) this.throwVbaError(448, `Named argument not found: '${entry.name}'`);
                if (paramIndex === paramArrayIndex) {
                    this.throwVbaError(448, `Named argument not found: '${entry.name}'`);
                }
            } else {
                while (nextPositional < (paramArrayIndex >= 0 ? paramArrayIndex : supplied.length)
                    && supplied[nextPositional]) nextPositional++;
                if (paramArrayIndex >= 0 && nextPositional >= paramArrayIndex) {
                    const valueExpr = entry.expression;
                    const value = evaluatedArgs?.[sourceIndex] ??
                        this.resolveAutoInstance(valueExpr, this.evaluateExpression(valueExpr));
                    paramArraySupplied.push({ expr: valueExpr, value });
                    continue;
                }
                paramIndex = nextPositional++;
            }
            if (paramIndex >= supplied.length) {
                this.throwVbaError(VbaErrorCode.WRONG_NUMBER_OF_ARGUMENTS, 'Wrong number of arguments or invalid property assignment');
            }
            const valueExpr = entry.expression;
            const value = evaluatedArgs?.[sourceIndex] ??
                this.resolveAutoInstance(valueExpr, this.evaluateExpression(valueExpr));
            supplied[paramIndex] = { expr: valueExpr, value };
        }
        let lastProvided = -1;
        for (let i = supplied.length - 1; i >= 0; i--) {
            if (supplied[i]) { lastProvided = i; break; }
        }
        const aligned = supplied.slice(0, lastProvided + 1);
        const expressions = aligned.map(slot => slot?.expr ?? null);
        // A ByVal argument is already evaluated above.  Resolving a reference
        // for it would evaluate member-call expressions a second time (for
        // example, TextStream.ReadLine would consume two lines).
        const references = expressions.map((expr, i) =>
            this.argumentReference(expr ?? undefined, proc.parameters[i])
        );
        const args = Array.from(aligned, (slot, i) => slot
            ? (references[i] ? references[i]!.get() : slot.value)
            : vbaOmitted);
        if (paramArrayIndex >= 0) {
            for (const slot of paramArraySupplied) {
                expressions.push(slot.expr);
                references.push(this.argumentReference(slot.expr, proc.parameters[paramArrayIndex]));
                const i = expressions.length - 1;
                args.push(references[i] ? references[i]!.get() : slot.value);
            }
        }
        return { args, references, expressions };
    }

    /** Use the common reference writeback path for already-evaluated values. */
    private callClassMethodWithValueReferences(instance: any, proc: ProcedureDeclaration, values: any[]): any {
        const references: VbaLValueReference[] = values.map((_, i) => ({
            get: () => values[i],
            set: (value: any) => { values[i] = value; },
        }));
        const byRefValues: any[] = [];
        const result = this.callClassMethod(instance, proc, values, byRefValues);
        for (let i = 0; i < proc.parameters.length && i < references.length; i++) {
            if (!proc.parameters[i].isByVal) references[i].set(byRefValues[i]);
        }
        return result;
    }

    private callClassMethod(instance: any, proc: ProcedureDeclaration, args: any[], byRefWriteback?: any[]): any {
        const instanceEnv = instance.__instanceEnv__ as Environment;
        const localEnv = new Environment(instanceEnv);

        // Make Me available as the current instance
        localEnv.setLocally('Me', instance);

        // Validate argument count
        this.checkArgCount(proc, args);

        this.bindProcedureParameters(proc, args, localEnv, undefined, !proc.isProperty);

        if (proc.isFunction || (proc.isProperty && proc.propertyType === 'get')) {
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
                else if (this.classDefinitions.has(proc.returnType.toLowerCase())) {
                    localEnv.setVariableType(proc.name.name, {
                        vbaType: 'Object',
                        objectTypeName: proc.returnType,
                    });
                }
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
        const previousProcedureReturnType = this.currentProcedureReturnType;
        const previousProcedureReturnsArray = this.currentProcedureReturnsArray;
        const previousProcIsStatic = this.currentProcIsStatic;
        const previousStaticVars = this.staticVarsInCurrentProc;
        const previousExecutingModule = this.executingModuleName;
        this.env = localEnv;
        this.errorHandlerLabel = null;
        this.errorHandlingMode = 'None';
        this.isInErrorHandler = false;
        this.lastErrorIndex = null;
        this.currentProcBody = proc.body;
        this.currentProcedureName = proc.name.name;
        this.currentProcedureType = proc.propertyType || (proc.isFunction ? 'function' : 'sub');
        this.currentProcedureReturnType = proc.returnType;
        this.currentProcedureReturnsArray = proc.returnsArray ?? false;
        this.currentProcIsStatic = false;
        this.staticVarsInCurrentProc = new Set();
        this.executingModuleName = proc.moduleName ?? previousExecutingModule;

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
            // Event handlers pass a mutable argument array so ByRef changes can
            // propagate back through RaiseEvent.
            if (byRefWriteback) {
                for (let i = 0; i < proc.parameters.length; i++) {
                    const param = proc.parameters[i];
                    if (!param.isByVal) {
                        byRefWriteback[i] = localEnv.get(param.name);
                    }
                }
            }
            this.env = previousEnv;
            this.errorHandlerLabel = previousErrorHandler;
            this.errorHandlingMode = previousErrorHandlingMode;
            this.isInErrorHandler = previousIsInErrorHandler;
            this.lastErrorIndex = previousLastErrorIndex;

            // Scope exit: release references held by local variables (including params).
            // Skip `me` (not addRef'd on entry). The retVar object (retCapture) is released
            // like any local, but Terminate is suppressed — even if a sibling local var holds
            // the same object and its release brings refCount to 0.
            {
                const retVarName = (proc.isFunction || (proc.isProperty && proc.propertyType === 'get'))
                    ? proc.name.name.toLowerCase() : null;
                const retCapture = retVarName !== null ? localEnv.getLocalVariables().get(retVarName) : undefined;
                for (const [name, val] of localEnv.getLocalVariables()) {
                    if (name === 'me') continue;
                    if (val && val.__vbaClass__) {
                        val.__refCount__ = (val.__refCount__ ?? 0) - 1;
                        if (val.__refCount__ <= 0 && val !== retCapture) {
                            this.triggerTerminate(val);
                        }
                    }
                }
            }

            this.currentProcBody = previousProcBody;
            this.currentProcedureName = previousProcedureName;
            this.currentProcedureType = previousProcedureType;
            this.currentProcedureReturnType = previousProcedureReturnType;
            this.currentProcedureReturnsArray = previousProcedureReturnsArray;
            this.currentProcIsStatic = previousProcIsStatic;
            this.staticVarsInCurrentProc = previousStaticVars;
            this.executingModuleName = previousExecutingModule;
        }

        if (proc.isFunction || (proc.isProperty && proc.propertyType === 'get')) {
            const result = localEnv.get(proc.name.name.toLowerCase());
            if (proc.returnsArray && Array.isArray(result)) (result as any).__vbaArrayReturn__ = true;
            return result;
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
        // VBA identifiers are case-insensitive.  Perform this pass before
        // resolving constants so collisions are reported as compile errors,
        // rather than whichever declaration happened to be loaded last.
        for (const { ast, moduleName } of modules) {
            this.checkDuplicateDeclarationsInScope(ast.body, moduleName);
        }
        // 全モジュールレベル定数を「module:name」修飾キーで収集する。
        // 1 ConstDeclaration に複数の declarator が含まれる場合も個別にエントリを作る。
        const allConsts = new Map<string, { stmt: ConstDeclaration; decl: ConstDeclaratorItem; moduleName: string; name: string }>();
        const collect = (stmt: ConstDeclaration, moduleName: string) => {
            for (const decl of stmt.declarations) {
                const name = decl.name.name.toLowerCase();
                allConsts.set(`${moduleName.toLowerCase()}:${name}`, { stmt, decl, moduleName, name });
            }
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
        for (const [qkey, { decl, moduleName }] of allConsts) {
            const exprDeps = this.collectConstExprDeps(decl.value);
            const resolved = new Set<string>();
            for (const d of exprDeps) {
                const r = resolveDep(d, moduleName);
                if (r) {
                    // 他モジュールの Private Const への参照はエラー
                    const { moduleName: depModule, stmt: depStmt, decl: depDecl } = allConsts.get(r)!;
                    if (depModule.toLowerCase() !== moduleName.toLowerCase() &&
                        (depStmt as any).scope === 'private') {
                        throw new Error(
                            `Constant expression required: '${decl.name.name}' references private constant '${depDecl.name.name}' in module '${depModule}'`
                        );
                    }
                    resolved.add(r);
                }
            }
            deps.set(qkey, resolved);
        }

        // グローバルトポロジカルソート（循環参照を検出）
        const order = this.topologicalSortConsts(allConsts, deps);

        // 正しい順序で評価（各 declarator を個別に評価）
        for (const qkey of order) {
            const { decl, moduleName } = allConsts.get(qkey)!;
            const prev = this.currentSourceModule;
            this.currentSourceModule = moduleName;
            this.evaluateOneConst(decl);
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
        for (const decl of stmt.declarations) {
            this.evaluateOneConst(decl);
        }
    }

    private evaluateOneConst(decl: ConstDeclaratorItem) {
        const value = this.evaluateConstValue(decl.value);
        const name = decl.name.name;
        if (!this.currentProcedureName && this.currentSourceModule) {
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
    private detachWithEventsHandlers(fieldName: string, instance: any): void {
        const bindings = instance?.__withEventsBindings__ as Map<string, { source: any, eventName: string, handler: (...args: any[]) => void }[]> | undefined;
        const fieldKey = fieldName.toLowerCase();
        const entries = bindings?.get(fieldKey);
        if (!entries) return;

        for (const { source, eventName, handler } of entries) {
            const handlers = source?.__events__?.get(eventName) as ((...args: any[]) => void)[] | undefined;
            if (!handlers) continue;
            const index = handlers.indexOf(handler);
            if (index >= 0) handlers.splice(index, 1);
        }
        bindings!.delete(fieldKey);
    }

    private bindWithEventsHandlers(fieldName: string, value: any, classDef: ClassDeclaration | undefined, instance: any) {
        this.detachWithEventsHandlers(fieldName, instance);
        const fieldKey = fieldName.toLowerCase();
        const bindings: Map<string, { source: any, eventName: string, handler: (...args: any[]) => void }[]> | undefined = instance
            ? (instance.__withEventsBindings__ ??= new Map())
            : undefined;
        for (const eventName of value.__events__.keys()) {
            const handlerName = `${fieldName}_${eventName}`;
            const handlerNameLower = handlerName.toLowerCase();
            const handlers = value.__events__.get(eventName);
            let eventHandler: ((...args: any[]) => void) | undefined;
            if (classDef) {
                const classProc = classDef.procedures.find(
                    p => p.name.name.toLowerCase() === handlerNameLower
                );
                if (classProc) {
                    const capturedInstance = instance;
                    const capturedProc = classProc;
                    eventHandler = (...args: any[]) => {
                        this.callClassMethodWithValueReferences(capturedInstance, capturedProc, args);
                        return args;
                    };
                    (eventHandler as any).__vbaEventByRef__ = true;
                }
            }
            if (!eventHandler) {
                const handler = this.env.getProcedure(handlerName);
                if (handler) {
                    eventHandler = (...args: any[]) => {
                        this.callProcedure(handlerName, args);
                        return args;
                    };
                    (eventHandler as any).__vbaEventByRef__ = true;
                }
            }
            if (eventHandler) {
                handlers.push(eventHandler);
                const entries = bindings?.get(fieldKey) ?? [];
                entries.push({ source: value, eventName, handler: eventHandler });
                bindings?.set(fieldKey, entries);
            }
        }
    }

    private evaluateSetStatement(stmt: SetStatement) {
        let value = this.evaluateExpression(stmt.right);
        value = this.resolveAutoInstance(stmt.right, value);

        // VBA requires Set target to be an object (or Nothing)
        if (!isVbaObjectReferenceCompatible(value)) {
            this.throwVbaError(VbaErrorCode.OBJECT_REQUIRED, 'Object required');
        }
        // If the right side evaluates to a variable name (string), resolve it
        if (typeof value === 'string' && stmt.right.type === 'Identifier') {
            value = this.env.get(value);
        }

        if (stmt.left.type === 'Identifier') {
            const name = (stmt.left as Identifier).name;
            const procNameLower = name.toLowerCase();

            // Special case: assigning to current function/property name (return value).
            // Unlike B-3 for regular vars, the retVar slot owns exactly one ref at a time:
            // addRef the new value, releaseRef the old. The scope exit releases this ref
            // without calling Terminate (it's the handoff to the caller's Set statement).
            if (this.currentProcedureName && this.currentProcedureName.toLowerCase() === procNameLower) {
                if (this.currentProcedureType === 'function' || this.currentProcedureType === 'get') {
                    const oldRetVal = this.env.get(name);
                    this.addRef(value);
                    this.releaseRef(oldRetVal);
                    this.env.setLocally(name, value);
                    return;
                }
            }

            let oldVal = this.env.get(name);
            this.rejectTypedArrayWholeAssignment(oldVal, value, stmt.right);
            // `Dim w As New T` + `Set w = Nothing` without any prior access: real VBA still
            // instantiates the object (Class_Initialize) before destroying it (Class_Terminate).
            // Resolve the placeholder first so both lifecycle hooks fire correctly.
            if (value === vbaNothing && isAutoInstancePlaceholder(oldVal)
                && !(oldVal as any).__explicitlyCleared__) {
                oldVal = this.instantiateClass(oldVal.__className__);
                this.env.set(name, oldVal);
            }
            const isWithEvents = this.env.isWithEvents(name);
            const me = this.env.get('me');
            if (isWithEvents) {
                this.detachWithEventsHandlers(name, me);
                // A WithEvents field owns its event subscription and its object reference.
                // Replacing or clearing it must release the old source independently of
                // any references that other variables retain.
                if (oldVal !== value) {
                    this.addRef(value);
                    this.releaseRef(oldVal);
                }
            } else {
                // addRef new value; releaseRef old value only on explicit Nothing (B-3: don't
                // release on reassignment because containers like Dictionary may still hold the ref).
                this.addRef(value);
                if (value === vbaNothing && oldVal !== value) this.releaseRef(oldVal);
            }

            const proc = this.env.getProcedure(name, 'set');
            if (proc) {
                this.callProcedure(name, [value], 'set');
                return;
            }

            // Auto-Instantiation: `Set x = Nothing` で auto-instance 変数なら placeholder に戻す。
            // 仕様: 再度参照したら新しいインスタンスが生成される。
            const className = this.autoInstanceVars.get(name.toLowerCase());
            if (className && value === vbaNothing) {
                const placeholder = createAutoInstancePlaceholder(className);
                // Remember that this auto-instance was explicitly cleared.
                // A repeated `Set x = Nothing` must not materialize and
                // terminate a fresh, unused object a second time.  Accessing
                // x still resolves the placeholder normally and then a later
                // clear gets the usual initialize/terminate lifecycle.
                (placeholder as any).__explicitlyCleared__ = true;
                this.env.set(name, placeholder);
                return;
            }
            this.env.set(name, value);

            // Check for WithEvents binding
            if (isWithEvents && value && value.__events__) {
                // Binding: look for VarName_EventName subs — first in class scope, then global
                this.bindWithEventsHandlers(name, value, me?.__classDef__, me);
            }
        } else if (stmt.left.type === 'MemberExpression') {
            const member = stmt.left as MemberExpression;
            const obj = this.resolveAutoInstance(member.object, this.evaluateExpression(member.object));
            const propName = member.property.name.toLowerCase();
            if (obj && obj.__vbaClass__) {
                const classDef = obj.__classDef__ as ClassDeclaration;
                const instanceEnv = obj.__instanceEnv__ as Environment;
                const oldVal = this.getClassField(instanceEnv, propName);
                this.rejectTypedArrayWholeAssignment(oldVal, value, stmt.right);
                const isWithEvents = instanceEnv.hasOwnWithEvents(propName);
                const setter = classDef.procedures.find(
                    p => p.isProperty && p.propertyType === 'set' && p.name.name.toLowerCase() === propName
                );
                if (setter) {
                    this.callClassMethodWithExpressions(obj, setter, [stmt.right], [value]);
                } else {
                    if (isWithEvents) {
                        this.detachWithEventsHandlers(propName, obj);
                        if (oldVal !== value) {
                            this.addRef(value);
                            this.releaseRef(oldVal);
                        }
                    } else if (value === vbaNothing && oldVal !== value) {
                        this.triggerTerminate(oldVal);
                    }
                    instanceEnv.setLocally(propName, value);
                    // obj.Field = New X 形式（外部からの WithEvents フィールド代入）でも
                    // バインドが行われるようにする。bare identifier 代入のみを見ていたため、
                    // member access 経由の代入では handler が一切ワイヤリングされず
                    // イベントが静かに発火しなかった。
                    if (isWithEvents && value && value.__events__) {
                        this.bindWithEventsHandlers(propName, value, classDef, obj);
                    }
                }
            } else if (obj && typeof obj === 'object') {
                const oldVal = obj[propName];
                this.rejectTypedArrayWholeAssignment(oldVal, value, stmt.right);
                if (value === vbaNothing && oldVal !== value) this.triggerTerminate(oldVal);
                obj[propName] = value;
            } else {
                if (obj === null || obj === undefined || obj === vbaNothing) {
                    this.throwVbaError(VbaErrorCode.OBJECT_VARIABLE_NOT_SET, 'Object variable or With block variable not set');
                } else {
                    this.throwVbaError(VbaErrorCode.OBJECT_REQUIRED, 'Object required');
                }
            }
        } else if (stmt.left.type === 'ImplicitWithObjectExpression') {
            if (this.withObjectStack.length === 0) {
                this.throwVbaError(VbaErrorCode.OBJECT_VARIABLE_NOT_SET, 'Object variable or With block variable not set');
            }
            const obj = this.withObjectStack[this.withObjectStack.length - 1];
            const propName = (stmt.left as ImplicitWithObjectExpression).property.name.toLowerCase();
            if (obj && obj.__vbaClass__) {
                const classDef = obj.__classDef__ as ClassDeclaration;
                const instanceEnv = obj.__instanceEnv__ as Environment;
                this.rejectTypedArrayWholeAssignment(this.getClassField(instanceEnv, propName), value, stmt.right);
                const setter = classDef.procedures.find(
                    p => p.isProperty && p.propertyType === 'set' && p.name.name.toLowerCase() === propName
                );
                if (setter) {
                    this.callClassMethodWithExpressions(obj, setter, [stmt.right], [value]);
                } else {
                    instanceEnv.setLocally(propName, value);
                }
            } else if (obj && typeof obj === 'object') {
                this.rejectTypedArrayWholeAssignment(obj[propName], value, stmt.right);
                obj[propName] = value;
            } else {
                this.throwVbaError(VbaErrorCode.OBJECT_REQUIRED, 'Object required');
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
                    const instanceEnv = obj.__instanceEnv__ as Environment;
                    const fieldArr = this.getClassField(instanceEnv, methodName);
                    if (Array.isArray(fieldArr)) {
                        const requiredType = (fieldArr as any).__vbaElementObjectTypeName__ as string | undefined;
                        if (requiredType) this.validateObjectArrayElementValue(value, requiredType);
                        const dims = (fieldArr as any).__vbaDimensions__ as { lower: number, upper: number }[] | undefined;
                        if (dims && call.args.length !== dims.length) {
                            this.throwVbaError(VbaErrorCode.SUBSCRIPT_OUT_OF_RANGE, 'Subscript out of range');
                        }
                        let current = fieldArr;
                        for (let i = 0; i < call.args.length - 1; i++) {
                            const index = this.evaluateArrayIndex(call.args[i], dims, i);
                            current = current[index];
                        }
                        const lastIndex = this.evaluateArrayIndex(call.args[call.args.length - 1], dims, call.args.length - 1);
                        current[lastIndex] = value;
                        return;
                    }
                    const setter = classDef.procedures.find(
                        p => p.isProperty && p.propertyType === 'set' && p.name.name.toLowerCase() === methodName
                    );
                    if (setter) {
                        const argsVals = this.evaluateExpressions(call.args);
                        this.callClassMethodWithExpressions(obj, setter, [...call.args, stmt.right], [...argsVals, value]);
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
                } else if (Array.isArray(target) && (target as any).__vbaElementObjectTypeName__) {
                    const requiredType = (target as any).__vbaElementObjectTypeName__ as string;
                    this.validateObjectArrayElementValue(value, requiredType);
                    const dims = (target as any).__vbaDimensions__ as { lower: number, upper: number }[] | undefined;
                    this.validateArrayArity(dims, call.args.length);
                    let current: any = target;
                    for (let i = 0; i < call.args.length - 1; i++) {
                        const index = this.evaluateExpression(call.args[i]) as number;
                        this.validateArrayIndex(dims, i, index);
                        current = current[index];
                        if (!Array.isArray(current)) {
                            this.throwVbaError(VbaErrorCode.SUBSCRIPT_OUT_OF_RANGE, 'Subscript out of range');
                        }
                    }
                    const lastIndex = this.evaluateExpression(call.args[call.args.length - 1]) as number;
                    this.validateArrayIndex(dims, call.args.length - 1, lastIndex);
                    current[lastIndex] = value;
                } else if (target && typeof target === 'object') {
                    const key = String(this.evaluateExpression(call.args[0]));
                    target[key] = value;
                } else {
                    this.throwVbaError(VbaErrorCode.INVALID_PROCEDURE_CALL, 'Invalid procedure call or argument');
                }
            } else if (call.callee.type === 'ImplicitWithObjectExpression') {
                if (this.withObjectStack.length === 0) {
                    this.throwVbaError(VbaErrorCode.OBJECT_VARIABLE_NOT_SET, 'Object variable or With block variable not set');
                }
                const obj = this.withObjectStack[this.withObjectStack.length - 1];
                const propName = (call.callee as ImplicitWithObjectExpression).property.name.toLowerCase();
                if (obj && obj.__vbaClass__) {
                    const instanceEnv = obj.__instanceEnv__ as Environment;
                    const fieldArr = this.getClassField(instanceEnv, propName);
                    if (Array.isArray(fieldArr)) {
                        const requiredType = (fieldArr as any).__vbaElementObjectTypeName__ as string | undefined;
                        if (requiredType) this.validateObjectArrayElementValue(value, requiredType);
                        const dims = (fieldArr as any).__vbaDimensions__ as { lower: number, upper: number }[] | undefined;
                        if (dims && call.args.length !== dims.length) {
                            this.throwVbaError(VbaErrorCode.SUBSCRIPT_OUT_OF_RANGE, 'Subscript out of range');
                        }
                        let current = fieldArr;
                        for (let i = 0; i < call.args.length - 1; i++) {
                            const index = this.evaluateArrayIndex(call.args[i], dims, i);
                            current = current[index];
                        }
                        const lastIndex = this.evaluateArrayIndex(call.args[call.args.length - 1], dims, call.args.length - 1);
                        current[lastIndex] = value;
                        return;
                    }
                    const classDef = obj.__classDef__ as ClassDeclaration;
                    const setter = classDef.procedures.find(
                        p => p.isProperty && p.propertyType === 'set' && p.name.name.toLowerCase() === propName
                    );
                    if (setter) {
                        this.callClassMethodWithExpressions(obj, setter, [...call.args, stmt.right],
                            [...this.evaluateExpressions(call.args), value]);
                    } else {
                        this.throwVbaError(VbaErrorCode.OBJECT_DOESNT_SUPPORT_PROPERTY, "Object doesn't support this property or method");
                    }
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
        // Err.Clear resets Err's visible fields, but leaves the active GoTo
        // handler in place. Resume remains valid until that handler completes.
        if (!this.isInErrorHandler) {
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
        const fileNum = this.evaluateFileNumber(stmt.fileNumber);

        if (this.fileHandles.has(fileNum)) {
            this.throwVbaError(VbaErrorCode.FILE_ALREADY_OPEN, "File already open");
        }
        const access = stmt.access ?? this.defaultOpenAccess(stmt.mode);
        const lock = stmt.lock ?? 'Lock Read Write';
        for (const h of this.fileHandles.values()) {
            if (h.path !== realPath) continue;
            if (this.openLockDenies(h.lock, access) || this.openLockDenies(lock, h.access)) {
                this.throwVbaError(VbaErrorCode.PERMISSION_DENIED, "Permission denied");
            }
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
                    if (dir === this.sandbox.getRoot()) {
                        this.fs.mkdirSync(dir, { recursive: true });
                    } else {
                        this.throwVbaError(VbaErrorCode.PATH_NOT_FOUND, 'Path not found');
                    }
                }
            } else if (flags === 'r' && !this.fs.existsSync(realPath)) {
                this.throwVbaError(VbaErrorCode.FILE_NOT_FOUND, "File not found");
            }

            const fd = this.fs.openSync(realPath, flags);
            let recordLen: number | undefined;
            if (stmt.mode === 'Random') {
                recordLen = stmt.recordLen ? Number(this.evaluateExpression(stmt.recordLen)) : 128;
                if (!Number.isInteger(recordLen) || recordLen < 1) {
                    this.throwVbaError(VbaErrorCode.INVALID_PROCEDURE_CALL, 'Invalid procedure call or argument');
                }
            }
            const initialPosition = stmt.mode === 'Append'
                ? this.fs.statSync(realPath).size
                : 0;
            this.fileHandles.set(fileNum, {
                fd,
                fileNumber: fileNum,
                mode: stmt.mode,
                path: realPath,
                access,
                lock,
                pos: initialPosition,
                lastRecord: 0,
                nextRecord: 1,
                eof: stmt.mode === 'Output' || stmt.mode === 'Append',
                width: 0,
                lineColumn: 0,
                recordLen,
                locks: []
            });
        } catch (e: any) {
            if (e.code === 'ENOENT') this.throwVbaError(VbaErrorCode.FILE_NOT_FOUND, "File not found");
            if (e.code === 'EACCES') this.throwVbaError(VbaErrorCode.PATH_FILE_ACCESS_ERROR, "Path/File access error");
            throw e;
        }
    }

    private defaultOpenAccess(mode: OpenStatement['mode']): 'Read' | 'Write' | 'Read Write' {
        if (mode === 'Input') return 'Read';
        if (mode === 'Output' || mode === 'Append') return 'Write';
        return 'Read Write';
    }

    private openLockDenies(lock: 'Shared' | 'Lock Read' | 'Lock Write' | 'Lock Read Write', access: 'Read' | 'Write' | 'Read Write'): boolean {
        if (lock === 'Shared') return false;
        const reads = access === 'Read' || access === 'Read Write';
        const writes = access === 'Write' || access === 'Read Write';
        return (reads && (lock === 'Lock Read' || lock === 'Lock Read Write'))
            || (writes && (lock === 'Lock Write' || lock === 'Lock Read Write'));
    }

    private evaluateCloseStatement(stmt: CloseStatement) {
        const nums = stmt.fileNumbers.length > 0
            ? stmt.fileNumbers.map(n => this.evaluateFileNumber(n))
            : Array.from(this.fileHandles.keys());

        for (const num of nums) {
            const handle = this.fileHandles.get(num);
            if (handle) {
                this.releaseFileLocks(handle);
                this.fs.closeSync(handle.fd);
                this.fileHandles.delete(num);
            }
        }
    }

    private evaluatePrintStatement(stmt: PrintStatement) {
        const fileNum = this.evaluateFileNumber(stmt.fileNumber);
        const handle = this.fileHandles.get(fileNum);
        if (!handle) this.throwVbaError(VbaErrorCode.BAD_FILE_NAME_OR_NUMBER, "Bad file name or number");
        this.requireFileMode(handle, ['Output', 'Append']);

        let output = "";
        const startingColumn = handle.lineColumn ?? 0;
        for (const expr of stmt.expressions) {
            if (expr === 'Comma') {
                const currentLen = startingColumn + output.length;
                const target = Math.ceil((currentLen + 1) / 14) * 14;
                output += " ".repeat(Math.max(0, target - currentLen));
            } else if (expr === 'Semicolon') {
                // Continue
            } else if (typeof expr === 'object' && expr !== null && 'type' in expr) {
                 if (expr.type === 'Spc') {
                     const n = Number(this.evaluateExpression((expr as any).val));
                     output += " ".repeat(Math.max(0, n));
                } else if (expr.type === 'Tab') {
                    // Tab(n): 次の出力を n 桁目（1 始まり）から始める → 長さ n-1 まで空白
                    const n = Number(this.evaluateExpression((expr as any).val));
                    output += " ".repeat(Math.max(0, n - 1 - startingColumn - output.length));
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

        // Width # controls automatic line breaks for Print #. Width 0 means
        // unlimited output. Track the current line across semicolon-terminated
        // Print statements so repeated calls wrap at the configured boundary.
        const width = handle.width ?? 0;
        if (width > 0) {
            let wrapped = '';
            let column = handle.lineColumn ?? 0;
            for (let i = 0; i < output.length; i++) {
                const ch = output[i];
                if (ch === '\r' && output[i + 1] === '\n') {
                    wrapped += '\r\n';
                    column = 0;
                    i++;
                    continue;
                }
                if (column >= width) {
                    wrapped += '\r\n';
                    column = 0;
                }
                wrapped += ch;
                column++;
            }
            output = wrapped;
            handle.lineColumn = column;
        } else if (output.endsWith('\r\n')) {
            handle.lineColumn = 0;
        } else {
            handle.lineColumn = (handle.lineColumn ?? 0) + output.length;
        }

        const outputBytes = this.encodeVbaFileText(output);
        const writePosition = handle.mode === 'Append' ? null : (handle.pos ?? null);
        this.fs.writeSync(handle.fd, outputBytes, 0, outputBytes.length, writePosition);
        handle.pos = handle.mode === 'Append'
            ? this.fs.statSync(handle.path).size
            : handle.pos! + outputBytes.length;
    }

    private evaluateDebugPrintStatement(stmt: DebugPrintStatement) {
        let output = "";
        for (const expr of stmt.expressions) {
            if (expr === 'Comma') {
                const currentLen = output.length;
                const target = Math.ceil((currentLen + 1) / 14) * 14;
                output += " ".repeat(target - currentLen);
            } else if (expr === 'Semicolon') {
                // no separator
            } else if (typeof expr === 'object' && expr !== null && 'type' in expr) {
                if (expr.type === 'Spc') {
                    const n = Number(this.evaluateExpression((expr as any).val));
                    output += " ".repeat(Math.max(0, n));
                } else if (expr.type === 'Tab') {
                    // Tab(n): 次の出力を n 桁目（1 始まり）から始める → 長さ n-1 まで空白
                    const n = Number(this.evaluateExpression((expr as any).val));
                    output += " ".repeat(Math.max(0, n - 1 - output.length));
                } else {
                    output += this.toDisplayString(this.evaluateExpression(expr as any));
                }
            } else {
                output += this.toDisplayString(this.evaluateExpression(expr as any));
            }
        }
        this.onPrint(output);
    }

    private evaluateLineInputStatement(stmt: LineInputStatement) {
        const fileNum = this.evaluateFileNumber(stmt.fileNumber);
        const handle = this.fileHandles.get(fileNum);
        if (!handle) this.throwVbaError(VbaErrorCode.BAD_FILE_NAME_OR_NUMBER, "Bad file name or number");
        this.requireFileMode(handle, ['Input', 'Binary']);

        const buffer = new Uint8Array(1);
        const lineBytes: number[] = [];
        let bytesRead = 0;

        while (true) {
            bytesRead = this.fs.readSync(handle.fd, buffer, 0, 1, handle.pos ?? null);
            if (bytesRead === 0) {
                if (lineBytes.length === 0) {
                    handle.eof = true;
                    this.throwVbaError(VbaErrorCode.INPUT_PAST_END_OF_FILE, 'Input past end of file');
                }
                break;
            }
            const byte = buffer[0];
            handle.pos!++;
            if (byte === 0x0a) break;
            if (byte !== 0x0d) lineBytes.push(byte);
        }

        const line = this.decodeVbaFileText(new Uint8Array(lineBytes));
        this.evaluateAssignmentToVariable(stmt.variable, line);
    }

    private evaluatePutStatement(stmt: PutStatement) {
        const fileNum = this.evaluateFileNumber(stmt.fileNumber);
        const handle = this.fileHandles.get(fileNum);
        if (!handle) this.throwVbaError(VbaErrorCode.BAD_FILE_NAME_OR_NUMBER, "Bad file name or number");
        this.requireFileMode(handle, ['Random', 'Binary']);

        const data = this.evaluateExpression(stmt.data);
        const layout = this.getBinaryValueLayout(stmt.data, data);
        const buffer = this.encodeBinaryValue(data, layout);

        const position = this.resolveFileRecordPosition(handle, stmt.recordNumber);

        this.fs.writeSync(handle.fd, buffer, 0, buffer.length, position);
        handle.pos = (position ?? handle.pos ?? 0) + buffer.length;
        if (handle.mode === 'Random') {
            handle.lastRecord = Math.floor((position ?? handle.pos ?? 0) / (handle.recordLen ?? 128)) + 1;
            handle.nextRecord = handle.lastRecord + 1;
            handle.eof = false;
        } else if (handle.mode === 'Binary') {
            handle.eof = false;
        }
    }

    /** VBA binary file strings use the active ANSI code page.  The runner's
     * supported Windows-compatible code page is CP932, matching the project
     * fixtures and the Excel differential tests. */
    private static readonly VBA_BINARY_ENCODING = 'cp932';

    /** Shared text-file encoding boundary for Print/Write/Input/Line Input. */
    private encodeVbaFileText(text: string): Uint8Array {
        return iconv.encode(text, Evaluator.VBA_BINARY_ENCODING);
    }

    private decodeVbaFileText(bytes: Uint8Array): string {
        return iconv.decode(Buffer.from(bytes), Evaluator.VBA_BINARY_ENCODING);
    }

    private getBinaryValueLayout(expr: Expression, value: any): BinaryValueLayout {
        const elementType = Array.isArray(value) ? (value as any).__vbaElementType__?.toLowerCase() : undefined;
        const elementTypeName = Array.isArray(value) ? (value as any).__vbaElementTypeName__ as string | undefined : undefined;
        if (elementTypeName && this.env.getType(elementTypeName)) {
            return { typeName: 'Array', elementType: elementTypeName };
        }
        if (elementType && [
            'byte', 'integer', 'boolean', 'long', 'longlong', 'longptr',
            'single', 'double', 'date', 'currency',
        ].includes(elementType)) {
            return { typeName: 'Array', elementType };
        }
        if (elementType === 'string' && (value as any).__vbaElementFixedLength__ !== undefined) {
            return {
                typeName: 'Array',
                elementType,
                fixedLength: (value as any).__vbaElementFixedLength__,
            };
        }
        if (expr.type === 'Identifier') {
            const typeInfo = this.env.getVariableType((expr as Identifier).name);
            if (typeInfo && typeInfo.vbaType !== 'Variant') {
                return { typeName: typeInfo.vbaType, fixedLength: typeInfo.fixedLength };
            }
        }
        if (expr.type === 'CallExpression' && (expr as CallExpression).callee.type === 'Identifier') {
            const name = ((expr as CallExpression).callee as Identifier).name.toLowerCase();
            const conversionTypes: Record<string, string> = {
                cbyte: 'Byte', cint: 'Integer', clng: 'Long',
            };
            if (conversionTypes[name]) return { typeName: conversionTypes[name] };
        }
        if (expr.type === 'NumberLiteral') {
            const n = Number(value);
            return { typeName: Number.isInteger(n) && n >= -32768 && n <= 32767 ? 'Integer' : 'Long' };
        }
        if (value && typeof value === 'object' && value.__vbaTypeName__ && this.env.getType(value.__vbaTypeName__)) {
            return { typeName: value.__vbaTypeName__ };
        }
        if (typeof value === 'string') return { typeName: 'String' };
        if (value === vbaEmpty) return { typeName: 'String', readToEnd: true };
        this.throwVbaError(VbaErrorCode.TYPE_MISMATCH, 'Binary Put/Get requires a declared scalar, String, or UDT');
    }

    private encodeBinaryValue(value: any, layout: BinaryValueLayout): Uint8Array {
        const typeName = layout.typeName.toLowerCase();
        switch (typeName) {
            case 'array': {
                if (!layout.elementType) {
                    this.throwVbaError(VbaErrorCode.TYPE_MISMATCH, 'Missing Binary Put/Get array element type');
                }
                const dims = (value as any).__vbaDimensions__ as { lower: number, upper: number }[] | undefined;
                const fields: Uint8Array[] = [];
                const dimensions = dims?.length ?? 1;
                const coordinates: number[] = [];
                const readAt = (array: any, indexes: number[]): any => {
                    let current = array;
                    for (const index of indexes) current = current[index];
                    return current;
                };
                // VBA stores the left-most dimension contiguously: for a(0,0)
                // ... a(1,1), the physical order is (0,0),(1,0),(0,1),(1,1).
                const collect = (dimension: number): void => {
                    const lower = dims?.[dimension].lower ?? (dimension === 0 ? (value as any).vbaBase ?? 0 : 0);
                    const upper = dims?.[dimension].upper ?? (dimension === 0 ? value.length - 1 : value[0].length - 1);
                    for (let i = lower; i <= upper; i++) {
                        coordinates[dimension] = i;
                        if (dimension === 0) fields.push(this.encodeBinaryValue(readAt(value, coordinates), {
                            typeName: layout.elementType!, fixedLength: layout.fixedLength,
                        }));
                        else collect(dimension - 1);
                    }
                };
                collect(dimensions - 1);
                const length = fields.reduce((total, field) => total + field.length, 0);
                const buffer = new Uint8Array(length);
                let offset = 0;
                for (const field of fields) {
                    buffer.set(field, offset);
                    offset += field.length;
                }
                return buffer;
            }
            case 'byte':
                return Uint8Array.of(Number(value) & 0xff);
            case 'integer': {
                const buffer = new Uint8Array(2);
                new DataView(buffer.buffer).setInt16(0, Number(value), true);
                return buffer;
            }
            case 'boolean': {
                const buffer = new Uint8Array(2);
                const booleanValue = value instanceof VbaBoolean ? value.value : (value ? -1 : 0);
                new DataView(buffer.buffer).setInt16(0, booleanValue, true);
                return buffer;
            }
            case 'long': {
                const buffer = new Uint8Array(4);
                new DataView(buffer.buffer).setInt32(0, Number(value), true);
                return buffer;
            }
            case 'longlong':
            case 'longptr': {
                const buffer = new Uint8Array(8);
                new DataView(buffer.buffer).setBigInt64(0, BigInt(value), true);
                return buffer;
            }
            case 'single': {
                const buffer = new Uint8Array(4);
                new DataView(buffer.buffer).setFloat32(0, Number(value), true);
                return buffer;
            }
            case 'double': {
                const buffer = new Uint8Array(8);
                new DataView(buffer.buffer).setFloat64(0, Number(value), true);
                return buffer;
            }
            case 'date': {
                const buffer = new Uint8Array(8);
                new DataView(buffer.buffer).setFloat64(0, value instanceof VbaDate ? value.value : Number(value), true);
                return buffer;
            }
            case 'currency': {
                const buffer = new Uint8Array(8);
                const internal = value instanceof VbaCurrency ? value.internal : VbaCurrency.fromNumber(Number(value)).internal;
                new DataView(buffer.buffer).setBigInt64(0, internal, true);
                return buffer;
            }
            case 'string': {
                const data = iconv.encode(String(value), Evaluator.VBA_BINARY_ENCODING);
                if (!layout.variableLengthDescriptor) return data;
                if (data.length > 0xffff) {
                    this.throwVbaError(VbaErrorCode.OVERFLOW, 'Variable-length UDT String is too long');
                }
                const buffer = new Uint8Array(2 + data.length);
                new DataView(buffer.buffer).setUint16(0, data.length, true);
                buffer.set(data, 2);
                return buffer;
            }
            default: {
                const members = this.env.getType(layout.typeName);
                if (!members || !value || typeof value !== 'object') {
                    this.throwVbaError(VbaErrorCode.TYPE_MISMATCH, `Unsupported binary value type: ${layout.typeName}`);
                }
                const fields = members.map(member => {
                    if (member.isArray) {
                        return this.encodeBinaryUdtArray(value[member.name.toLowerCase()], member);
                    }
                    return this.encodeBinaryValue(value[member.name.toLowerCase()], {
                        typeName: member.memberType,
                        fixedLength: member.fixedLength,
                        variableLengthDescriptor: member.memberType.toLowerCase() === 'string' && member.fixedLength === undefined,
                    });
                });
                const length = fields.reduce((total, field) => total + field.length, 0);
                const buffer = new Uint8Array(length);
                let offset = 0;
                for (const field of fields) {
                    buffer.set(field, offset);
                    offset += field.length;
                }
                return buffer;
            }
        }
    }

    /** Serialize a fixed-size UDT member array in VBA's declared index order. */
    private encodeBinaryUdtArray(value: any, member: TypeMember): Uint8Array {
        const bounds = member.arrayBounds;
        if (!bounds || bounds.length === 0) {
            this.throwVbaError(VbaErrorCode.TYPE_MISMATCH, 'Binary Put/Get requires fixed-size UDT array members');
        }

        const encodeDimension = (array: any, dimension: number): Uint8Array => {
            if (!Array.isArray(array)) {
                this.throwVbaError(VbaErrorCode.TYPE_MISMATCH, 'Invalid UDT array member for Binary Put/Get');
            }
            const fields: Uint8Array[] = [];
            const { lower, upper } = bounds![dimension];
            for (let index = lower; index <= upper; index++) {
                fields.push(dimension < bounds!.length - 1
                    ? encodeDimension(array[index], dimension + 1)
                    : this.encodeBinaryValue(array[index], {
                        typeName: member.memberType,
                        fixedLength: member.fixedLength,
                        variableLengthDescriptor: member.memberType.toLowerCase() === 'string' && member.fixedLength === undefined,
                    }));
            }
            const length = fields.reduce((total, field) => total + field.length, 0);
            const buffer = new Uint8Array(length);
            let offset = 0;
            for (const field of fields) {
                buffer.set(field, offset);
                offset += field.length;
            }
            return buffer;
        };

        return encodeDimension(value, 0);
    }

    private encodedByteLengthForCharacters(bytes: Uint8Array, offset: number, characters: number): number {
        let cursor = offset;
        let count = 0;
        while (cursor < bytes.length && count < characters) {
            const first = bytes[cursor++];
            // CP932 lead-byte ranges. A lead byte without a following byte is
            // consumed as one byte; the decoder will supply its replacement.
            if ((first >= 0x81 && first <= 0x9f) || (first >= 0xe0 && first <= 0xfc)) {
                if (cursor < bytes.length) cursor++;
            }
            count++;
        }
        return cursor - offset;
    }

    private decodeBinaryValue(bytes: Uint8Array, offset: number, layout: BinaryValueLayout, currentValue: any): { value: any, length: number } {
        const typeName = layout.typeName.toLowerCase();
        const requireBytes = (length: number): void => {
            if (offset + length > bytes.length) {
                throw new RangeError('Binary Get reached end of file');
            }
        };
        switch (typeName) {
            case 'array': {
                if (!layout.elementType) {
                    this.throwVbaError(VbaErrorCode.TYPE_MISMATCH, 'Missing Binary Get array element type');
                }
                const target = Array.isArray(currentValue) ? currentValue : [];
                const dims = (target as any).__vbaDimensions__ as { lower: number, upper: number }[] | undefined;
                let length = 0;
                const dimensions = dims?.length ?? 1;
                const coordinates: number[] = [];
                const writeAt = (array: any, indexes: number[], nextValue: any): void => {
                    let current = array;
                    for (let i = 0; i < indexes.length - 1; i++) current = current[indexes[i]];
                    current[indexes[indexes.length - 1]] = nextValue;
                };
                const restore = (dimension: number): void => {
                    const lower = dims?.[dimension].lower ?? (dimension === 0 ? (target as any).vbaBase ?? 0 : 0);
                    const upper = dims?.[dimension].upper ?? (dimension === 0 ? target.length - 1 : target[0].length - 1);
                    for (let i = lower; i <= upper; i++) {
                        coordinates[dimension] = i;
                        if (dimension === 0) {
                            const current = readNestedArrayValue(target, coordinates);
                            const field = this.decodeBinaryValue(bytes, offset + length,
                                { typeName: layout.elementType!, fixedLength: layout.fixedLength }, current);
                            writeAt(target, coordinates, field.value);
                            length += field.length;
                        } else restore(dimension - 1);
                    }
                };
                const readNestedArrayValue = (array: any, indexes: number[]): any => {
                    let current = array;
                    for (const index of indexes) current = current[index];
                    return current;
                };
                restore(dimensions - 1);
                return { value: target, length };
            }
            case 'byte':
                requireBytes(1);
                return { value: bytes[offset], length: 1 };
            case 'integer':
                requireBytes(2);
                return { value: new DataView(bytes.buffer, bytes.byteOffset + offset, 2).getInt16(0, true), length: 2 };
            case 'boolean':
                requireBytes(2);
                return {
                    value: new DataView(bytes.buffer, bytes.byteOffset + offset, 2).getInt16(0, true) !== 0 ? vbaTrue : vbaFalse,
                    length: 2,
                };
            case 'long':
                requireBytes(4);
                return { value: new DataView(bytes.buffer, bytes.byteOffset + offset, 4).getInt32(0, true), length: 4 };
            case 'longlong':
            case 'longptr':
                requireBytes(8);
                return { value: new DataView(bytes.buffer, bytes.byteOffset + offset, 8).getBigInt64(0, true), length: 8 };
            case 'single':
                requireBytes(4);
                return { value: new DataView(bytes.buffer, bytes.byteOffset + offset, 4).getFloat32(0, true), length: 4 };
            case 'double':
                requireBytes(8);
                return { value: new DataView(bytes.buffer, bytes.byteOffset + offset, 8).getFloat64(0, true), length: 8 };
            case 'date':
                requireBytes(8);
                return { value: new VbaDate(new DataView(bytes.buffer, bytes.byteOffset + offset, 8).getFloat64(0, true)), length: 8 };
            case 'currency':
                requireBytes(8);
                return { value: new VbaCurrency(new DataView(bytes.buffer, bytes.byteOffset + offset, 8).getBigInt64(0, true)), length: 8 };
            case 'string': {
                if (layout.variableLengthDescriptor) {
                    if (offset + 2 > bytes.length) {
                        this.throwVbaError(VbaErrorCode.INPUT_PAST_END_OF_FILE, 'Input past end of file');
                    }
                    const byteLength = new DataView(bytes.buffer, bytes.byteOffset + offset, 2).getUint16(0, true);
                    if (offset + 2 + byteLength > bytes.length) {
                        this.throwVbaError(VbaErrorCode.INPUT_PAST_END_OF_FILE, 'Input past end of file');
                    }
                    return {
                        value: iconv.decode(Buffer.from(bytes.subarray(offset + 2, offset + 2 + byteLength)), Evaluator.VBA_BINARY_ENCODING),
                        length: 2 + byteLength,
                    };
                }
                const length = layout.readToEnd
                    ? bytes.length - offset
                    : this.encodedByteLengthForCharacters(bytes, offset,
                        layout.fixedLength ?? String(currentValue ?? '').length);
                if (!layout.readToEnd && length < (layout.fixedLength ?? String(currentValue ?? '').length)) {
                    throw new RangeError('Binary Get reached end of file');
                }
                return {
                    value: iconv.decode(Buffer.from(bytes.subarray(offset, offset + length)), Evaluator.VBA_BINARY_ENCODING),
                    length,
                };
            }
            default: {
                const members = this.env.getType(layout.typeName);
                if (!members) this.throwVbaError(VbaErrorCode.TYPE_MISMATCH, `Unsupported binary value type: ${layout.typeName}`);
                const value = this.instantiateType(layout.typeName);
                let length = 0;
                for (const member of members) {
                    if (member.isArray) {
                        const field = this.decodeBinaryUdtArray(bytes, offset + length, member,
                            value[member.name.toLowerCase()]);
                        value[member.name.toLowerCase()] = field.value;
                        length += field.length;
                        continue;
                    }
                    const field = this.decodeBinaryValue(bytes, offset + length, {
                        typeName: member.memberType,
                        fixedLength: member.fixedLength,
                        variableLengthDescriptor: member.memberType.toLowerCase() === 'string' && member.fixedLength === undefined,
                    }, currentValue?.[member.name.toLowerCase()]);
                    value[member.name.toLowerCase()] = field.value;
                    length += field.length;
                }
                return { value, length };
            }
        }
    }

    /** Restore a fixed-size UDT member array while retaining its VBA bounds metadata. */
    private decodeBinaryUdtArray(bytes: Uint8Array, offset: number, member: TypeMember, value: any): { value: any, length: number } {
        const bounds = member.arrayBounds;
        if (!bounds || bounds.length === 0) {
            this.throwVbaError(VbaErrorCode.TYPE_MISMATCH, 'Binary Put/Get requires fixed-size UDT array members');
        }

        const decodeDimension = (array: any, dimension: number, position: number): number => {
            if (!Array.isArray(array)) {
                this.throwVbaError(VbaErrorCode.TYPE_MISMATCH, 'Invalid UDT array member for Binary Get');
            }
            const { lower, upper } = bounds![dimension];
            let length = 0;
            for (let index = lower; index <= upper; index++) {
                if (dimension < bounds!.length - 1) {
                    length += decodeDimension(array[index], dimension + 1, position + length);
                } else {
                    const field = this.decodeBinaryValue(bytes, position + length, {
                        typeName: member.memberType,
                        fixedLength: member.fixedLength,
                        variableLengthDescriptor: member.memberType.toLowerCase() === 'string' && member.fixedLength === undefined,
                    }, array[index]);
                    array[index] = field.value;
                    length += field.length;
                }
            }
            return length;
        };

        return { value, length: decodeDimension(value, 0, offset) };
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
                if (this.fs.existsSync(realPath) && this.fs.statSync(realPath).isDirectory()) {
                    this.throwVbaError(VbaErrorCode.PATH_FILE_ACCESS_ERROR, 'Path/File access error');
                }
                this.fs.unlinkSync(realPath);
            } catch (e) {
                if (e && typeof e === 'object' && 'type' in e && (e as { type?: string }).type === 'VbaError') {
                    throw e;
                }
                this.throwVbaError(VbaErrorCode.FILE_NOT_FOUND, 'File not found');
            }
        }
    }

    private evaluateKillStatement(stmt: KillStatement) {
        const vbaPath = this.toVbaString(this.evaluateExpression(stmt.path) ?? '');
        this.executeKill(vbaPath);
    }

    private evaluateWriteStatement(stmt: WriteStatement) {
        const fileNum = this.evaluateFileNumber(stmt.fileNumber);
        const handle = this.fileHandles.get(fileNum);
        if (!handle) this.throwVbaError(VbaErrorCode.BAD_FILE_NAME_OR_NUMBER, `Bad file name or number: #${fileNum}`);
        this.requireFileMode(handle, ['Output', 'Append']);

        const output = stmt.items.map(item => {
            const val = this.evaluateExpression(item);
            if (typeof val === 'string') return `"${val}"`;
            if (val instanceof VbaDate) {
                // Write # の日付は universal date format（#yyyy-mm-dd hh:nn:ss#）
                const d = fromVbaDate(val.value);
                const pad = (n: number) => String(n).padStart(2, '0');
                const datePart = `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
                const hasFraction = Math.abs(val.value - Math.round(val.value)) > 1e-10;
                const timePart = `${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`;
                if (!hasFraction) return `#${datePart}#`;
                return Math.round(val.value) === 0 ? `#${timePart}#` : `#${datePart} ${timePart}#`;
            }
            if (val instanceof VbaBoolean) return val.value !== 0 ? '#TRUE#' : '#FALSE#';
            if (val === vbaEmpty || val === undefined) return '""';
            if (val === vbaNull) return '#NULL#';
            return String(val);
        }).join(",");

        const lineOutput = this.encodeVbaFileText(output + "\r\n");  // Print # と同じく CP932/CRLF
        const writePosition = handle.mode === 'Append' ? null : (handle.pos ?? null);
        this.fs.writeSync(handle.fd, lineOutput, 0, lineOutput.length, writePosition);
        handle.pos = handle.mode === 'Append'
            ? this.fs.statSync(handle.path).size
            : handle.pos! + lineOutput.length;
    }

    private evaluateInputStatement(stmt: InputStatement) {
        const fileNum = this.evaluateFileNumber(stmt.fileNumber);
        const handle = this.fileHandles.get(fileNum);
        if (!handle) this.throwVbaError(VbaErrorCode.BAD_FILE_NAME_OR_NUMBER, `Bad file name or number: #${fileNum}`);
        this.requireFileMode(handle, ['Input', 'Binary']);

        // Input # consumes comma-separated Write # values across line boundaries.
        // Read the remaining bytes as one CP932 stream so multibyte characters and
        // records split across CRLF are not truncated at the first line.
        const contentBytes: number[] = [];
        const buf = new Uint8Array(1024);
        let bytesRead = 0;
        let readPos = handle.pos || 0;
        while ((bytesRead = this.fs.readSync(handle.fd, buf, 0, 1024, readPos)) > 0) {
            for (let i = 0; i < bytesRead; i++) contentBytes.push(buf[i]);
            readPos += bytesRead;
        }

        if (contentBytes.length === 0 && stmt.variables.length > 0) {
            handle.eof = true;
            this.throwVbaError(VbaErrorCode.INPUT_PAST_END_OF_FILE, 'Input past end of file');
        }

        const content = this.decodeVbaFileText(new Uint8Array(contentBytes));
        // 引用符内のカンマ・改行で分割しないフィールド分割（Bug 32-C）
        const splitInputFields = (src: string): Array<{ raw: string; end: number }> => {
            const fields: Array<{ raw: string; end: number }> = [];
            let cur = '';
            let inQuotes = false;
            let endedWithDelimiter = false;
            for (let i = 0; i < src.length; i++) {
                const ch = src[i];
                if (inQuotes) {
                    if (ch === '"') {
                        if (src[i + 1] === '"') { cur += '"'; i++; continue; }
                        inQuotes = false;
                    }
                    cur += ch;
                } else if (ch === '"') {
                    inQuotes = true;
                    cur += ch;
                } else if (ch === ',' || ch === '\r' || ch === '\n') {
                    let end = i + 1;
                    if (ch === '\r' && src[i + 1] === '\n') { end++; i++; }
                    fields.push({ raw: cur, end });
                    cur = '';
                    endedWithDelimiter = ch === ',';
                } else {
                    cur += ch;
                    endedWithDelimiter = false;
                }
            }
            if (cur.length > 0 || fields.length === 0 || endedWithDelimiter) {
                fields.push({ raw: cur, end: src.length });
            }
            return fields;
        };
        const fields = splitInputFields(content);
        const rawValues = fields.slice(0, stmt.variables.length).map(field => field.raw.trim());
        if (fields.length > 0 && stmt.variables.length > 0) {
            const consumedEnd = fields[Math.min(stmt.variables.length, fields.length) - 1].end;
            handle.pos = (handle.pos || 0) + this.encodeVbaFileText(content.slice(0, consumedEnd)).length;
        }
        const parseInputValue = (raw: string): any => {
            const t = raw.trim();
            if (t.length === 0) return vbaEmpty;
            if (t.startsWith('"') && t.endsWith('"')) return t.slice(1, -1);
            const lower = t.toLowerCase();
            if (lower === '#true#') return vbaTrue;
            if (lower === '#false#') return vbaFalse;
            if (lower === '#null#') return vbaNull;
            // Write # が出力した日付 #yyyy-mm-dd [hh:nn:ss]# を VbaDate に復元
            if (t.length > 2 && t.startsWith('#') && t.endsWith('#')) {
                try { return new VbaDate(toVbaDate(parseVbaDate(t.slice(1, -1)))); } catch { /* fall through */ }
            }
            const n = Number(t);
            return isNaN(n) ? t : n;
        };
        for (let i = 0; i < stmt.variables.length; i++) {
            if (i >= rawValues.length) {
                this.throwVbaError(VbaErrorCode.INPUT_PAST_END_OF_FILE, 'Input past end of file');
            }
            this.evaluateAssignmentToVariable(stmt.variables[i], parseInputValue(rawValues[i]));
        }
    }

    private evaluateGetStatement(stmt: GetStatement) {
        const fileNum = this.evaluateFileNumber(stmt.fileNumber);
        const handle = this.fileHandles.get(fileNum);
        if (!handle) this.throwVbaError(VbaErrorCode.BAD_FILE_NAME_OR_NUMBER, `Bad file name or number: #${fileNum}`);
        this.requireFileMode(handle, ['Random', 'Binary']);

        const position = this.resolveFileRecordPosition(handle, stmt.recordNumber);

        const currentValue = this.evaluateExpression(stmt.variable);
        const layout = this.getBinaryValueLayout(stmt.variable, currentValue);
        // Read the remaining file so a CP932 variable-length string can find
        // its character boundary. decodeBinaryValue consumes only its field.
        const remaining = Math.max(0, this.fs.statSync(handle.path).size - (position ?? 0));
        const buffer = new Uint8Array(remaining);
        const bytesRead = this.fs.readSync(handle.fd, buffer, 0, buffer.length, position);
        let decoded: { value: any, length: number };
        try {
            decoded = this.decodeBinaryValue(buffer.subarray(0, bytesRead), 0, layout, currentValue);
        } catch (error) {
            const errorNumber = error && typeof error === 'object' && 'number' in error
                ? (error as { number?: number }).number
                : undefined;
            if (error instanceof RangeError || errorNumber === VbaErrorCode.INPUT_PAST_END_OF_FILE) {
                handle.eof = true;
                this.throwVbaError(VbaErrorCode.INPUT_PAST_END_OF_FILE, 'Input past end of file');
            }
            throw error;
        }
        // Binary/Random Get writes into the existing destination array.
        this.evaluateAssignmentToVariable(stmt.variable, decoded.value, undefined, true);
        handle.pos = (position ?? handle.pos ?? 0) + decoded.length;
        if (handle.mode === 'Random') {
            handle.lastRecord = Math.floor((position ?? handle.pos ?? 0) / (handle.recordLen ?? 128)) + 1;
            handle.nextRecord = handle.lastRecord + 1;
            handle.eof = false;
        } else if (handle.mode === 'Binary') {
            handle.eof = false;
        }
    }

    private resolveFileRecordPosition(
        handle: { mode: OpenStatement['mode']; pos?: number; recordLen?: number; nextRecord?: number },
        recordNumber?: Expression,
    ): number | null {
        if (!recordNumber) {
            if (handle.mode === 'Random') {
                const nextRecord = handle.nextRecord ?? 1;
                return (nextRecord - 1) * (handle.recordLen ?? 128);
            }
            return handle.pos ?? null;
        }
        const value = this.evaluateExpression(recordNumber);
        if (value === vbaNull) {
            this.throwVbaError(VbaErrorCode.INVALID_USE_OF_NULL, 'Invalid use of Null');
        }
        const record = this.toVbaNumber(value);
        if (!Number.isInteger(record) || record < 1) {
            this.throwVbaError(VbaErrorCode.INVALID_PROCEDURE_CALL, 'Invalid procedure call or argument');
        }
        return handle.mode === 'Random'
            ? (record - 1) * (handle.recordLen ?? 128)
            : record - 1;
    }

    private evaluateSeekStatement(stmt: SeekStatement) {
        const fileNum = this.evaluateFileNumber(stmt.fileNumber);
        const handle = this.fileHandles.get(fileNum);
        if (!handle) this.throwVbaError(VbaErrorCode.BAD_FILE_NAME_OR_NUMBER, `Bad file name or number: #${fileNum}`);

        const positionValue = this.evaluateExpression(stmt.position);
        if (positionValue === vbaNull) {
            this.throwVbaError(VbaErrorCode.INVALID_USE_OF_NULL, 'Invalid use of Null');
        }
        const rawPos = this.toVbaNumber(positionValue);
        if (!Number.isFinite(rawPos)) {
            this.throwVbaError(VbaErrorCode.BAD_RECORD_NUMBER, 'Bad record number');
        }
        const pos = Math.round(rawPos);
        if (pos < 1 || (handle.mode === 'Random' && pos > 2147483647)) {
            this.throwVbaError(VbaErrorCode.BAD_RECORD_NUMBER, 'Bad record number');
        }
        // Node doesn't have seekSync on FD directly without lseek,
        // but we can track it in our handle if we use it for subsequent read/write.
        if (handle.mode === 'Random') {
            handle.nextRecord = Math.max(1, pos);
            handle.pos = (handle.nextRecord - 1) * (handle.recordLen ?? 128);
            handle.eof = false;
        } else {
            handle.pos = Math.max(0, pos - 1);
            if (handle.mode === 'Binary') handle.eof = false;
        }
    }

    private evaluateResetStatement(_stmt: ResetStatement) {
        for (const [_num, handle] of this.fileHandles) {
            this.releaseFileLocks(handle);
            this.fs.closeSync(handle.fd);
        }
        this.fileHandles.clear();
    }

    private releaseFileLocks(handle: { fileNumber: number, path: string, locks?: Array<{ start: number, end: number, key: string }> }): void {
        const pathLocks = this.fileLocks.get(handle.path);
        if (!pathLocks) return;
        const remaining = pathLocks.filter(lock => lock.owner !== handle.fileNumber);
        if (remaining.length > 0) this.fileLocks.set(handle.path, remaining);
        else this.fileLocks.delete(handle.path);
        handle.locks?.splice(0, handle.locks.length);
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
            const params: ArgBinderParam[] = stmt.parameters.map(param => ({
                isOptional: !!param.isOptional || param.defaultValue != null,
                isParamArray: !!param.isParamArray,
            }));
            this.checkArgCountGeneric(params, _args.length);
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
        const decoratedFactory = () => this.decorateComObject(factory());
        const register = (key: string) => {
            const k = key.toLowerCase();
            if (!this.externalObjectFactories.has(k)) {
                this.externalObjectFactories.set(k, decoratedFactory);
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
            const sample = decoratedFactory();
            const progId = sample.__progId__;
            // 主キー登録
            this.externalObjectFactories.set(progId.toLowerCase(), decoratedFactory);
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

    /**
     * Attach the same parameter metadata used by builtins to host/COM methods.
     * COM factories opt in with a non-VBA `__vbaParamSpecs__` map so dynamic
     * objects share argument binding and coercion without hard-coding dispatch
     * rules in the evaluator.
     */
    private decorateComObject(object: any): any {
        const specs = object?.__vbaParamSpecs__ as Record<string, BuiltinParamSpec[]> | undefined;
        if (!specs) return object;
        for (const [name, params] of Object.entries(specs)) {
            const method = object[name];
            if (typeof method === 'function' && !(method as any).__vbaParamSpec__) {
                (method as any).__vbaParamSpec__ = params;
            }
        }
        return object;
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
            let compareMode = 0;
            const matchingKey = (key: any): any => {
                if (key === vbaNull) {
                    this.throwVbaError(VbaErrorCode.INVALID_USE_OF_NULL, 'Invalid use of Null');
                }
                if (compareMode !== 1 || typeof key !== 'string') return key;
                return Array.from(dict.keys()).find(existing =>
                    typeof existing === 'string' && existing.toLowerCase() === key.toLowerCase()) ?? key;
            };
            const dictionary: any = {
                __isVbaDict__: true,
                __progId__: 'Scripting.Dictionary',
                __map__: dict,
                __resolveKey__: matchingKey,
                add: (k: any, v: any) => {
                    const key = matchingKey(k);
                    if (dict.has(key)) this.throwVbaError(VbaErrorCode.KEY_ALREADY_EXISTS, 'This key is already associated with an element of this collection');
                    dict.set(key, v);
                },
                exists: (k: any) => dict.has(matchingKey(k)) ? vbaTrue : vbaFalse,
                remove: (k: any) => dict.delete(matchingKey(k)),
                removeall: () => dict.clear(),
                count: () => dict.size,
                keys: () => Array.from(dict.keys()),
                items: () => Array.from(dict.values()),
                item: (k: any, v?: any) => {
                    const key = matchingKey(k);
                    if (v !== undefined) {
                        dict.set(key, v);
                    } else if (!dict.has(key)) {
                        // VBA auto-creates the key with Empty when reading a missing key
                        const loc = `${this.executingModuleName || this.currentSourceModule}:${this.currentLine}`;
                        console.warn(`[vba-runner] ${loc}: Dictionary.Item("${k}"): key not found, auto-creating with Empty (VBA compatible)`);
                        dict.set(key, vbaEmpty);
                    }
                    return dict.get(key);
                }
            };
            Object.defineProperty(dictionary, 'comparemode', {
                get: () => compareMode,
                set: (value: any) => { compareMode = Number(value); },
                enumerable: true,
            });
            return dictionary;
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
                __vbaParamSpecs__: {
                    add: [
                        { name: 'Item' },
                        { name: 'Key', optional: true, coerce: 'string' },
                        { name: 'Before', optional: true },
                        { name: 'After', optional: true },
                    ],
                },
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
        const throwFsoPathError = (error: any): never => {
            if (error?.type === 'VbaError') throw error;
            if (error?.code === 'ENOENT' || /ENOENT|not found/i.test(String(error?.message ?? error))) {
                this.throwVbaError(VbaErrorCode.PATH_NOT_FOUND, 'Path not found');
            }
            throw error;
        };
        const textStreamParamSpecs: Record<string, BuiltinParamSpec[]> = {
            write: [{ name: 'Text', coerce: 'string' }],
            writeline: [{ name: 'Text', coerce: 'string' }],
            writeblanklines: [{ name: 'NumberOfLines' }],
            read: [{ name: 'Characters' }],
            skip: [{ name: 'Characters' }],
        };
        // File/Folder は FSO の複数メソッドから返される同じ COM 境界の
        // オブジェクトである。ここで共通のパス属性と操作を組み立て、
        // GetFile/GetFolder ごとに個別の薄いオブジェクトを作らない。
        const pathObjectParamSpecs: Record<string, BuiltinParamSpec[]> = {
            delete: [{ name: 'Force', optional: true, coerce: 'boolean' }],
            copy: [
                { name: 'Destination', coerce: 'string' },
                { name: 'Overwrite', optional: true, coerce: 'boolean' },
            ],
            move: [{ name: 'Destination', coerce: 'string' }],
            openastextstream: [
                { name: 'IOMode', optional: true },
                { name: 'Format', optional: true },
            ],
        };
        const makePathObject = (kind: 'file' | 'folder', requestedPath: string, stats: any): any => {
            const full = this.sandbox.toRealPath(requestedPath);
            const isFile = kind === 'file';
            const base = path.win32.basename(requestedPath);
            const parentPath = path.win32.dirname(requestedPath);
            const common: any = {
                path: requestedPath,
                name: base,
                parentfolder: this.decorateComObject({
                    path: parentPath,
                    name: path.win32.basename(parentPath),
                }),
                datecreated: new VbaDate(toVbaDate(stats.birthtime || stats.mtime)),
                datelastaccessed: new VbaDate(toVbaDate(stats.atime || stats.mtime)),
                datelastmodified: new VbaDate(toVbaDate(stats.mtime)),
                attributes: stats.mode || 0,
                __vbaParamSpecs__: pathObjectParamSpecs,
            };
            if (isFile) {
                common.size = stats.size;
                common.type = path.win32.extname(base).replace(/^\./, '');
                common.delete = (force: any = false) => {
                    try { this.fs.unlinkSync(full); }
                    catch (e: any) {
                        if (vbaFlagIsTrue(force) && e?.code === 'EACCES') {
                            (this.fs as any).chmodSync?.(full, 0o666);
                            this.fs.unlinkSync(full);
                            return;
                        }
                        throwFsoPathError(e);
                    }
                };
                common.copy = (destination: string, overwrite: any = false) => {
                    const target = this.sandbox.toRealPath(destination);
                    if (this.fs.existsSync(target) && !vbaFlagIsTrue(overwrite)) {
                        this.throwVbaError(VbaErrorCode.FILE_ALREADY_EXISTS, 'File already exists');
                    }
                    if (!this.fs.copyFileSync) this.throwVbaError(VbaErrorCode.INVALID_PROCEDURE_CALL, 'File copy is not supported');
                    try { this.fs.copyFileSync(full, target); } catch (e) { throwFsoPathError(e); }
                };
                common.move = (destination: string) => {
                    const target = this.sandbox.toRealPath(destination);
                    if (this.fs.existsSync(target)) {
                        this.throwVbaError(VbaErrorCode.FILE_ALREADY_EXISTS, 'File already exists');
                    }
                    if (!this.fs.copyFileSync) this.throwVbaError(VbaErrorCode.INVALID_PROCEDURE_CALL, 'File move is not supported');
                    try {
                        this.fs.copyFileSync(full, target);
                        this.fs.unlinkSync(full);
                    } catch (e) { throwFsoPathError(e); }
                };
            } else {
                common.delete = (force: any = false) => {
                    if (!this.fs.existsSync(full)) this.throwVbaError(VbaErrorCode.PATH_NOT_FOUND, 'Path not found');
                    this.fs.rmSync?.(full, { recursive: vbaFlagIsTrue(force), force: vbaFlagIsTrue(force) });
                };
                common.copy = (destination: string, overwrite: any = false) => {
                    const target = this.sandbox.toRealPath(destination);
                    if (this.fs.existsSync(target) && !vbaFlagIsTrue(overwrite)) {
                        this.throwVbaError(VbaErrorCode.FILE_ALREADY_EXISTS, 'File already exists');
                    }
                    (this.fs as any).cpSync?.(full, target, { recursive: true, force: vbaFlagIsTrue(overwrite) });
                };
                common.move = (destination: string) => {
                    const target = this.sandbox.toRealPath(destination);
                    if (this.fs.existsSync(target)) this.throwVbaError(VbaErrorCode.FILE_ALREADY_EXISTS, 'File already exists');
                    (this.fs as any).renameSync?.(full, target);
                };
            }
            return this.decorateComObject(common);
        };
        let fsoObject: any;
        this.registerComObject( () => (fsoObject = ({
            __isVbaFso__: true,
            __progId__: 'Scripting.FileSystemObject',
            __vbaParamSpecs__: {
                createtextfile: [
                    { name: 'FileName' },
                    { name: 'Overwrite', optional: true, coerce: 'boolean' },
                    { name: 'Unicode', optional: true, coerce: 'boolean' },
                ],
                opentextfile: [
                    { name: 'FileName' },
                    { name: 'IOMode', optional: true },
                    { name: 'Create', optional: true, coerce: 'boolean' },
                    { name: 'Format', optional: true },
                ],
                copyfile: [
                    { name: 'Source' },
                    { name: 'Destination' },
                    { name: 'Overwrite', optional: true, coerce: 'boolean' },
                ],
                copyfolder: [
                    { name: 'Source' },
                    { name: 'Destination' },
                    { name: 'Overwrite', optional: true, coerce: 'boolean' },
                ],
                movefolder: [
                    { name: 'Source' },
                    { name: 'Destination' },
                ],
                fileexists: [{ name: 'filespec', coerce: 'string' }],
                folderexists: [{ name: 'folderspec', coerce: 'string' }],
                deletefile: [
                    { name: 'filespec', coerce: 'string' },
                    { name: 'force', optional: true, coerce: 'boolean' },
                ],
                movefile: [
                    { name: 'source', coerce: 'string' },
                    { name: 'destination', coerce: 'string' },
                ],
                createfolder: [{ name: 'path', coerce: 'string' }],
                deletefolder: [
                    { name: 'path' },
                    { name: 'force', optional: true, coerce: 'boolean' },
                ],
                getfile: [{ name: 'filespec', coerce: 'string' }],
                getfolder: [{ name: 'folderspec', coerce: 'string' }],
                getbasename: [{ name: 'path', coerce: 'string' }],
                getextensionname: [{ name: 'path', coerce: 'string' }],
                getparentfoldername: [{ name: 'path', coerce: 'string' }],
                getabsolutepathname: [{ name: 'path', coerce: 'string' }],
            },
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
            createtextfile: (p: string, overwrite: any = true, unicode: any = false) => {
                const full = this.sandbox.toRealPath(p);
                if (!vbaFlagIsTrue(overwrite) && this.fs.existsSync(full)) {
                    this.throwVbaError(VbaErrorCode.FILE_ALREADY_EXISTS, "File already exists");
                }
                let fd = -1;
                try { fd = this.fs.openSync(full, 'w'); } catch (e) { throwFsoPathError(e); }
                const useUnicode = vbaFlagIsTrue(unicode);
                if (useUnicode) this.fs.writeSync(fd, encodeTextStream('', true));
                let closed = false;
                const ensureOpen = () => {
                    if (closed) this.throwVbaError(VbaErrorCode.INVALID_PROCEDURE_CALL,
                        'Operation is not allowed when object is closed');
                };
                const writeText = (s: string) => {
                    ensureOpen();
                    const bytes = encodeTextStream(s, useUnicode);
                    // encodeTextStream includes a BOM; only the first write owns it.
                    const payload = useUnicode ? bytes.subarray(2) : bytes;
                    this.fs.writeSync(fd, payload);
                };
                return this.decorateComObject({
                    __vbaParamSpecs__: textStreamParamSpecs,
                    write: writeText,
                    writeline: (s: string) => writeText(s + "\r\n"),
                    writeblanklines: (count: number) => writeText("\r\n".repeat(Math.max(0, Math.trunc(Number(count))))),
                    close: () => {
                        if (closed) return;
                        this.fs.closeSync(fd);
                        closed = true;
                    }
                });
            },
            opentextfile: (p: string, iomode: any = 1, create: any = false, format: any = -2) => {
                const full = this.sandbox.toRealPath(p);
                if (iomode === vbaNull || format === vbaNull) {
                    this.throwVbaError(VbaErrorCode.INVALID_USE_OF_NULL, 'Invalid use of Null');
                }
                const modeValue = iomode === vbaEmpty ? 1 : Number(iomode?.valueOf?.() ?? iomode);
                if (![1, 2, 8].includes(modeValue)) {
                    this.throwVbaError(VbaErrorCode.INVALID_PROCEDURE_CALL, 'Invalid procedure call or argument');
                }
                const formatValue = format === vbaEmpty ? -2 : Number(format?.valueOf?.() ?? format);
                if (![ -2, -1, 0 ].includes(formatValue)) {
                    this.throwVbaError(VbaErrorCode.INVALID_PROCEDURE_CALL, 'Invalid procedure call or argument');
                }
                const mode = modeValue === 1 ? 'r' : (modeValue === 2 ? 'w' : 'a');
                const shouldCreate = vbaFlagIsTrue(create);
                if (shouldCreate && !this.fs.existsSync(full)) this.fs.openSync(full, 'w');
                let fd: number;
                try {
                    fd = this.fs.openSync(full, mode);
                } catch (e: any) {
                    if (e?.code === 'ENOENT' || /ENOENT|not found/i.test(String(e?.message ?? e))) {
                        this.throwVbaError(VbaErrorCode.FILE_NOT_FOUND, 'File not found');
                    }
                    throw e;
                }
                const evalFs = this.fs;
                let closed = false;
                const ensureOpen = () => {
                    if (closed) this.throwVbaError(VbaErrorCode.INVALID_PROCEDURE_CALL,
                        'Operation is not allowed when object is closed');
                };
                let pos = 0;
                const useUnicode = formatValue === -1;
                const readContent = () => {
                    ensureOpen();
                    const size = evalFs.statSync(full).size;
                    const rawFd = evalFs.openSync(full, 'r');
                    const bytes = new Uint8Array(size);
                    let read = 0;
                    try {
                        while (read < size) {
                            const n = evalFs.readSync(rawFd, bytes, read, size - read, null);
                            if (n <= 0) break;
                            read += n;
                        }
                    } finally {
                        evalFs.closeSync(rawFd);
                    }
                    return decodeTextStream(bytes.subarray(0, read), useUnicode);
                };
                const nextLineBreak = (content: string, from: number) => {
                    for (let index = Math.max(0, from); index < content.length; index++) {
                        if (content[index] === '\r') {
                            return { index, length: content[index + 1] === '\n' ? 2 : 1 };
                        }
                        if (content[index] === '\n') return { index, length: 1 };
                    }
                    return undefined;
                };
                return this.decorateComObject({
                    __vbaParamSpecs__: textStreamParamSpecs,
                    get atendofstream() {
                        const content = readContent();
                        return pos >= content.length ? vbaTrue : vbaFalse;
                    },
                    get atendofline() {
                        const content = readContent();
                        return pos >= content.length || content[pos] === '\r' || content[pos] === '\n' ? vbaTrue : vbaFalse;
                    },
                    get line() {
                        const content = readContent();
                        const bounded = Math.min(pos, content.length);
                        return (content.slice(0, bounded).match(/\r\n|\r|\n/g)?.length ?? 0) + 1;
                    },
                    get column() {
                        const content = readContent();
                        const bounded = Math.min(pos, content.length);
                        const lastBreak = Math.max(content.lastIndexOf('\n', bounded - 1), content.lastIndexOf('\r', bounded - 1));
                        return bounded - lastBreak;
                    },
                    readall: () => {
                        const content = readContent();
                        const result = content.slice(pos);
                        pos = content.length;
                        return result;
                    },
                    read: (numChars: number) => {
                        const content = readContent();
                        const count = Math.max(0, Math.trunc(Number(numChars)));
                        const result = content.slice(pos, pos + count);
                        pos += result.length;
                        return result;
                    },
                    skip: (numChars: number) => {
                        const content = readContent();
                        const count = Math.max(0, Math.trunc(Number(numChars)));
                        pos = Math.min(content.length, pos + count);
                    },
                    readline: () => {
                        const content = readContent();
                        const breakAt = nextLineBreak(content, pos);
                        if (!breakAt) {
                            const line = content.slice(pos);
                            pos = content.length;
                            return line;
                        }
                        const line = content.slice(pos, breakAt.index);
                        pos = breakAt.index + breakAt.length;
                        return line;
                    },
                    skipline: () => {
                        const content = readContent();
                        const breakAt = nextLineBreak(content, pos);
                        pos = breakAt ? breakAt.index + breakAt.length : content.length;
                    },
                    write: (s: string) => {
                        ensureOpen();
                        const bytes = encodeTextStream(s, useUnicode);
                        const payload = useUnicode && (evalFs.statSync(full).size === 0)
                            ? bytes
                            : (useUnicode ? bytes.subarray(2) : bytes);
                        this.fs.writeSync(fd, payload);
                    },
                    writeline: (s: string) => {
                        ensureOpen();
                        const text = s + "\r\n";
                        const bytes = encodeTextStream(text, useUnicode);
                        const payload = useUnicode && (evalFs.statSync(full).size === 0)
                            ? bytes
                            : (useUnicode ? bytes.subarray(2) : bytes);
                        this.fs.writeSync(fd, payload);
                    },
                    close: () => {
                        if (closed) return;
                        this.fs.closeSync(fd);
                        closed = true;
                    }
                });
            },
            createfolder: (p: string) => {
                const full = this.sandbox.toRealPath(p);
                this.validateSingleDirectoryCreation(full, VbaErrorCode.FILE_ALREADY_EXISTS,
                    'File already exists', false);
                this.fs.mkdirSync(full);
                return { path: p };
            },
            deletefile: (p: string, force: any = false) => {
                const full = this.sandbox.toRealPath(p);
                try {
                    if (this.fs.statSync(full).isDirectory()) {
                        this.throwVbaError(VbaErrorCode.PATH_FILE_ACCESS_ERROR, 'Path/File access error');
                    }
                    const attributes = this.fs.getAttributes?.(full);
                    const readOnly = attributes !== undefined
                        ? (attributes & VBA_FILE_ATTRIBUTE.READ_ONLY) !== 0
                        : (() => {
                            const stat = this.fs.statSync(full);
                            return stat.mode !== undefined && (stat.mode & 0o222) === 0;
                        })();
                    if (readOnly) {
                        if (!vbaFlagIsTrue(force)) {
                            this.throwVbaError(VbaErrorCode.PERMISSION_DENIED, 'Permission denied');
                        }
                        if (this.fs.setAttributes && attributes !== undefined) {
                            this.fs.setAttributes(full, attributes & ~VBA_FILE_ATTRIBUTE.READ_ONLY);
                        } else if ((this.fs as any).chmodSync) {
                            (this.fs as any).chmodSync(full, 0o666);
                        }
                    }
                    this.fs.unlinkSync(full);
                } catch (e: any) {
                    if (e?.code === 'ENOENT' || /ENOENT|not found/i.test(String(e?.message ?? e)))
                        this.throwVbaError(VbaErrorCode.FILE_NOT_FOUND, 'File not found');
                    throw e;
                }
            },
            copyfile: (source: string, destination: string, overwrite: boolean = false) => {
                const sourcePath = this.sandbox.toRealPath(source);
                const destinationPath = this.sandbox.toRealPath(destination);
                try {
                    if (!this.fs.existsSync(sourcePath) || !this.fs.statSync(sourcePath).isFile()) {
                        this.throwVbaError(VbaErrorCode.FILE_NOT_FOUND, 'File not found');
                    }
                    if (this.fs.existsSync(destinationPath) && !vbaFlagIsTrue(overwrite)) {
                        this.throwVbaError(VbaErrorCode.FILE_ALREADY_EXISTS, 'File already exists');
                    }
                    if (!this.fs.copyFileSync) {
                        this.throwVbaError(VbaErrorCode.INVALID_PROCEDURE_CALL, 'File copy is not supported');
                    }
                    this.fs.copyFileSync(sourcePath, destinationPath);
                } catch (e: any) {
                    if (e?.type === 'VbaError') throw e;
                    if (e?.code === 'ENOENT' || /ENOENT|not found/i.test(String(e?.message ?? e))) {
                        this.throwVbaError(VbaErrorCode.FILE_NOT_FOUND, 'File not found');
                    }
                    throw e;
                }
            },
            copyfolder: (source: string, destination: string, overwrite: any = false) => {
                const sourcePath = this.sandbox.toRealPath(source);
                const destinationPath = this.sandbox.toRealPath(destination);
                if (overwrite === vbaNull) {
                    this.throwVbaError(VbaErrorCode.INVALID_USE_OF_NULL, 'Invalid use of Null');
                }
                try {
                    if (!this.fs.existsSync(sourcePath) || !this.fs.statSync(sourcePath).isDirectory()) {
                        this.throwVbaError(VbaErrorCode.PATH_NOT_FOUND, 'Path not found');
                    }
                    if (this.fs.existsSync(destinationPath)) {
                        this.throwVbaError(VbaErrorCode.FILE_ALREADY_EXISTS, 'File already exists');
                    }
                    const parent = path.dirname(destinationPath);
                    if (!this.fs.existsSync(parent)) {
                        this.throwVbaError(VbaErrorCode.PATH_NOT_FOUND, 'Path not found');
                    }
                    const copyTree = (from: string, to: string): void => {
                        this.fs.mkdirSync(to);
                        for (const name of this.fs.readdirSync(from)) {
                            const fromChild = path.join(from, name);
                            const toChild = path.join(to, name);
                            if (this.fs.statSync(fromChild).isDirectory()) copyTree(fromChild, toChild);
                            else this.fs.copyFileSync?.(fromChild, toChild);
                        }
                    };
                    copyTree(sourcePath, destinationPath);
                } catch (e: any) {
                    if (e?.type === 'VbaError') throw e;
                    if (e?.code === 'ENOENT' || /ENOENT|not found/i.test(String(e?.message ?? e))) {
                        this.throwVbaError(VbaErrorCode.PATH_NOT_FOUND, 'Path not found');
                    }
                    throw e;
                }
            },
            movefolder: (source: string, destination: string) => {
                const sourcePath = this.sandbox.toRealPath(source);
                const destinationPath = this.sandbox.toRealPath(destination);
                try {
                    if (!this.fs.existsSync(sourcePath) || !this.fs.statSync(sourcePath).isDirectory()) {
                        this.throwVbaError(VbaErrorCode.PATH_NOT_FOUND, 'Path not found');
                    }
                    if (this.fs.existsSync(destinationPath)) {
                        this.throwVbaError(VbaErrorCode.FILE_ALREADY_EXISTS, 'File already exists');
                    }
                    if (!this.fs.existsSync(path.dirname(destinationPath))) {
                        this.throwVbaError(VbaErrorCode.PATH_NOT_FOUND, 'Path not found');
                    }
                    const moveTree = (from: string, to: string): void => {
                        this.fs.mkdirSync(to);
                        for (const name of this.fs.readdirSync(from)) {
                            const fromChild = path.join(from, name);
                            const toChild = path.join(to, name);
                            if (this.fs.statSync(fromChild).isDirectory()) moveTree(fromChild, toChild);
                            else this.fs.copyFileSync?.(fromChild, toChild);
                        }
                    };
                    moveTree(sourcePath, destinationPath);
                    this.fs.rmSync?.(sourcePath, { recursive: true, force: true });
                } catch (e: any) {
                    if (e?.type === 'VbaError') throw e;
                    if (e?.code === 'ENOENT' || /ENOENT|not found/i.test(String(e?.message ?? e))) {
                        this.throwVbaError(VbaErrorCode.PATH_NOT_FOUND, 'Path not found');
                    }
                    throw e;
                }
            },
            movefile: (source: string, destination: string) => {
                const sourcePath = this.sandbox.toRealPath(source);
                const destinationPath = this.sandbox.toRealPath(destination);
                try {
                    if (!this.fs.existsSync(sourcePath) || !this.fs.statSync(sourcePath).isFile()) {
                        this.throwVbaError(VbaErrorCode.FILE_NOT_FOUND, 'File not found');
                    }
                    if (this.fs.existsSync(destinationPath)) {
                        this.throwVbaError(VbaErrorCode.FILE_ALREADY_EXISTS, 'File already exists');
                    }
                    if (!this.fs.copyFileSync) {
                        this.throwVbaError(VbaErrorCode.INVALID_PROCEDURE_CALL, 'File move is not supported');
                    }
                    this.fs.copyFileSync(sourcePath, destinationPath);
                    this.fs.unlinkSync(sourcePath);
                } catch (e: any) {
                    if (e?.type === 'VbaError') throw e;
                    if (e?.code === 'ENOENT' || /ENOENT|not found/i.test(String(e?.message ?? e))) {
                        this.throwVbaError(VbaErrorCode.FILE_NOT_FOUND, 'File not found');
                    }
                    throw e;
                }
            },
            deletefolder: (p: string, force: any = false) => {
                if (force === vbaNull) {
                    this.throwVbaError(VbaErrorCode.INVALID_USE_OF_NULL, 'Invalid use of Null');
                }
                const full = this.sandbox.toRealPath(p);
                if (!this.fs.existsSync(full)) {
                    this.throwVbaError(VbaErrorCode.PATH_NOT_FOUND, 'Path not found');
                }
                if (!this.fs.statSync(full).isDirectory()) {
                    this.throwVbaError(VbaErrorCode.PATH_FILE_ACCESS_ERROR, 'Path/File access error');
                }
                const forceDelete = vbaFlagIsTrue(force);
                if (!forceDelete && this.fs.readdirSync(full).length > 0) {
                    this.throwVbaError(VbaErrorCode.PATH_FILE_ACCESS_ERROR, 'Path/File access error');
                }
                this.fs.rmSync?.(full, { recursive: forceDelete, force: forceDelete });
            },
            getfile: (p: string) => {
                const full = this.sandbox.toRealPath(p);
                let stats;
                try { stats = this.fs.statSync(full); } catch (e) {
                    if (e instanceof Error && /ENOENT|not found/i.test(e.message)) {
                        this.throwVbaError(VbaErrorCode.FILE_NOT_FOUND, 'File not found');
                    }
                    throw e;
                }
                if (!stats.isFile()) this.throwVbaError(VbaErrorCode.FILE_NOT_FOUND, 'File not found');
                const file = makePathObject('file', p, stats);
                file.openastextstream = (iomode: any = 1, format: any = -2) =>
                    fsoObject.opentextfile(p, iomode, false, format);
                return this.decorateComObject(file);
            },
            getfolder: (p: string) => {
                const full = this.sandbox.toRealPath(p);
                if (!this.fs.existsSync(full) || !this.fs.statSync(full).isDirectory()) {
                    this.throwVbaError(VbaErrorCode.PATH_NOT_FOUND, 'Path not found');
                }
                const folder = makePathObject('folder', p, this.fs.statSync(full));
                const files = new VbaCollection();
                const subfolders = new VbaCollection();
                for (const name of this.fs.readdirSync(full)) {
                    const childPath = path.win32.join(p, name);
                    const childFull = this.sandbox.toRealPath(childPath);
                    let childStats: any;
                    try { childStats = this.fs.statSync(childFull); } catch { continue; }
                    if (childStats.isFile()) files.add(makePathObject('file', childPath, childStats));
                    else if (childStats.isDirectory()) subfolders.add(makePathObject('folder', childPath, childStats));
                }
                Object.defineProperty(folder, 'files', { get: () => files, enumerable: true });
                Object.defineProperty(folder, 'subfolders', { get: () => subfolders, enumerable: true });
                Object.defineProperty(folder, 'isrootfolder', { get: () => path.win32.dirname(p) === p, enumerable: true });
                return this.decorateComObject(folder);
            },
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
        })));

        // --- MSXML2.XMLHTTP / Microsoft.XMLHTTP ---
        const xmlhttpFactory = () => {
            let responseText = "";
            let status = 0;
            return {
                __progId__: 'MSXML2.XMLHTTP',
                __vbaParamSpecs__: {
                    open: [
                        { name: 'bstrMethod', coerce: 'string' },
                        { name: 'bstrUrl', coerce: 'string' },
                        { name: 'varAsync', optional: true },
                        { name: 'bstrUser', optional: true, coerce: 'string' },
                        { name: 'bstrPassword', optional: true, coerce: 'string' },
                    ],
                    setrequestheader: [
                        { name: 'bstrHeader', coerce: 'string' },
                        { name: 'bstrValue', coerce: 'string' },
                    ],
                    send: [{ name: 'varBody', optional: true }],
                },
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
            let closed = true;
            const evaluator = this;
            const rejectNullArgument = (value: any) => {
                if (value === vbaNull) evaluator.throwVbaError(VbaErrorCode.INVALID_USE_OF_NULL, 'Invalid use of Null');
                return value;
            };
            const ensureOpen = () => {
                if (closed) evaluator.throwVbaError(VbaErrorCode.INVALID_PROCEDURE_CALL, 'Operation is not allowed when object is closed');
            };
            return {
                __progId__: 'ADODB.Stream',
                __vbaParamSpecs__: {
                    open: [{ name: 'Options', optional: true }],
                    write: [{ name: 'Buffer' }],
                    writetext: [{ name: 'Text', coerce: 'string' }],
                    read: [{ name: 'NumBytes', optional: true }],
                    readtext: [{ name: 'NumChars', optional: true }],
                    savetofile: [
                        { name: 'FileName', coerce: 'string' },
                        { name: 'SaveOptions', optional: true },
                    ],
                    loadfromfile: [{ name: 'FileName', coerce: 'string' }],
                },
                open: (options?: any) => { rejectNullArgument(options); closed = false; streamPos = 0; },
                close: () => { closed = true; },
                write: (data: any) => {
                    ensureOpen();
                    if (data === vbaNull) this.throwVbaError(VbaErrorCode.INVALID_USE_OF_NULL, 'Invalid use of Null');
                    content += String(data);
                    streamPos = content.length;
                },
                writetext: (text: string) => { ensureOpen(); content += text; streamPos = content.length; },
                read: (len?: number) => {
                    ensureOpen();
                    rejectNullArgument(len);
                    const requested = len === undefined || Number(len) < 0 ? content.length - streamPos : Math.max(0, Math.trunc(Number(len)));
                    const r = content.slice(streamPos, streamPos + requested);
                    streamPos += r.length;
                    return r;
                },
                readtext: (...args: any[]) => {
                    ensureOpen();
                    const numChars = args[0] as number | undefined;
                    rejectNullArgument(numChars);
                    const requested = numChars === undefined || Number(numChars) < 0
                        ? content.length - streamPos
                        : Math.max(0, Math.trunc(Number(numChars)));
                    const r = content.slice(streamPos, streamPos + requested);
                    streamPos += r.length;
                    return r;
                },
                savetofile: (p: string, _mode: number = 1) => {
                    ensureOpen();
                    rejectNullArgument(_mode);
                    const full = this.sandbox.toRealPath(p);
                    this.fs.writeFileSync(full, content);
                },
                loadfromfile: (p: string) => {
                    ensureOpen();
                    const full = this.sandbox.toRealPath(p);
                    try {
                        content = this.fs.readFileSync(full, 'utf8');
                    } catch (e: any) {
                        if (e?.code === 'ENOENT' || /ENOENT|not found/i.test(String(e?.message ?? e))) {
                            this.throwVbaError(VbaErrorCode.FILE_NOT_FOUND, 'File not found');
                        }
                        throw e;
                    }
                    streamPos = 0;
                },
                type: 2,
                charset: 'utf-8',
                get position() { ensureOpen(); return streamPos; },
                set position(value: any) {
                    ensureOpen();
                    if (value === vbaNull) evaluator.throwVbaError(VbaErrorCode.INVALID_USE_OF_NULL, 'Invalid use of Null');
                    streamPos = Math.max(0, Math.min(content.length, Math.trunc(Number(value))));
                },
                get size() { ensureOpen(); return content.length; }
            };
        });
    }

    // Execute a sequence of statements starting from startIndex, with error handling support.
    // GoSub is handled in every block using the owning procedure body, so a
    // Return can resume the statement following the GoSub even inside loops or
    // conditional clauses. Other procedure-level control flow still bubbles
    // from nested blocks to their owning procedure executor.
    // inlineClause=true: 単一行 If の Then/Else 節。節内の文は「If 文の一部」なので、
    // GoTo ハンドラー処理を行わず外側へバブルさせる（Resume は If 文全体を再実行する）。
    private executeStatements(body: Statement[], startIndex: number, isTopLevel: boolean = true, inlineClause: boolean = false) {
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
                if (!isTopLevel && e && (e.type === 'GoTo' || e.type === 'Return' || e.type === 'Resume')) {
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
                    this.executeGoSubLabel(e.label);
                    i++;
                    continue;
                }

                if (e && e.type === 'Return') {
                    if (this.gosubTargetDepth > 0) {
                        throw e;
                    }
                    {
                        this.throwVbaError(VbaErrorCode.RETURN_WITHOUT_GOSUB, 'Return without GoSub');
                    }
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
                    this.clearErrState();
                    continue;
                }

                // Handle VbaError or standard JS Error
                if (!this.isInErrorHandler) {
                    // vbaBareMessage があれば "Run-time error 'N': ... (line X)" の枠組みを
                    // 含まない生のメッセージを使う（throwVbaError 経由のエラー）。
                    // Err.Raise も invokeBuiltin で枠組み付き Error に包み直されるため
                    // vbaBareMessage を持つ（Err.Description には生メッセージが入る）。
                    this.recordErrState(e);

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
                        // ラベルがこのブロックにない = ネストブロック（For/If 等）内のエラー。
                        // ループフレームを巻き戻すと Resume Next が失敗文の次ではなく
                        // ブロック全体の次に飛んでしまうため、フレームを保ったまま
                        // プロシージャ本体のハンドラーをここから実行し、Resume を受け取る。
                        const procBody = inlineClause ? null : this.currentProcBody;
                        const topIndex = procBody && procBody !== body ? procBody.findIndex(s =>
                            s.type === 'LabelStatement' &&
                            (s as any).label.toLowerCase() === labelName
                        ) : -1;
                        if (procBody && topIndex >= 0) {
                            this.isInErrorHandler = true;
                            let resumeInfo: any = null;
                            let fellOffEnd = false;
                            try {
                                // isTopLevel=false: ハンドラー内の Resume/GoTo をこの catch まで伝播させる
                                this.executeStatements(procBody, topIndex, false);
                                fellOffEnd = true;
                            } catch (r: any) {
                                if (r && r.type === 'Resume') resumeInfo = r;
                                else throw r; // Exit / 再エラー / GoTo はそのまま上位へ
                            }
                            if (fellOffEnd) {
                                // ハンドラーが End Sub に到達 → プロシージャを正常終了させる
                                throw { type: 'Exit', target: 'Sub' };
                            }
                            this.isInErrorHandler = false;
                            this.clearErrState();
                            if (resumeInfo.mode === 'Current') continue;      // 失敗文を再実行
                            if (resumeInfo.mode === 'Next') { i++; continue; } // 失敗文の次へ
                            // Resume <label>: ラベルはプロシージャレベルにあるため GoTo として伝播
                            throw { type: 'GoTo', label: resumeInfo.label };
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
        for (const target of stmt.names) {
            let arr: any;
            let setValue: (value: any) => void;
            let typeInfoEnv: Environment | undefined;
            let typeInfoName: string | undefined;
            if (target.type === 'Identifier') {
                const varName = (target as Identifier).name;
                arr = this.env.get(varName);
                setValue = (value) => this.env.set(varName, value);
                typeInfoEnv = this.env;
                typeInfoName = varName;
            } else if (target.type === 'MemberExpression') {
                const member = target as MemberExpression;
                const obj = this.resolveAutoInstance(member.object, this.evaluateExpression(member.object));
                const propName = member.property.name.toLowerCase();
                if (obj?.__vbaClass__) {
                    typeInfoEnv = obj.__instanceEnv__ as Environment;
                    arr = typeInfoEnv.get(propName);
                    setValue = (value) => typeInfoEnv!.set(propName, value);
                    typeInfoName = propName;
                } else {
                    arr = obj?.[propName];
                    setValue = (value) => { obj[propName] = value; };
                }
            } else if (target.type === 'ImplicitWithObjectExpression') {
                if (this.withObjectStack.length === 0) this.throwVbaError(91, 'Object variable or With block variable not set');
                const obj = this.withObjectStack[this.withObjectStack.length - 1];
                const propName = (target as ImplicitWithObjectExpression).property.name.toLowerCase();
                if (obj?.__vbaClass__) {
                    typeInfoEnv = obj.__instanceEnv__ as Environment;
                    arr = typeInfoEnv.get(propName);
                    setValue = (value) => typeInfoEnv!.set(propName, value);
                    typeInfoName = propName;
                } else {
                    arr = obj?.[propName];
                    setValue = (value) => { obj[propName] = value; };
                }
            } else {
                this.throwVbaError(VbaErrorCode.INVALID_PROCEDURE_CALL, 'Invalid procedure call or argument');
            }
            if (Array.isArray(arr)) {
                if ((arr as any).vbaFixed) {
                    const defaultValue = (arr as any).__vbaDefaultValue__
                        ?? ((arr as any).__vbaElementObjectTypeName__ ? vbaNothing : 0);
                    const autoNewType = (arr as any).__vbaElementAutoNew__
                        ? (arr as any).__vbaElementObjectTypeName__
                        : undefined;
                    this.reinitializeArray(arr, defaultValue, autoNewType);
                } else {
                    // Dynamic array: Erase deallocates (uninitialized). Null signals this
                    // so UBound/LBound/access throw Error 9 until ReDim is called again.
                    if (typeInfoEnv && typeInfoName) typeInfoEnv.setArrayTypeInfo(typeInfoName, arr);
                    setValue(null);
                }
            }
        }
    }

    private reinitializeArray(arr: any[], defaultValue: any, autoNewType?: string) {
        for (let i = 0; i < arr.length; i++) {
            if (Array.isArray(arr[i]) && !(arr[i] as any).__vbaClass__) {
                this.reinitializeArray(arr[i], defaultValue, autoNewType);
            } else {
                arr[i] = autoNewType ? createAutoInstancePlaceholder(autoNewType) : defaultValue;
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

    /** Fill an `As New` object array with one lazy placeholder per element. */
    private fillArrayWithAutoInstancePlaceholders(arr: any[], typeName: string, skipExisting = false) {
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
                    if (!skipExisting || a[i] === undefined || a[i] === vbaNothing) {
                        a[i] = createAutoInstancePlaceholder(typeName);
                    }
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
        for (const decl of stmt.declarations) {
            this.evaluateReDimDeclarator(decl, stmt.isPreserve);
        }
    }

    private evaluateReDimDeclarator(decl: ReDimDeclarator, isPreserve: boolean) {
        // Resolve get/set accessors for the three target forms
        let oldArr: any;
        let setNewArr: (arr: any[]) => void;
        let storedArrayTypeInfo: { elementType?: string; elementTypeName?: string; elementObjectTypeName?: string; elementAutoNew?: boolean } | undefined;
        if (decl.name.type === 'Identifier') {
            const varName = (decl.name as Identifier).name;
            oldArr = this.env.get(varName);
            setNewArr = (arr) => this.env.set(varName, arr);
            storedArrayTypeInfo = this.env.getArrayTypeInfo(varName);
        } else if (decl.name.type === 'MemberExpression') {
            const mem = decl.name as MemberExpression;
            const obj = this.resolveAutoInstance(mem.object, this.evaluateExpression(mem.object));
            const propName = (mem.property as Identifier).name.toLowerCase();
            if (obj?.__vbaClass__) {
                const instanceEnv = obj.__instanceEnv__ as Environment;
                oldArr = this.getClassField(instanceEnv, propName);
                setNewArr = (arr) => instanceEnv.setLocally(propName, arr);
                storedArrayTypeInfo = instanceEnv.getOwnArrayTypeInfo(propName);
            } else {
                oldArr = obj?.[propName];
                setNewArr = (arr) => { obj[propName] = arr; };
            }
        } else {
            // ImplicitWithObjectExpression: .Items inside With block
            const prop = (decl.name as ImplicitWithObjectExpression).property as Identifier;
            if (this.withObjectStack.length === 0) this.throwVbaError(91, 'Object variable or With block variable not set');
            const obj = this.withObjectStack[this.withObjectStack.length - 1];
            const propName = prop.name.toLowerCase();
            if (obj?.__vbaClass__) {
                const instanceEnv = obj.__instanceEnv__ as Environment;
                oldArr = this.getClassField(instanceEnv, propName);
                setNewArr = (arr) => instanceEnv.setLocally(propName, arr);
                storedArrayTypeInfo = instanceEnv.getOwnArrayTypeInfo(propName);
            } else {
                oldArr = obj?.[propName];
                setNewArr = (arr) => { obj[propName] = arr; };
            }
        }

        const returnArrayType = decl.name.type === 'Identifier' && this.currentProcedureName && this.currentProcedureReturnsArray &&
            (decl.name as Identifier).name.toLowerCase() === this.currentProcedureName.toLowerCase()
            ? this.currentProcedureReturnType
            : undefined;

        // UDT 配列の場合、Dim 時に保存した要素型名を引き継ぐ
        const elementTypeName: string | undefined =
            (Array.isArray(oldArr) ? (oldArr as any).__vbaElementTypeName__ : undefined) ??
            storedArrayTypeInfo?.elementTypeName ??
            (decl.objectType && this.env.getType(decl.objectType) ? decl.objectType : undefined) ??
            (returnArrayType && this.env.getType(returnArrayType) ? returnArrayType : undefined);
        const elementType: string | undefined =
            (Array.isArray(oldArr) ? (oldArr as any).__vbaElementType__ : undefined) ??
            storedArrayTypeInfo?.elementType ??
            (decl.objectType && !this.env.getType(decl.objectType) ? decl.objectType.toLowerCase() : undefined) ??
            (returnArrayType && !this.env.getType(returnArrayType) ? returnArrayType.toLowerCase() : undefined);
        const elementObjectTypeName: string | undefined =
            (Array.isArray(oldArr) ? (oldArr as any).__vbaElementObjectTypeName__ : undefined) ??
            storedArrayTypeInfo?.elementObjectTypeName ??
            (decl.objectType && (
                this.classDefinitions.has(decl.objectType.toLowerCase()) ||
                this.externalObjectFactories.has(decl.objectType.toLowerCase()) ||
                ['object', 'collection'].includes(decl.objectType.toLowerCase())
            ) ? decl.objectType : undefined) ??
            (returnArrayType && (
                this.classDefinitions.has(returnArrayType.toLowerCase()) ||
                this.externalObjectFactories.has(returnArrayType.toLowerCase()) ||
                ['object', 'collection'].includes(returnArrayType.toLowerCase())
            ) ? returnArrayType : undefined);
        const elementAutoNew = (Array.isArray(oldArr) && (oldArr as any).__vbaElementAutoNew__)
            || !!storedArrayTypeInfo?.elementAutoNew;

        let defaultValue: any = 0;
        if (decl.objectType) {
            const t = decl.objectType.toLowerCase();
            if (t === 'string') defaultValue = '';
            else if (t === 'boolean') defaultValue = 0;
            else if (
                this.classDefinitions.has(t) || this.externalObjectFactories.has(t) ||
                t === 'object' || t === 'collection'
            ) defaultValue = vbaNothing;
        } else if (returnArrayType && (
            this.classDefinitions.has(returnArrayType.toLowerCase()) ||
            this.externalObjectFactories.has(returnArrayType.toLowerCase()) ||
            ['object', 'collection'].includes(returnArrayType.toLowerCase())
        )) {
            defaultValue = vbaNothing;
        } else if (Array.isArray(oldArr)) {
            defaultValue = (oldArr as any).__vbaDefaultValue__
                ?? ((oldArr as any).__vbaElementObjectTypeName__ ? vbaNothing : 0);
        } else if (storedArrayTypeInfo?.elementObjectTypeName) {
            defaultValue = vbaNothing;
        }

        if (decl.bounds.length > 0) {
            if (isPreserve && Array.isArray(oldArr) && (oldArr as any).__vbaDimensions__) {
                const oldDims = (oldArr as any).__vbaDimensions__ as { lower: number, upper: number }[];
                const newDims = decl.bounds.map(bound => {
                    const lower = bound.lower ? this.evaluateExpression(bound.lower) : 0;
                    const upper = this.evaluateExpression(bound.upper);
                    return { lower, upper };
                });
                if (newDims.length !== oldDims.length) {
                    this.throwVbaError(VbaErrorCode.SUBSCRIPT_OUT_OF_RANGE, "Subscript out of range");
                }
                for (let i = 0; i < newDims.length; i++) {
                    if (newDims[i].lower !== oldDims[i].lower) {
                        this.throwVbaError(VbaErrorCode.SUBSCRIPT_OUT_OF_RANGE, "Subscript out of range");
                    }
                }
                for (let i = 0; i < newDims.length - 1; i++) {
                    if (newDims[i].upper !== oldDims[i].upper) {
                        this.throwVbaError(VbaErrorCode.SUBSCRIPT_OUT_OF_RANGE, "Subscript out of range");
                    }
                }
            }

            const arr = this.createMultiDimArray(decl.bounds, defaultValue);
            (arr as any).vbaFixed = false;
            if (elementType) (arr as any).__vbaElementType__ = elementType;

            if (elementTypeName) {
                (arr as any).__vbaElementTypeName__ = elementTypeName;
                if (!isPreserve) {
                    this.fillArrayWithUDT(arr, (arr as any).__vbaDimensions__, 0, elementTypeName);
                }
            }
            if (elementObjectTypeName) {
                (arr as any).__vbaElementObjectTypeName__ = elementObjectTypeName;
            }
            if (elementAutoNew && elementObjectTypeName && this.classDefinitions.has(elementObjectTypeName.toLowerCase())) {
                (arr as any).__vbaElementAutoNew__ = true;
            }

            if (isPreserve && Array.isArray(oldArr)) {
                this.copyPreservedData(oldArr, arr, (arr as any).__vbaDimensions__);
                if (elementTypeName) {
                    this.fillArrayWithUDT(arr, (arr as any).__vbaDimensions__, 0, elementTypeName, true);
                }
            }
            if (elementAutoNew && elementObjectTypeName &&
                this.classDefinitions.has(elementObjectTypeName.toLowerCase())) {
                this.fillArrayWithAutoInstancePlaceholders(
                    arr,
                    elementObjectTypeName,
                    isPreserve,
                );
            }

            setNewArr(arr);
            if (decl.name.type === 'Identifier') {
                this.env.setArrayTypeInfo((decl.name as Identifier).name, arr);
            } else if (decl.name.type === 'MemberExpression') {
                const member = decl.name as MemberExpression;
                const obj = this.resolveAutoInstance(member.object, this.evaluateExpression(member.object));
                if (obj?.__vbaClass__) (obj.__instanceEnv__ as Environment).setArrayTypeInfo(member.property.name, arr);
            } else if (decl.name.type === 'ImplicitWithObjectExpression') {
                const obj = this.withObjectStack[this.withObjectStack.length - 1];
                if (obj?.__vbaClass__) {
                    (obj.__instanceEnv__ as Environment).setArrayTypeInfo((decl.name as ImplicitWithObjectExpression).property.name, arr);
                }
            }
        }
    }

    private fillArrayWithUDT(arr: any[], dimensions: { lower: number, upper: number }[], dimIdx: number, typeName: string, skipExisting = false) {
        const { lower, upper } = dimensions[dimIdx];
        if (dimIdx < dimensions.length - 1) {
            for (let i = lower; i <= upper; i++) {
                this.fillArrayWithUDT(arr[i], dimensions, dimIdx + 1, typeName, skipExisting);
            }
        } else {
            for (let i = lower; i <= upper; i++) {
                if (!skipExisting || arr[i] === undefined || arr[i] === 0) {
                    arr[i] = this.instantiateType(typeName);
                }
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
                if (!Number.isFinite(lit.value)) {
                    this.throwVbaError(VbaErrorCode.OVERFLOW, 'Overflow');
                }
                return this.applyLiteralTypeSuffix(lit.value, lit.typeSuffix, lit.rawIntegerText);
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
                // env.get() は暗黙初期化で変数を作るため、存在チェックは先に取る
                const hadVariable = this.env.hasVariable(idName);
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
                // クラスメソッド内の非修飾自メンバー参照（Bug 33-A）:
                // 変数にもプロシージャにも解決できない場合、Me のクラスメンバー
                // （Property Get / Function、必須引数 0）を暗黙の Me.<name> として呼ぶ
                if (!p && !this.inConstEval && !hadVariable) {
                    const me = this.env.getConst('me');  // getConst: 暗黙初期化なし
                    if (me && me.__vbaClass__ && me.__classDef__) {
                        const isCurrentProcReturn = this.currentProcedureName &&
                            idName.toLowerCase() === this.currentProcedureName.toLowerCase();
                        if (!isCurrentProcReturn) {
                            const lower = idName.toLowerCase();
                            const member = (me.__classDef__ as ClassDeclaration).procedures.find(cp =>
                                cp.name.name.toLowerCase() === lower &&
                                (cp.isFunction || (cp.isProperty && cp.propertyType === 'get')) &&
                                cp.parameters.filter(pp => !pp.isOptional && !pp.isParamArray).length === 0
                            );
                            if (member) {
                                return this.callClassMethod(me, member, []);
                            }
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
            case 'ImplicitWithDictionaryAccessExpression':
                return this.evaluateImplicitWithDictionaryAccessExpression(expr as ImplicitWithDictionaryAccessExpression);
            case 'ParenthesizedExpression':
                return this.evaluateExpression((expr as ParenthesizedExpression).expression);
            case 'NewExpression':
                return this.instantiateClass((expr as NewExpression).className);
            case 'NamedArgument':
                // Named arguments should only appear as call arguments, but evaluate the value if encountered
                return this.evaluateExpression((expr as NamedArgument).value);
            case 'ByValArgument':
                return this.evaluateExpression((expr as ByValArgument).value);
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
        // Time-only literal (e.g. #12:30:45#): VBA uses the zero date 1899/12/30
        if (parts.length === 0 && timeMatch !== null) {
            return { year: 1899, month: 12, day: 30, hour, minute, second };
        }
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

        // VBA's two-digit-year rule: 00-29 -> 2000-2029; 30-99 -> 1930-1999.
        if (year < 100) {
            year += year <= 29 ? 2000 : 1900;
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
    private applyLiteralTypeSuffix(val: number, suffix: NumberLiteral['typeSuffix'], rawIntegerText?: string): any {
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
            case '@':
                return VbaCurrency.fromNumber(this.vbaRound(val, 4));
            case '^':
                // TypeName → 'LongLong' は AST の typeSuffix 経由で返る。
                // Preserve the source digits before JavaScript Number loses
                // precision, then let declared-type coercion validate the range.
                if (rawIntegerText !== undefined) return BigInt(rawIntegerText);
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
        if (lit.baseWidth === 16) return 'Integer';
        if (lit.baseWidth === 32) return 'Long';
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
        if (lit.baseWidth === 16) return 2; // vbInteger
        if (lit.baseWidth === 32) return 3; // vbLong
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
        'datediff': 'Long',
        // Integer を返す関数
        'asc': 'Integer', 'ascb': 'Integer', 'ascw': 'Integer',
        'vartype': 'Integer',
        'year': 'Integer', 'month': 'Integer', 'day': 'Integer',
        'hour': 'Integer', 'minute': 'Integer', 'second': 'Integer',
        'weekday': 'Integer',
        'sgn': 'Integer', 'strcomp': 'Integer', 'datepart': 'Integer',
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
                const subtype = this.resolveNumericSubtype(argExpr);
                if (subtype) return subtype;
                return 'Double';
            }
            return this.env.get('typename')(val);
        } else {
            if (typeof val === 'number') {
                if (argExpr.type === 'NumberLiteral') return this.inferLiteralVarType(argExpr as NumberLiteral);
                const subtype = this.resolveNumericSubtype(argExpr);
                if (subtype) return vtMap[subtype] ?? 5;
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

    /**
     * Determines the VBA numeric subtype of an expression at the AST level.
     * Used to track the dynamic subtype of Variant variables at assignment time.
     * Returns undefined when the subtype cannot be determined statically.
     */
    private resolveNumericSubtype(expr: Expression): VbaVarType | undefined {
        // 括弧は型に影響しない: (2) は Integer のまま（実 VBA 差分コーパスで発見）
        if (expr.type === 'ParenthesizedExpression') {
            return this.resolveNumericSubtype((expr as ParenthesizedExpression).expression);
        }
        if (expr.type === 'NumberLiteral') {
            return this.inferLiteralTypeName(expr as NumberLiteral) as VbaVarType;
        }
        // True/False/Empty リテラルと日付リテラルの静的な型（実 VBA 差分で裁定）
        if (expr.type === 'Identifier') {
            const n0 = (expr as Identifier).name.toLowerCase();
            if (n0 === 'true' || n0 === 'false' || n0 === 'empty') return 'Integer';
        }
        if (expr.type === 'DateLiteral') return 'Date';
        // Int/Fix は引数の数値型を保存する（Int(256)→Integer、Int(1E10)→Double。実 VBA 差分で裁定）
        if (expr.type === 'CallExpression') {
            const ce0 = expr as CallExpression;
            if (ce0.callee.type === 'Identifier' && ce0.args.length === 1) {
                const fn = (ce0.callee as Identifier).name.toLowerCase();
                if (fn === 'int' || fn === 'fix') {
                    const inner = this.resolveNumericSubtype(ce0.args[0]);
                    if (inner) return inner;
                    return 'Double';  // 文字列等、静的に型が分からない引数は Double
                }
            }
        }
        if (expr.type === 'BinaryExpression') {
            return this.resolveBinaryExprNumericSubtype(expr as BinaryExpression);
        }
        if (expr.type === 'UnaryExpression') {
            const ue = expr as UnaryExpression;
            if (ue.operator === '-' || ue.operator === '+') {
                return this.resolveNumericSubtype(ue.argument);
            }
        }
        const declared = this.resolveDeclaredReturnType(expr);
        if (declared && declared !== 'Variant' && declared !== 'Object') return declared;
        // Propagate subtype from another Variant variable
        if (expr.type === 'Identifier') {
            return this.env.getVariantSubtype((expr as Identifier).name);
        }
        // Bug CG: Look up subtype from Variant array element (arr(i))
        if (expr.type === 'CallExpression') {
            const ce = expr as CallExpression;
            if (ce.callee.type === 'Identifier') {
                const arrName = (ce.callee as Identifier).name;
                const arr = this.env.get(arrName);
                if (Array.isArray(arr) && (arr as any).__vbaSubtypes__) {
                    const idxs = this.evaluateIndexExpressions(ce.args);
                    return (arr as any).__vbaSubtypes__[idxs.join(',')];
                }
            }
        }
        return undefined;
    }

    /**
     * VBA 型昇格規則に従って BinaryExpression の数値サブタイプを解決する。
     *
     * 昇格階層（+, -, * の場合）: Byte < Integer < Long < LongLong < Single < Currency < Double
     * / は常に Double（両辺 Single のみ Single）
     * \ と Mod は整数除算結果（Integer または Long）
     * ^ は常に Double
     */
    private resolveBinaryExprNumericSubtype(expr: BinaryExpression): VbaVarType | undefined {
        const op = expr.operator.toLowerCase();

        if (op === '^') return 'Double';

        const lt = this.resolveNumericSubtype(expr.left);
        const rt = this.resolveNumericSubtype(expr.right);

        // 数値文字列オペランドは Double として演算される（実 VBA 差分で裁定）:
        // (2) + ("7") → Double。\ と Mod は丸め後の整数演算なので Long
        const isStringOperand = (e: Expression): boolean => {
            if (e.type === 'ParenthesizedExpression') return isStringOperand((e as ParenthesizedExpression).expression);
            if (e.type === 'StringLiteral') return true;
            if (e.type === 'Identifier') {
                const ti = this.env.getVariableType((e as Identifier).name);
                return ti?.vbaType === 'String';
            }
            return false;
        };
        const logicOps = ['and', 'or', 'xor', 'eqv', 'imp'];
        if (['+', '-', '*', '/', '\\', 'mod', ...logicOps].includes(op) && (isStringOperand(expr.left) || isStringOperand(expr.right))) {
            // 論理演算子も \/Mod と同様、整数（Long）へ丸めてからビット演算される
            // （実 VBA 差分で裁定: True And "7" は Double ではなく Long）
            return (op === '\\' || op === 'mod' || logicOps.includes(op)) ? 'Long' : 'Double';
        }

        if (op === '/') {
            // / は常に Double（両辺が Single のときのみ Single）
            const singleCompatible = new Set<VbaVarType | undefined>(['Single', 'Byte', 'Integer']);
            if (singleCompatible.has(lt) && singleCompatible.has(rt) && (lt === 'Single' || rt === 'Single')) return 'Single';
            return 'Double';
        }

        // \/Mod/論理演算子共通: 非整数型は Long に丸めてから演算される
        const toIntType = (t: VbaVarType | undefined): 'Integer' | 'Long' | 'LongLong' | undefined => {
            if (t === 'Byte' || t === 'Integer') return 'Integer';
            if (t === 'Long') return 'Long';
            if (t === 'LongLong') return 'LongLong';
            if (t === 'Single' || t === 'Double' || t === 'Currency' || t === 'Date') return 'Long'; // 丸め後 Long
            return undefined;
        };
        if (op === '\\' || op === 'mod' || logicOps.includes(op)) {
            const nl = toIntType(lt), nr = toIntType(rt);
            if (!nl || !nr) return undefined;
            if (nl === 'LongLong' || nr === 'LongLong') return 'LongLong';
            if (nl === 'Long' || nr === 'Long') return 'Long';
            return 'Integer';
        }

        // +, -, * の型昇格
        if (op === '+' || op === '-' || op === '*') {
            // Date は + / - では評価結果側で Date 型が維持されるが、* では数値（Double）になる。
            // ただし Date - Date は数値（Double）を返す
            if (op === '-' && lt === 'Date' && rt === 'Date') return 'Double';
            if ((op === '+' || op === '-') && (lt === 'Date' || rt === 'Date')) return 'Date';
            if (op === '*' && (lt === 'Date' || rt === 'Date')) return 'Double';
            // どちらかが Double → Double
            if (lt === 'Double' || rt === 'Double') return 'Double';
            // どちらかが Currency → Currency（ただし Double との組み合わせは上で処理済み）
            if (lt === 'Currency' || rt === 'Currency') return 'Currency';
            // Long + Single / LongLong + Single は Double（Single の精度不足）
            const isLongish = (t: VbaVarType | undefined) => t === 'Long' || t === 'LongLong';
            if ((isLongish(lt) && rt === 'Single') || (lt === 'Single' && isLongish(rt))) return 'Double';
            // どちらかが Single → Single
            if (lt === 'Single' || rt === 'Single') return 'Single';
            // どちらかが LongLong → LongLong
            if (lt === 'LongLong' || rt === 'LongLong') return 'LongLong';
            // どちらかが Long → Long
            if (lt === 'Long' || rt === 'Long') return 'Long';
            // Byte + Byte → Integer（オーバーフロー防止のため VBA 仕様上 Integer に昇格）。
            // 片方の型が静的に不明な場合は Integer と断定しない（安全側に倒す。
            // 例: Sgn(x)（Integer 確定）* Abs(x)（型不明）を誤って Integer 扱いすると
            // オーバーフロー判定が過検出になる）
            const isIntegerish = (t: VbaVarType | undefined) => t === 'Integer' || t === 'Byte';
            if (isIntegerish(lt) && isIntegerish(rt)) return 'Integer';
            // 両辺が不明、または片方だけ不明 → 不明
            return undefined;
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

            // CallByName receives its argument expressions through a ParamArray. For
            // VBA class procedures, preserve those expressions so ByRef updates are
            // written back to the caller instead of being lost in evaluated values.
            if (nameLower === 'callbyname' && expr.args.length >= 3) {
                const target = this.resolveAutoInstance(expr.args[0], this.evaluateExpression(expr.args[0]));
                if (target?.__vbaClass__) {
                    const procName = vbaToString(this.evaluateExpression(expr.args[1])).toLowerCase();
                    const callType = this.evaluateExpression(expr.args[2]);
                    if (callType === vbaNull) this.throwVbaError(VbaErrorCode.INVALID_USE_OF_NULL, 'Invalid use of Null');
                    const propertyArgs = expr.args.slice(3);
                    const evaluatedPropertyArgs = callType === 8
                        ? this.evaluateExpressions(propertyArgs)
                        : undefined;
                    if (callType === 8 && evaluatedPropertyArgs && evaluatedPropertyArgs.length > 0 &&
                        !isVbaObjectReferenceCompatible(evaluatedPropertyArgs[evaluatedPropertyArgs.length - 1])) {
                        this.throwVbaError(VbaErrorCode.OBJECT_REQUIRED, 'Object required');
                    }
                    const procedures = (target.__classDef__ as ClassDeclaration).procedures;
                    const propertyType = callType === 4 ? 'let' : callType === 8 ? 'set' : undefined;
                    const proc = propertyType
                        ? findClassProperty(procedures, procName, propertyType)
                        : (callType === 2
                            ? findClassProperty(procedures, procName, 'get')
                            : procedures.find(p => !p.isProperty && p.name.name.toLowerCase() === procName));
                    if (proc) {
                        this.checkNoGapOnRequiredParam(proc.parameters, propertyArgs);
                        return this.callClassMethodWithExpressions(target, proc, propertyArgs, evaluatedPropertyArgs);
                    }
                    this.throwVbaError(VbaErrorCode.OBJECT_DOESNT_SUPPORT_PROPERTY, `Object doesn't support this property or method: '${procName}'`);
                }
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
                    this.vbaCallStack.push({ name: classProc.name.name, moduleName: me.__className__ ?? '', line: this.currentLine });
                    try {
                        return this.callClassMethodWithExpressions(me, classProc, expr.args);
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
                const byRefArgs: { paramName: string, reference: VbaLValueReference }[] = [];
                let paramArrayParamName: string | null = null;
                let paramArrayReferences: (VbaLValueReference | null)[] = [];
                const namedArgs = new Map<string, any>();
                const namedArgExpressions = new Map<string, Expression>();
                const positionalArgs: any[] = [];
                const positionalArgExpressions: Expression[] = [];

                const splitArgs = this.splitArgumentExpressions(expr.args);
                for (const name of splitArgs.named.keys()) {
                    if (!proc.parameters.some(parameter => parameter.name.toLowerCase() === name)) {
                        this.throwVbaError(448, `Named argument not found: '${name}'`);
                    }
                }
                for (const entry of splitArgs.ordered) {
                    const name = entry.name;
                    const argExpr = entry.expression;
                    if (name === undefined) {
                        const reference = this.createLValueReference(argExpr);
                        positionalArgs.push(reference
                            ? reference.get()
                            : this.resolveAutoInstance(argExpr, this.evaluateExpression(argExpr)));
                        positionalArgExpressions.push(argExpr);
                        continue;
                    }
                    // As New auto-instance は引数として渡す時点で呼び出し元の変数に実体化する
                    // （Bug 33-C: ByVal だと callee 側実体化が呼び出し元に反映されない）
                    const reference = this.createLValueReference(argExpr);
                    namedArgs.set(name, reference
                        ? reference.get()
                        : this.resolveAutoInstance(argExpr, this.evaluateExpression(argExpr)));
                    namedArgExpressions.set(name, argExpr);
                }
                // Variant サブタイプは呼び出し元 env で引数式から解決しておき、
                // パラメーターへ伝播する（実 VBA 差分: TypeName(引数) が Double に化けるのを防ぐ）
                const positionalSubtypes = positionalArgs.map((v, i) =>
                    typeof v === 'number' ? this.resolveNumericSubtype(positionalArgExpressions[i]) : undefined);
                const namedSubtypes = new Map<string, VbaVarType | undefined>();
                for (const [k, e] of namedArgExpressions) {
                    const v = namedArgs.get(k);
                    namedSubtypes.set(k, typeof v === 'number' ? this.resolveNumericSubtype(e) : undefined);
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
                        paramArrayReferences = positionalArgExpressions.slice(i)
                            .map(expr => this.createLValueReference(expr));
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
                    this.validateArrayParameterContainer(argVal, param);
                    const originalArgExpr = namedArgExpressions.has(paramNameLower)
                        ? namedArgExpressions.get(paramNameLower)
                        : (i < positionalArgExpressions.length ? positionalArgExpressions[i] : undefined);
                    // Parenthesizing an argument forces a ByVal temporary in VBA,
                    // even when the procedure declares the parameter ByRef.
                    const forcedByVal = originalArgExpr?.type === 'ParenthesizedExpression' ||
                        originalArgExpr?.type === 'ByValArgument';
                    // Register parameter type metadata (but not for array parameters)
                    if (param.paramType && !param.isArray) {
                        const typeMap: Record<string, VbaVarType> = {
                            'byte': 'Byte', 'integer': 'Integer', 'long': 'Long',
                            'single': 'Single', 'double': 'Double', 'currency': 'Currency',
                            'string': 'String', 'boolean': 'Boolean', 'date': 'Date',
                        };
                        const mapped = typeMap[param.paramType.toLowerCase()];
                        if (mapped) {
                            // Bug CH: include fixedLength so setLocally coerces String * N parameters
                            localEnv.setVariableType(param.name, {
                                vbaType: mapped,
                                fixedLength: mapped === 'String' ? param.fixedLength : undefined,
                            });
                        } else if (this.classDefinitions.has(param.paramType.toLowerCase())) {
                            localEnv.setVariableType(param.name, {
                                vbaType: 'Object',
                                objectTypeName: param.paramType,
                            });
                        }
                    }
                    // ByVal values must not share mutable VBA arrays or UDTs with
                    // the caller. Class instances remain references, as in VBA.
                    argVal = this.prepareArgumentValue(argVal, param, forcedByVal);
                    localEnv.setLocally(param.name, argVal);
                    // Variant（型なし）パラメーターへ数値サブタイプを伝播
                    if (!param.paramType || param.paramType.toLowerCase() === 'variant') {
                        const st = namedArgs.has(paramNameLower)
                            ? namedSubtypes.get(paramNameLower)
                            : (i < positionalSubtypes.length && !isMissingSlot ? positionalSubtypes[i] : undefined);
                        if (st) localEnv.setVariantSubtype(param.name, st);
                    }
                    this.addRef(argVal);

                    // ByRef handling
                    if (!param.isByVal) {
                        let originalExpr: Expression | undefined;
                        if (namedArgExpressions.has(paramNameLower)) {
                            originalExpr = namedArgExpressions.get(paramNameLower);
                        } else if (i < positionalArgExpressions.length) {
                            originalExpr = positionalArgExpressions[i];
                        }
                        const reference = this.argumentReference(originalExpr, param, forcedByVal);
                        if (reference) {
                            byRefArgs.push({
                                paramName: param.name,
                                reference,
                            });
                        }
                    }
                }

                return this.execProcBody(proc, localEnv, {
                    byRefArgs,
                    paramArrayParamName,
                    paramArrayReferences,
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
                    return this.invokeBuiltin(variable, this.resolveCallArgs(variable, expr.args, name));
                } else if (Array.isArray(variable)) {
                    if (expr.args.length === 0) this.throwVbaError(VbaErrorCode.SUBSCRIPT_OUT_OF_RANGE, 'Subscript out of range');
                    const dims = (variable as any).__vbaDimensions__ as { lower: number, upper: number }[] | undefined;
                    const indexes = this.evaluateIndexExpressions(expr.args);
                    return this.readArrayAtIndexes(variable, indexes, dims, true, true);
                } else if (variable && variable.__isVbaDict__) {
                    // Dictionary read: dict("key")
                    if (expr.args.length === 0) this.throwVbaError(VbaErrorCode.ARGUMENT_NOT_OPTIONAL, 'Argument not optional');
                    const key = variable.__resolveKey__(this.evaluateExpression(expr.args[0]));
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
                        return this.callClassMethodWithExpressions(variable, defaultProperty, expr.args);
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
                                const argsVals = this.resolveCallArgs(tier6Member, expr.args, name);
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
                                return this.invokeBuiltin(builtin, this.resolveCallArgs(builtin, expr.args, member.property.name));
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
                            const aligned = this.alignProcedureCallExpressions(qualifiedProc, expr.args);
                            // Variant サブタイプを呼び出し元 env で解決して伝播
                            const subtypes = aligned.args.map((v, i) =>
                                typeof v === 'number' ? this.resolveNumericSubtype(aligned.expressions[i] ?? expr.args[i]) : undefined);
                            const result = this.callProcedure(member.property.name, aligned.args, undefined, possibleModuleName, subtypes);
                            for (let i = 0; i < qualifiedProc.parameters.length && i < aligned.args.length; i++) {
                                if (!qualifiedProc.parameters[i].isByVal && aligned.references[i]) {
                                    aligned.references[i]!.set(aligned.args[i]);
                                }
                            }
                            return result;
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
                // Bug CE: In a read/call context, Property Get must take priority over Property Set/Let
                // when both share the same name. `classDef.procedures.find()` returns whichever was
                // declared first, so if Property Set precedes Property Get, the setter is incorrectly
                // invoked on a call like `w.Inner("key")`.
                const getter = classDef.procedures.find(
                    p => p.name.name.toLowerCase() === methodNameLower && p.isProperty && p.propertyType === 'get'
                );
                const setter = classDef.procedures.find(
                    p => p.name.name.toLowerCase() === methodNameLower &&
                        p.isProperty && (p.propertyType === 'let' || p.propertyType === 'set')
                );
                // A call-style Property Let can supply its value by name (for example,
                // Call obj.Value(Value:=x)). Prefer the setter when a named argument
                // exists only in the setter; ordinary property reads keep Getter priority.
                const namedArgs = expr.args
                    .filter(arg => arg.type === 'NamedArgument')
                    .map(arg => (arg as NamedArgument).name.toLowerCase());
                const getterNames = new Set((getter?.parameters ?? []).map(param => param.name.toLowerCase()));
                const setterNames = new Set((setter?.parameters ?? []).map(param => param.name.toLowerCase()));
                const setterOnlyNamedArg = namedArgs.some(name => setterNames.has(name) && !getterNames.has(name));
                const proc = setterOnlyNamedArg ? setter : getter ?? setter ??
                    classDef.procedures.find(p => p.name.name.toLowerCase() === methodNameLower);
                if (proc) {
                    // Bug BZ: Property Get with 0 params but called with args → get the object first,
                    // then subscript/index the returned value with the given args.
                    // e.g. c.TheObj("k") where Property Get TheObj() As Object returns a Dictionary.
                    if (proc.isProperty && proc.propertyType === 'get' &&
                        proc.parameters.length === 0 && expr.args.length > 0) {
                        const returned = this.callClassMethod(obj, proc, []);
                        const argsVals = this.evaluateCallArgumentValues(expr.args);
                        if (returned && returned.__isVbaDict__) {
                            return returned.__map__.get(returned.__resolveKey__(argsVals[0]));
                        }
                        if (returned && returned.__isVbaCollection__) {
                            return (returned as VbaCollection).item(argsVals[0]);
                        }
                        if (Array.isArray(returned)) {
                            const dims = (returned as any).__vbaDimensions__ as { lower: number, upper: number }[] | undefined;
                            if (dims && argsVals.length !== dims.length) {
                                this.throwVbaError(VbaErrorCode.SUBSCRIPT_OUT_OF_RANGE, 'Subscript out of range');
                            }
                            let current: any = returned;
                            for (let i = 0; i < argsVals.length; i++) {
                                if (!Array.isArray(current)) {
                                    this.throwVbaError(VbaErrorCode.SUBSCRIPT_OUT_OF_RANGE, 'Subscript out of range');
                                }
                                const index = Number(argsVals[i]);
                                this.validateArrayIndex(dims, i, index);
                                current = current[index];
                            }
                            return current;
                        }
                        if (returned && returned.__vbaClass__) {
                            const retClassDef = returned.__classDef__ as ClassDeclaration;
                            const itemProp = retClassDef.procedures.find(
                                p2 => p2.isProperty && p2.propertyType === 'get' && p2.name.name.toLowerCase() === 'item'
                            );
                            if (itemProp) return this.callClassMethod(returned, itemProp, argsVals);
                        }
                        // fallthrough: let checkNoGapOnRequiredParam produce Error 450
                    }
                    this.checkNoGapOnRequiredParam(proc.parameters, expr.args);
                    return this.callClassMethodWithExpressions(obj, proc, expr.args);
                }
                // Implements interface dispatch: obj.Speak -> obj.IAnimal_Speak
                const ifaceProc = this.findInterfaceDispatch(obj, methodNameOriginal);
                if (ifaceProc) {
                    this.checkNoGapOnRequiredParam(ifaceProc.parameters, expr.args);
                    return this.callClassMethodWithExpressions(obj, ifaceProc, expr.args);
                }
                // プロシージャが見つからない場合、配列フィールドへの外部インデックスアクセスを試みる
                // 例: obj.Items(0)
                {
                    const instanceEnvForRead = obj.__instanceEnv__ as Environment;
                    const fieldArrRead = instanceEnvForRead.get(methodNameLower);
                    if (Array.isArray(fieldArrRead) && expr.args.length > 0) {
                        const dims = (fieldArrRead as any).__vbaDimensions__ as { lower: number, upper: number }[] | undefined;
                        const indexes = this.evaluateIndexExpressions(expr.args);
                        return this.readArrayAtIndexes(fieldArrRead, indexes, dims, false, false);
                    }
                }
                this.throwVbaError(VbaErrorCode.OBJECT_DOESNT_SUPPORT_PROPERTY, `Object doesn't support this property or method: '${methodNameOriginal}'`);
            }

            // UDT array member subscript access: obj.Items(0) where obj is a UDT instance
            if (obj && obj.__vbaTypeName__) {
                const memberKey = this.resolveObjectMemberKey(obj, methodNameLower);
                const fieldArr = memberKey !== undefined ? obj[memberKey] : undefined;
                if (Array.isArray(fieldArr) && expr.args.length > 0) {
                    const dims = (fieldArr as any).__vbaDimensions__ as { lower: number, upper: number }[] | undefined;
                    let current: any = fieldArr;
                    for (let i = 0; i < expr.args.length - 1; i++) {
                        const d = this.evaluateArrayIndex(expr.args[i], dims, i);
                        current = current[d];
                        if (!Array.isArray(current)) this.throwVbaError(VbaErrorCode.SUBSCRIPT_OUT_OF_RANGE, 'Subscript out of range');
                    }
                    const lastIdx = this.evaluateArrayIndex(expr.args[expr.args.length - 1], dims, expr.args.length - 1);
                    return current[lastIdx];
                }
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
                    return this.invokeBuiltin(targetMethod, this.resolveCallArgs(targetMethod, expr.args, methodNameOriginal), obj);
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
                const idx = this.evaluateArrayIndex(expr.args[i], (target as any).__vbaDimensions__, i);
                current = current[idx];
            }
            return current === undefined ? vbaEmpty : current;
                } else if (target && target.__isVbaDict__) {
            if (expr.args.length === 0) this.throwVbaError(VbaErrorCode.ARGUMENT_NOT_OPTIONAL, 'Argument not optional');
            const key = target.__resolveKey__(this.evaluateExpression(expr.args[0]));
            return target.__map__.get(key);
        } else if (typeof target === 'function') {
            return this.invokeBuiltin(target, this.resolveCallArgs(target, expr.args, expr.callee.type));
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

    private evaluateImplicitWithDictionaryAccessExpression(expr: ImplicitWithDictionaryAccessExpression): any {
        if (this.withObjectStack.length === 0) {
            this.throwVbaError(VbaErrorCode.OBJECT_VARIABLE_NOT_SET, 'Object variable or With block variable not set');
        }
        const obj = this.withObjectStack[this.withObjectStack.length - 1];
        if (obj && obj.__isVbaDict__) return obj.__map__.get(expr.property.name);
        this.throwVbaError(VbaErrorCode.OBJECT_DOESNT_SUPPORT_PROPERTY, "Object doesn't support this property or method");
    }

    private evaluateTypeOfIsExpression(expr: TypeOfIsExpression): any {
        const raw = this.evaluateExpression(expr.expression);
        // `Dim x As New ClassName` creates an auto-instance placeholder until first member access.
        // Force materialization so TypeOf can inspect the real instance.
        const obj = this.resolveAutoInstance(expr.expression as any, raw);
        const typeName = expr.typeName.toLowerCase();

        if (!isVbaBoundObject(obj)) return vbaFalse;

        // Check for built-in types
        if (typeName === 'object') return vbaTrue; // Everything that reaches here is an object
        if (typeName === 'dictionary' && obj.__isVbaDict__) return vbaTrue;
        if (typeName === 'collection' && obj.__isVbaCollection__) return vbaTrue;

        // COM ProgID 照合 (e.g. "Scripting.Dictionary")
        if (obj.__progId__ && obj.__progId__.toLowerCase() === typeName) return vbaTrue;

        // User defined types or classes (if we store metadata)
        if (obj.__vbaTypeName__ && obj.__vbaTypeName__.toLowerCase() === typeName) return vbaTrue;

        // Bug CF: `TypeOf obj Is InterfaceName` when obj's class Implements the interface.
        // The object's __vbaTypeName__ is the concrete class (e.g. "Dog"), not the interface
        // ("Animal"). Check the class body for ImplementsDirective nodes.
        if (obj.__vbaClass__ && obj.__classDef__) {
            const classDef = obj.__classDef__ as ClassDeclaration;
            const implementsIt = classDef.body.some(
                stmt => stmt.type === 'ImplementsDirective' &&
                    (stmt as any).interfaceName.toLowerCase() === typeName
            );
            if (implementsIt) return vbaTrue;
        }

        return vbaFalse;
    }

    private evaluateUnaryExpression(expr: UnaryExpression): any {
        // A signed minimum Integer/Long literal is lexed as a positive
        // suffixed literal followed by unary minus.  Its positive magnitude
        // is outside the declared type, but the signed value is representable
        // (e.g. -32768% and -2147483648&).  Handle this boundary before the
        // operand is coerced so the suffix check sees the final signed value.
        if (expr.operator === '-' && expr.argument.type === 'NumberLiteral') {
            const literal = expr.argument as NumberLiteral;
            if (literal.typeSuffix === '%' && literal.value === 32768) return -32768;
            if (literal.typeSuffix === '&' && literal.value === 2147483648) return -2147483648;
        }
        let argument = this.evaluateExpression(expr.argument);
        const op = expr.operator.toLowerCase();

        // VBA Null / Empty の伝播ルール
        // - 算術系の単項 (-, +): Null は Null、Empty は 0 として扱う
        // - Not: Null は Null
        if (argument === vbaNull) return vbaNull;
        if (argument instanceof VbaErrorValue) {
            this.throwVbaError(VbaErrorCode.TYPE_MISMATCH, 'Type mismatch');
        }
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
                if (argument instanceof VbaCurrency) return new VbaCurrency(-argument.internal);
                if (argument instanceof VbaDecimal) return new VbaDecimal(-argument.mantissa, argument.scale);
                return -argument;
            case '+':
                // JavaScript の単項 + は BigInt を受け付けない。LongLong は
                // 値を変更しない単項演算としてそのまま保持する。
                if (typeof argument === 'bigint') return argument;
                return +argument;
            default:
                throw new Error(`Execution error: Unknown unary operator ${expr.operator}`);
        }
    }

    /**
     * Decimal の BigInt 固定小数点演算（96-bit mantissa + scale 0-28）。
     */
    private evaluateDecimalOp(op: string, a: any, b: any, toVbaNumber: (v: any) => number): any {
        const parseVal = (v: any): { m: bigint; scale: number } => {
            if (v instanceof VbaDecimal) return { m: v.mantissa, scale: v.scale };
            if (v instanceof VbaCurrency) return { m: v.internal, scale: 4 };
            if (typeof v === 'bigint') return { m: v, scale: 0 };
            if (typeof v === 'string') {
                const trimmed = v.trim().replace(/^\+/, '');
                const d = VbaDecimal.fromString(trimmed);
                return { m: d.mantissa, scale: d.scale };
            }
            const d = VbaDecimal.fromNumber(toVbaNumber(v));
            return { m: d.mantissa, scale: d.scale };
        };

        const { m: am, scale: as_ } = parseVal(a);
        const { m: bm, scale: bs_ } = parseVal(b);

        // スケール整列ヘルパー
        const align = () => {
            const maxScale = Math.max(as_, bs_);
            const aam = as_ < maxScale ? am * (10n ** BigInt(maxScale - as_)) : am;
            const bbm = bs_ < maxScale ? bm * (10n ** BigInt(maxScale - bs_)) : bm;
            return { aam, bbm, maxScale };
        };

        const normalizeDecimal = (m: bigint, scale: number): VbaDecimal => {
            let mAdj = m;
            let sAdj = scale;
            const maxM = VbaDecimal.MAX_MANTISSA;
            while (sAdj > VbaDecimal.MAX_SCALE) {
                mAdj = bankersDivide(mAdj, 10n);
                sAdj--;
            }
            while ((mAdj < 0n ? -mAdj : mAdj) > maxM && sAdj > 0) {
                mAdj = bankersDivide(mAdj, 10n);
                sAdj--;
            }
            if ((mAdj < 0n ? -mAdj : mAdj) > maxM) {
                this.throwVbaError(VbaErrorCode.OVERFLOW, 'Overflow');
            }
            return new VbaDecimal(mAdj, sAdj);
        };

        const toLongOperand = (m: bigint, scale: number): bigint => {
            const value = bankersDivide(m, 10n ** BigInt(scale));
            if (value < -2147483648n || value > 2147483647n) {
                this.throwVbaError(VbaErrorCode.OVERFLOW, 'Overflow');
            }
            return value;
        };

        switch (op) {
            case '+': { const { aam, bbm, maxScale } = align(); return normalizeDecimal(aam + bbm, maxScale); }
            case '-': { const { aam, bbm, maxScale } = align(); return normalizeDecimal(aam - bbm, maxScale); }
            case '*': return normalizeDecimal(am * bm, as_ + bs_);
            case '/': {
                if (bm === 0n) this.throwVbaError(VbaErrorCode.DIVISION_BY_ZERO, 'Division by zero');
                // result_mantissa = am × 10^(28 - as + bs) / bm, scale = 28
                const extExp = 28 - as_ + bs_;
                const extNum = am * (10n ** BigInt(extExp));
                return normalizeDecimal(bankersDivide(extNum, bm), 28);
            }
            case '\\': {
                const la = toLongOperand(am, as_);
                const lb = toLongOperand(bm, bs_);
                if (lb === 0n) this.throwVbaError(VbaErrorCode.DIVISION_BY_ZERO, 'Division by zero');
                return Number(la / lb);
            }
            case 'mod': {
                const la = toLongOperand(am, as_);
                const lb = toLongOperand(bm, bs_);
                if (lb === 0n) this.throwVbaError(VbaErrorCode.DIVISION_BY_ZERO, 'Division by zero');
                return Number(la % lb);
            }
            // 比較: スケール整列してから大小比較
            case '=':  { const { aam, bbm } = align(); return aam === bbm ? vbaTrue : vbaFalse; }
            case '<>': { const { aam, bbm } = align(); return aam !== bbm ? vbaTrue : vbaFalse; }
            case '<':  { const { aam, bbm } = align(); return aam < bbm ? vbaTrue : vbaFalse; }
            case '>':  { const { aam, bbm } = align(); return aam > bbm ? vbaTrue : vbaFalse; }
            case '<=': { const { aam, bbm } = align(); return aam <= bbm ? vbaTrue : vbaFalse; }
            case '>=': { const { aam, bbm } = align(); return aam >= bbm ? vbaTrue : vbaFalse; }
        }
        return undefined;
    }

    /**
     * Currency の BigInt 固定小数点演算。
     * Currency + {Byte/Integer/Long/Boolean/bigint/Currency} → Currency
     * Currency + Double(non-integer float) → Double
     * Currency / any → Double
     */
    private evaluateCurrencyOp(op: string, a: any, b: any, toVbaNumber: (v: any) => number): any {
        const toCurrencyInternal = (v: any): bigint | null => {
            if (v instanceof VbaCurrency) return v.internal;
            if (v instanceof VbaBoolean) return BigInt(v.value) * 10000n;
            if (typeof v === 'bigint') return v * 10000n;
            if (typeof v === 'number') {
                if (Number.isInteger(v)) return BigInt(v) * 10000n;
                return null; // non-integer float → Double path
            }
            if (typeof v === 'string') {
                const trimmed = v.trim().replace(/^\+/, '');
                if (/^-?(\d+\.?\d*|\.\d+)$/.test(trimmed)) {
                    return parseFixedPointString(trimmed, 4);
                }
                return null;
            }
            return null;
        };

        const ai = toCurrencyInternal(a);
        const bi = toCurrencyInternal(b);

        // If either side is a non-integer float (Single/Double), fall through to Double arithmetic
        if (ai === null || bi === null) {
            const la = a instanceof VbaCurrency ? Number(a.internal) / 10000 : toVbaNumber(a);
            const lb = b instanceof VbaCurrency ? Number(b.internal) / 10000 : toVbaNumber(b);
            switch (op) {
                case '+': return la + lb;
                case '-': return la - lb;
                case '*': return la * lb;
                case '/':
                    // 0/0 は Error 6 (Overflow)、x/0 (x≠0) は Error 11（実 VBA 差分で裁定）
                    if (lb === 0 && la === 0) this.throwVbaError(VbaErrorCode.OVERFLOW, 'Overflow');
                    if (lb === 0) this.throwVbaError(VbaErrorCode.DIVISION_BY_ZERO, 'Division by zero');
                    return la / lb;
                case '\\': {
                    const li = _vbaRound(la, 0), ri = _vbaRound(lb, 0);
                    if (ri === 0) this.throwVbaError(VbaErrorCode.DIVISION_BY_ZERO, 'Division by zero');
                    return Math.trunc(li / ri);
                }
                case 'mod': {
                    const lm = _vbaRound(la, 0), rm = _vbaRound(lb, 0);
                    if (rm === 0) this.throwVbaError(VbaErrorCode.DIVISION_BY_ZERO, 'Division by zero');
                    return lm % rm;
                }
                case '=':  return la === lb ? vbaTrue : vbaFalse;
                case '<>': return la !== lb ? vbaTrue : vbaFalse;
                case '<':  return la < lb ? vbaTrue : vbaFalse;
                case '>':  return la > lb ? vbaTrue : vbaFalse;
                case '<=': return la <= lb ? vbaTrue : vbaFalse;
                case '>=': return la >= lb ? vbaTrue : vbaFalse;
            }
        }

        switch (op) {
            case '+': return new VbaCurrency(ai! + bi!);
            case '-': return new VbaCurrency(ai! - bi!);
            case '*':
                // (a×10^4)(b×10^4) = real×10^8 → divide by 10^4 with banker's rounding
                return new VbaCurrency(bankersDivide(ai! * bi!, 10000n));
            case '/': {
                // Currency / Currency always yields Double per VBA spec
                if (bi === 0n) this.throwVbaError(VbaErrorCode.DIVISION_BY_ZERO, 'Division by zero');
                return Number(ai!) / Number(bi!);
            }
            case '\\': {
                const li = bankersDivide(ai!, 10000n);
                const ri = bankersDivide(bi!, 10000n);
                if (ri === 0n) this.throwVbaError(VbaErrorCode.DIVISION_BY_ZERO, 'Division by zero');
                if (li < -2147483648n || li > 2147483647n || ri < -2147483648n || ri > 2147483647n) {
                    this.throwVbaError(VbaErrorCode.OVERFLOW, 'Overflow');
                }
                return Number(li / ri);
            }
            case 'mod': {
                const lm = bankersDivide(ai!, 10000n);
                const rm = bankersDivide(bi!, 10000n);
                if (rm === 0n) this.throwVbaError(VbaErrorCode.DIVISION_BY_ZERO, 'Division by zero');
                if (lm < -2147483648n || lm > 2147483647n || rm < -2147483648n || rm > 2147483647n) {
                    this.throwVbaError(VbaErrorCode.OVERFLOW, 'Overflow');
                }
                return Number(lm % rm);
            }
            case '=':  return ai! === bi! ? vbaTrue : vbaFalse;
            case '<>': return ai! !== bi! ? vbaTrue : vbaFalse;
            case '<':  return ai! < bi! ? vbaTrue : vbaFalse;
            case '>':  return ai! > bi! ? vbaTrue : vbaFalse;
            case '<=': return ai! <= bi! ? vbaTrue : vbaFalse;
            case '>=': return ai! >= bi! ? vbaTrue : vbaFalse;
        }
        return undefined;
    }

    /**
     * Auto-Instance placeholder を実際のインスタンスに解決する。
     * 解決後は `Identifier` の場合 env に書き戻して、次回のアクセスではインスタンスが
     * 直接使えるようにする（毎回 instantiate するコストを避ける）。
     * Placeholder でない値はそのまま返す。
     */
    private deepCopyByValValue(value: any): any {
        if (Array.isArray(value)) {
            const copy: any = value.map(item => this.deepCopyByValValue(item));
            for (const key of Object.keys(value)) {
                if (!/^\d+$/.test(key)) copy[key] = (value as any)[key];
            }
            return copy;
        }
        if (value && typeof value === 'object' && value.__vbaUdt__ === true) {
            return this.deepCopyUdtValue(value);
        }
        return value;
    }

    /** UDT 値型の ByVal コピーを作成する。ネストした UDT メンバーも再帰コピーする。 */
    private deepCopyUdtValue(obj: any): any {
        if (!obj || typeof obj !== 'object' || obj.__vbaClass__) return obj;
        const copy: any = {};
        for (const key of Object.keys(obj)) {
            const val = obj[key];
            if (Array.isArray(val)) {
                copy[key] = this.deepCopyByValValue(val);
            } else if (val && typeof val === 'object' && val.__vbaUdt__ === true) {
                copy[key] = this.deepCopyUdtValue(val);
            } else {
                copy[key] = val;
            }
        }
        return copy;
    }

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

            // No-arg Sub/Function access without parens: call it (Sub or Function).
            const method = classDef.procedures.find(
                p => !p.isProperty && p.name.name.toLowerCase() === propName && p.parameters.length === 0
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
            if (instanceEnv.hasOwnVariable(propName)) return instanceEnv.get(propName);
            this.throwVbaError(VbaErrorCode.OBJECT_DOESNT_SUPPORT_PROPERTY, `Object doesn't support this property or method: '${propName}'`);
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
                // Host/COM properties may expose SAFEARRAY-like values
                // without VBA ProcedureDeclaration metadata. Treat the
                // evaluated value as a getter result for this assignment only.
                if (Array.isArray(val) && (obj as any).__progId__) {
                    (val as any).__vbaArrayReturn__ = true;
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
        const result = this.evaluateBinaryExpressionInner(expr);
        // Integer/Long 型どうしの算術はその型のまま演算され、範囲超過は Error 6 になる
        // （実 VBA 差分で裁定: (2)+(32767) は Long へ昇格せず Overflow）。
        // サブタイプ解決は範囲超過の疑いがあるときだけ行う
        if (typeof result === 'number' && Number.isFinite(result)
                && (result > 32767 || result < -32768)) {
            const op = expr.operator.toLowerCase();
            if (op === '+' || op === '-' || op === '*') {
                const st = this.resolveBinaryExprNumericSubtype(expr);
                if (st === 'Integer' || st === 'Byte') {
                    this.throwVbaError(VbaErrorCode.OVERFLOW, 'Overflow');
                }
                if (st === 'Long' && (result > 2147483647 || result < -2147483648)) {
                    this.throwVbaError(VbaErrorCode.OVERFLOW, 'Overflow');
                }
            }
        }
        return result;
    }

    private evaluateBinaryExpressionInner(expr: BinaryExpression): any {
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

        // CVErr values do not participate in implicit string, numeric, or
        // logical coercion. They may only be compared for equality or passed
        // explicitly to a conversion such as CStr.
        if ((op === '&' || logicalOps.has(op)) &&
            (leftVal instanceof VbaErrorValue || rightVal instanceof VbaErrorValue)) {
            this.throwVbaError(VbaErrorCode.TYPE_MISMATCH, 'Type mismatch');
        }

        if (op === '&') {
            // 文字列連結: Empty は "" 扱い。Null は両辺とも Null のときのみ Null を返し、
            // 片方のみ Null のときは "" として連結する（VBA 仕様）。
            if (leftVal === vbaNull && rightVal === vbaNull) return vbaNull;
            if (leftVal === vbaNull || leftVal === vbaEmpty) leftVal = '';
            if (rightVal === vbaNull || rightVal === vbaEmpty) rightVal = '';
        } else if (arithmeticOps.has(op) || comparisonOps.has(op) || logicalOps.has(op)) {
            // Null 伝播より先に、非 Null 側が数値変換不能な文字列なら Type Mismatch。
            // ただし "+" だけは例外（Null 伝播を優先）。実 VBA 差分で裁定:
            // ("abc") - (Null) は Error 13 だが ("abc") + (Null) は Null のまま
            if (op !== '+' && arithmeticOps.has(op) && (leftVal === vbaNull || rightVal === vbaNull)) {
                const nonNullSide = leftVal === vbaNull ? rightVal : leftVal;
                if (typeof nonNullSide === 'string') this.toVbaNumber(nonNullSide);
            }
            // Null 伝播: どちらかが Null なら Null（ただし And/Or/Imp には吸収元例外あり）
            if (leftVal === vbaNull || rightVal === vbaNull) {
                // VBA 三値論理の特例:
                //   False And Null = False  (False は And の吸収元)
                //   True  Or  Null = True   (True は Or の吸収元)
                //   False Imp Null = True   (左が False なら Imp は常に True)
                //   Null  Imp True = True   (右が True なら Imp は常に True)
                if (logicalOps.has(op)) {
                    const toNum = (v: any): number => v instanceof VbaBoolean ? v.value : (typeof v === 'number' ? v : Number(v));
                    const nonNull = leftVal === vbaNull ? rightVal : leftVal;
                    if (op === 'and' && toNum(nonNull) === 0) return vbaFalse;
                    if (op === 'or'  && toNum(nonNull) !== 0) return vbaTrue;
                    if (op === 'imp') {
                        if (leftVal !== vbaNull  && toNum(leftVal)  === 0) return vbaTrue;
                        if (rightVal !== vbaNull && toNum(rightVal) !== 0) return vbaTrue;
                    }
                }
                return vbaNull;
            }
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
        const toExactBigInt = (v: any): bigint => {
            if (typeof v === 'bigint') return v;
            if (typeof v === 'string') {
                const trimmed = v.trim();
                if (/^[+-]?\d+$/.test(trimmed)) return BigInt(trimmed);
            }
            return BigInt(toVbaNumber(v));
        };

        // Decimal: BigInt-backed string arithmetic for precision
        if (leftVal instanceof VbaDecimal || rightVal instanceof VbaDecimal) {
            if (arithmeticOps.has(op) || comparisonOps.has(op)) {
                return this.evaluateDecimalOp(op, leftVal, rightVal, toVbaNumber);
            }
        }

        // Currency: BigInt fixed-point arithmetic (before VbaBoolean normalisation for comparisons)
        if (leftVal instanceof VbaCurrency || rightVal instanceof VbaCurrency) {
            if (arithmeticOps.has(op) || comparisonOps.has(op)) {
                return this.evaluateCurrencyOp(op, leftVal, rightVal, toVbaNumber);
            }
        }

        // LongLong の整数演算は Number に落とさず BigInt のまま評価する。
        // JavaScript Number は 2^53 を超える整数を表現できないため、通常の
        // 算術経路へ進むと LongLong の下位桁が失われる。Double 昇格を伴う
        // `/` と `^` はここでは扱わず、整数演算の +,-,*,\\,Mod のみ対象にする。
        const longLongIntegerOps = new Set(['+', '-', '*', '\\', 'mod']);
        const isExactInteger = (v: any): boolean =>
            typeof v === 'bigint' ||
            (typeof v === 'number' && Number.isSafeInteger(v)) ||
            v instanceof VbaBoolean;
        const isIntegerString = (v: any): boolean =>
            typeof v === 'string' && /^[+-]?\d+$/.test(v.trim());
        const isExactIntegerOperand = (v: any): boolean =>
            isExactInteger(v) || isIntegerString(v);
        const hasStringOperand = typeof leftVal === 'string' || typeof rightVal === 'string';
        // VBA promotes LongLong +/− String to Double (XL-016), but the
        // integer-only operators retain an integer String exactly.
        const exactMixedIntegerOp = op === '*' || op === '\\' || op === 'mod';
        if (longLongIntegerOps.has(op) &&
            (typeof leftVal === 'bigint' || typeof rightVal === 'bigint') &&
            (!hasStringOperand || exactMixedIntegerOp) &&
            isExactIntegerOperand(leftVal) && isExactIntegerOperand(rightVal)) {
            const asBigInt = (v: any): bigint => {
                if (typeof v === 'bigint') return v;
                if (v instanceof VbaBoolean) return BigInt(v.value);
                if (typeof v === 'string') return BigInt(v.trim());
                return BigInt(v);
            };
            const l = asBigInt(leftVal);
            const r = asBigInt(rightVal);
            let result: bigint;
            if ((op === '\\' || op === 'mod') && r === 0n) {
                this.throwVbaError(VbaErrorCode.DIVISION_BY_ZERO, 'Division by zero');
            }
            switch (op) {
                case '+': result = l + r; break;
                case '-': result = l - r; break;
                case '*': result = l * r; break;
                case '\\': result = l / r; break;
                case 'mod': result = l % r; break;
                default: throw new Error(`Execution error: Unknown integer operator ${op}`);
            }
            if (result < -9223372036854775808n || result > 9223372036854775807n) {
                this.throwVbaError(VbaErrorCode.OVERFLOW, 'Overflow');
            }
            return result;
        }

        // 比較演算子では VbaBoolean を数値として扱う（True=-1, False=0）
        if (comparisonOps.has(op)) {
            // 片方が Boolean でもう片方が文字列なら、文字列側を CBool 変換して比較する
            // （数値変換ではない。実 VBA 差分で裁定: "7" = True は True。CBool("7")=True のため）
            const leftWasBoolean = leftVal instanceof VbaBoolean;
            const rightWasBoolean = rightVal instanceof VbaBoolean;
            if (leftWasBoolean) leftVal = leftVal.value;
            if (rightWasBoolean) rightVal = rightVal.value;
            if (leftWasBoolean && typeof rightVal === 'string') {
                rightVal = vbaToBoolean(rightVal).value;
            } else if (rightWasBoolean && typeof leftVal === 'string') {
                leftVal = vbaToBoolean(leftVal).value;
            } else {
                // 片方が文字列でもう片方が数値/日付なら、文字列側を数値変換して比較する
                // （変換失敗は Type Mismatch。実 VBA 差分で裁定: "abc" = #date# は Error 13）。
                // 文字列同士の比較（Text 比較モード含む）はここでは変換しない。
                const isNumericLike = (v: any) => typeof v === 'number' || typeof v === 'bigint' || v instanceof VbaDate;
                if (typeof leftVal === 'string' && isNumericLike(rightVal)) {
                    leftVal = typeof rightVal === 'bigint' ? toExactBigInt(leftVal) : toVbaNumber(leftVal);
                    if (rightVal instanceof VbaDate) rightVal = rightVal.value;
                } else if (isNumericLike(leftVal) && typeof rightVal === 'string') {
                    rightVal = typeof leftVal === 'bigint' ? toExactBigInt(rightVal) : toVbaNumber(rightVal);
                    if (leftVal instanceof VbaDate) leftVal = leftVal.value;
                }
            }
        }

        // \ と Mod: 左が Boolean で右が文字列のときだけ、右を CBool 変換し結果も
        // Boolean 型で返す（実 VBA 差分で裁定した非対称規則。左右逆・両方 Boolean
        // literal・左が文字列右が Boolean のいずれも該当せず通常の整数演算になる）:
        //   True \ "7"  → "7" を CBool すると True(-1)。True(-1)\True(-1)=1 → Boolean True
        //   "7" \ True  → 通常の数値変換。7\(-1) = -7 → Long -7
        //   True \ True → 両方すでに Boolean（変換不要）→ 通常の整数昇格 → Integer 1
        let resultAsBoolean = false;
        if ((op === '\\' || op === 'mod') && leftVal instanceof VbaBoolean && typeof rightVal === 'string') {
            rightVal = vbaToBoolean(rightVal);
            resultAsBoolean = true;
        }

        const isPlusConcatenation = op === '+' && typeof leftVal === 'string' && typeof rightVal === 'string';
        const numericArithOps = new Set(['-', '*', '/', '\\', 'mod', '^']);
        if (numericArithOps.has(op) || (op === '+' && !isPlusConcatenation)) {
            // "+" に限り、Date とペアになる "H.N" 形式の文字列は「H 時 N 分」の時刻として
            // 解釈される（CDate と同じ規則。実 VBA 差分で裁定: #date# + "3.5" は 03:05:00
            // を加算する）。"-"/"*"/"/"/"\"/"Mod" では同じ文字列でも通常どおりシリアル値
            // 3.5 として解釈される（+ 演算子だけの非対称な特殊ルール）
            if (op === '+') {
                if (typeof leftVal === 'string' && rightVal instanceof VbaDate) {
                    const frac = tryParseTimeFractionString(leftVal);
                    if (frac !== undefined) leftVal = frac;
                }
                if (typeof rightVal === 'string' && leftVal instanceof VbaDate) {
                    const frac = tryParseTimeFractionString(rightVal);
                    if (frac !== undefined) rightVal = frac;
                }
            }
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
                if ((leftVal instanceof VbaDate || rightVal instanceof VbaDate)
                        && (sum < -657434 || sum >= 2958466)) {
                    this.throwVbaError(VbaErrorCode.OVERFLOW, 'Overflow');
                }
                return (leftVal instanceof VbaDate || rightVal instanceof VbaDate) ? new VbaDate(sum) : sum;
            }
            case '&': return this.toDisplayString(leftVal) + this.toDisplayString(rightVal);
            case '-': {
                const l = leftVal instanceof VbaDate ? leftVal.value : leftVal;
                const r = rightVal instanceof VbaDate ? rightVal.value : rightVal;
                const diff = l - r;
                if (leftVal instanceof VbaDate && rightVal instanceof VbaDate) return diff; // Date - Date = Number
                // 数値 - Date、Date - 数値 のどちらも Date 型結果になる（実 VBA 差分で裁定）
                if ((leftVal instanceof VbaDate || rightVal instanceof VbaDate)
                        && (diff < -657434 || diff >= 2958466)) {
                    this.throwVbaError(VbaErrorCode.OVERFLOW, 'Overflow');
                }
                return (leftVal instanceof VbaDate || rightVal instanceof VbaDate) ? new VbaDate(diff) : diff;
            }
            case '*': {
                const product = leftVal * rightVal;
                if (!Number.isFinite(product)) this.throwVbaError(VbaErrorCode.OVERFLOW, 'Overflow');
                return product;
            }
            case '/':
                // 0/0 は Error 6 (Overflow)、x/0 (x≠0) は Error 11（実 VBA 差分で裁定）
                if (rightVal === 0 && leftVal === 0) this.throwVbaError(VbaErrorCode.OVERFLOW, 'Overflow');
                if (rightVal === 0) this.throwVbaError(VbaErrorCode.DIVISION_BY_ZERO, 'Division by zero');
                return leftVal / rightVal;
            case '\\': {
                const li = _vbaRound(leftVal, 0), ri = _vbaRound(rightVal, 0);
                if (ri === 0) this.throwVbaError(VbaErrorCode.DIVISION_BY_ZERO, 'Division by zero');
                const q = Math.trunc(li / ri);
                return resultAsBoolean ? VbaBoolean.from(q) : q;
            }
            case 'mod': {
                const lm = _vbaRound(leftVal, 0), rm = _vbaRound(rightVal, 0);
                if (rm === 0) this.throwVbaError(VbaErrorCode.DIVISION_BY_ZERO, 'Division by zero');
                const r = lm % rm;
                return resultAsBoolean ? VbaBoolean.from(r) : r;
            }
            case '^': {
                const powResult = Math.pow(leftVal, rightVal);
                if (isNaN(powResult)) this.throwVbaError(VbaErrorCode.INVALID_PROCEDURE_CALL, 'Invalid procedure call or argument');
                if (!Number.isFinite(powResult)) this.throwVbaError(VbaErrorCode.OVERFLOW, 'Overflow');
                return powResult;
            }
            case '=':
                if (leftVal instanceof VbaErrorValue && rightVal instanceof VbaErrorValue) {
                    return leftVal.code === rightVal.code ? vbaTrue : vbaFalse;
                }
                if (leftVal instanceof VbaDate && rightVal instanceof VbaDate) {
                    return leftVal.value === rightVal.value ? vbaTrue : vbaFalse;
                }
                if (typeof leftVal === 'string' && typeof rightVal === 'string' && this.getComparisonMode() === 'Text') {
                    return leftVal.toLowerCase() === rightVal.toLowerCase() ? vbaTrue : vbaFalse;
                }
                // Mixed string/number: coerce string to number (VBA auto-coercion)
                if (typeof leftVal === 'string' && typeof rightVal === 'number') {
                    return toVbaNumber(leftVal) === rightVal ? vbaTrue : vbaFalse;
                }
                if (typeof leftVal === 'number' && typeof rightVal === 'string') {
                    return leftVal === toVbaNumber(rightVal) ? vbaTrue : vbaFalse;
                }
                return leftVal === rightVal ? vbaTrue : vbaFalse;
            case '<>':
                if (leftVal instanceof VbaErrorValue && rightVal instanceof VbaErrorValue) {
                    return leftVal.code !== rightVal.code ? vbaTrue : vbaFalse;
                }
                if (leftVal instanceof VbaDate && rightVal instanceof VbaDate) {
                    return leftVal.value !== rightVal.value ? vbaTrue : vbaFalse;
                }
                if (typeof leftVal === 'string' && typeof rightVal === 'string' && this.getComparisonMode() === 'Text') {
                    return leftVal.toLowerCase() !== rightVal.toLowerCase() ? vbaTrue : vbaFalse;
                }
                // Mixed string/number: coerce string to number (VBA auto-coercion)
                if (typeof leftVal === 'string' && typeof rightVal === 'number') {
                    return toVbaNumber(leftVal) !== rightVal ? vbaTrue : vbaFalse;
                }
                if (typeof leftVal === 'number' && typeof rightVal === 'string') {
                    return leftVal !== toVbaNumber(rightVal) ? vbaTrue : vbaFalse;
                }
                return leftVal !== rightVal ? vbaTrue : vbaFalse;
            case '<':
                if (typeof leftVal === 'string' && typeof rightVal === 'string' && this.getComparisonMode() === 'Text') {
                    return leftVal.toLowerCase() < rightVal.toLowerCase() ? vbaTrue : vbaFalse;
                }
                return leftVal < rightVal ? vbaTrue : vbaFalse;
            case '>':
                if (typeof leftVal === 'string' && typeof rightVal === 'string' && this.getComparisonMode() === 'Text') {
                    return leftVal.toLowerCase() > rightVal.toLowerCase() ? vbaTrue : vbaFalse;
                }
                return leftVal > rightVal ? vbaTrue : vbaFalse;
            case '<=':
                if (typeof leftVal === 'string' && typeof rightVal === 'string' && this.getComparisonMode() === 'Text') {
                    return leftVal.toLowerCase() <= rightVal.toLowerCase() ? vbaTrue : vbaFalse;
                }
                return leftVal <= rightVal ? vbaTrue : vbaFalse;
            case '>=':
                if (typeof leftVal === 'string' && typeof rightVal === 'string' && this.getComparisonMode() === 'Text') {
                    return leftVal.toLowerCase() >= rightVal.toLowerCase() ? vbaTrue : vbaFalse;
                }
                return leftVal >= rightVal ? vbaTrue : vbaFalse;
            case 'is': return leftVal === rightVal ? vbaTrue : vbaFalse;
            case 'like':
                if (leftVal === vbaNull || rightVal === vbaNull) return vbaNull;
                return this.evaluateLike(leftVal, rightVal) ? vbaTrue : vbaFalse;
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

        // VBA class instance: same logic as evaluateMemberExpression
        if (obj && obj.__vbaClass__) {
            const classDef = obj.__classDef__ as ClassDeclaration;
            const instanceEnv = obj.__instanceEnv__ as Environment;
            const getter = classDef.procedures.find(
                p => p.isProperty && p.propertyType === 'get' && p.name.name.toLowerCase() === propName
            );
            if (getter) return this.callClassMethod(obj, getter, []);
            const method = classDef.procedures.find(
                p => !p.isProperty && p.name.name.toLowerCase() === propName && p.parameters.length === 0 && p.isFunction
            );
            if (method) return this.callClassMethod(obj, method, []);
            const ifaceProc = this.findInterfaceDispatch(obj, expr.property.name);
            if (ifaceProc) return this.callClassMethod(obj, ifaceProc, []);
            if (instanceEnv.hasOwnVariable(propName)) return instanceEnv.get(propName);
            this.throwVbaError(VbaErrorCode.OBJECT_DOESNT_SUPPORT_PROPERTY, `Object doesn't support this property or method: '${propName}'`);
        }

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
                let firstClassChar = true;
                while (i < patternStr.length && (firstClassChar || patternStr[i] !== ']')) {
                    const charInList = patternStr[i];
                    // Escape regex special chars inside [] if they are not part of range
                    if ('\\^$[]{}|()+.'.includes(charInList) && charInList !== '-') {
                        regexStr += '\\' + charInList;
                    } else {
                        regexStr += charInList;
                    }
                    i++;
                    firstClassChar = false;
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
            const flags = this.getComparisonMode() === 'Text' ? 'i' : '';
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

}
