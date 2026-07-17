import {
    VbaBoolean, VbaDate, VbaDecimal, VbaCurrency, VbaErrorValue,
    vbaEmpty, vbaNull, vbaNothing, vbaMissing, vbaTrue, vbaFalse,
    toVbaDate, fromVbaDate, parseVbaDate,
    parseFixedPointString,
} from './vba-types';
import { VbaErrorCode } from './vba-errors';
import { vbaToBoolean, vbaToString, vbaRound } from './coerce';
// VbaErrorCode is imported as a value-namespace for use in function bodies (VbaErrorCode.OVERFLOW etc.)
import type { ProcedureDeclaration } from './parser';
import { formatDate, formatNumber, formatString } from './format';

// ---------------------------------------------------------------------------
// Shared type definitions (also re-exported from evaluator.ts)
// ---------------------------------------------------------------------------

/**
 * VBA パラメーター仕様。registerBuiltin に渡し、名前付き引数解決・引数数チェックに使う。
 *
 * - 「必須引数のあとに Optional 引数が続く」通常の形は `BuiltinParamSpec[]` をそのまま使う。
 * - `InStr` のように先頭 Optional が個数で有無する不規則な組み込み関数は `BuiltinOverload` を使う。
 */
export interface BuiltinParamSpec {
    name: string;
    optional?: boolean;
    isParamArray?: boolean;
}

export interface BuiltinOverload {
    params: BuiltinParamSpec[];
}

// ---------------------------------------------------------------------------
// Context interface — passed from Evaluator to each register function
// ---------------------------------------------------------------------------

export interface StdlibCtx {
    reg(name: string, fn: any, params: BuiltinParamSpec[], variants?: string[]): void;
    regOvl(name: string, fn: any, overloads: BuiltinOverload[], variants?: string[]): void;
    envSet(name: string, val: any, variants?: string[]): void;
    envSetConst(name: string, val: any): void;
    envGet(name: string): any;
    toVbaNumber(val: any): number;
    throwError(code: number, msg: string): never;
    isTrue(val: any): boolean;
    round(n: number, digits?: number): number;
    callMethod(obj: any, proc: ProcedureDeclaration, args: any[]): any;
    readonly compMode: 'Binary' | 'Text';
    readonly arrayBase: number;
    print(msg: string): void;
    errNum(): number;
    getEnv(k: any): string;
    ptrNext(): number;
    createObj(id: string): any;
    registry: Record<string, Record<string, Record<string, string>>>;
}

// ---------------------------------------------------------------------------
// Information functions — IsEmpty, IsNumeric, TypeName, CallByName, etc.
// ---------------------------------------------------------------------------

export function registerInformationFunctions(ctx: StdlibCtx): void {
    ctx.reg('isempty', (val: any) => (val === undefined || val === null || val === vbaEmpty) ? vbaTrue : vbaFalse, [{ name: 'Expression' }]);
    ctx.reg('ismissing', (val: any) => val === vbaMissing ? vbaTrue : vbaFalse, [{ name: 'ArgName' }]);
    ctx.reg('isnumeric', (val: any) => {
        if (val === vbaNull) return vbaFalse;
        if (val === vbaEmpty || val === undefined) return vbaTrue;
        if (typeof val === 'number' || typeof val === 'bigint' || val instanceof VbaDecimal || val instanceof VbaCurrency || val instanceof VbaBoolean) return vbaTrue;
        if (typeof val === 'string') {
            const s = val.trim();
            if (s === "") return vbaFalse;
            const cleaned = s.replace(/[$,]/g, '');
            return (!isNaN(Number(cleaned)) && isFinite(Number(cleaned))) ? vbaTrue : vbaFalse;
        }
        return vbaFalse;
    }, [{ name: 'Expression' }]);
    ctx.reg('isdate', (val: any) => {
        if (val instanceof VbaDate) return vbaTrue;
        if (typeof val === 'number') return isFinite(val) ? vbaTrue : vbaFalse;
        if (typeof val === 'string') {
            const d = Date.parse(val);
            return !isNaN(d) ? vbaTrue : vbaFalse;
        }
        return vbaFalse;
    }, [{ name: 'Expression' }]);
    ctx.reg('isobject', (val: any) => (val === vbaNothing || (val && typeof val === 'object' && !Array.isArray(val) && val !== vbaNull)) ? vbaTrue : vbaFalse, [{ name: 'Identifier' }]);
    ctx.reg('iserror', (val: any) => (val instanceof VbaErrorValue) ? vbaTrue : vbaFalse, [{ name: 'Expression' }]);
    ctx.reg('isnull', (val: any) => (val === vbaNull) ? vbaTrue : vbaFalse, [{ name: 'Expression' }]);
    ctx.reg('isarray', (val: any) => Array.isArray(val) ? vbaTrue : vbaFalse, [{ name: 'VarName' }]);

    ctx.reg('vartype', (val: any) => {
        if (val === vbaEmpty || val === undefined) return 0;
        if (val === vbaNull) return 1;
        if (val === vbaNothing) return 9;
        if (val instanceof VbaBoolean) return 11;
        if (val instanceof VbaDate) return 7;
        if (val === vbaMissing || val instanceof VbaErrorValue) return 10;
        if (Array.isArray(val)) return 8192 + 12;
        if (val instanceof VbaCurrency) return 6;
        if (typeof val === 'number') return 5;
        if (typeof val === 'string') return 8;
        if (val instanceof VbaDecimal) return 14;
        if (typeof val === 'bigint') return 20;
        if (val && val.__vbaTypeName__) return 36;
        if (val && (val.__vbaClass__ || val.__isVbaDict__ || val.__isVbaCollection__)) return 9;
        if (typeof val === 'object' && val !== null) return 9;
        return 12;
    }, [{ name: 'VarName' }]);

    ctx.reg('typename', (val: any) => {
        if (val === vbaEmpty || val === undefined) return 'Empty';
        if (val === vbaNull) return 'Null';
        if (val === vbaNothing) return 'Nothing';
        if (val === vbaMissing || val instanceof VbaErrorValue) return 'Error';
        if (val instanceof VbaBoolean) return 'Boolean';
        if (val instanceof VbaDate) return 'Date';
        if (val instanceof VbaCurrency) return 'Currency';
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
    ctx.reg('callbyname', (obj: any, procName: string, callType: number, ...args: any[]) => {
        if (obj === null || obj === undefined || obj === vbaNothing) {
            ctx.throwError(VbaErrorCode.OBJECT_VARIABLE_NOT_SET, 'Object variable or With block variable not set');
        }
        const name = String(procName).toLowerCase();
        if (callType === 2 /* VbGet */ || callType === 1 /* VbMethod */) {
            if (obj.__vbaClass__) {
                const classDef = obj.__classDef__ as ProcedureDeclaration & { procedures: ProcedureDeclaration[] };
                const getter = (classDef as any).procedures.find(
                    (p: any) => p.isProperty && p.propertyType === 'get' && p.name.name.toLowerCase() === name
                );
                if (getter) return ctx.callMethod(obj, getter, args);
                const method = (classDef as any).procedures.find(
                    (p: any) => !p.isProperty && p.name.name.toLowerCase() === name
                );
                if (method) return ctx.callMethod(obj, method, args);
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
        if (callType === 4 /* VbLet */ || callType === 8 /* VbSet */) {
            const propType = callType === 4 ? 'let' : 'set';
            const fallbackType = callType === 4 ? 'set' : 'let';
            if (obj.__vbaClass__) {
                const classDef = obj.__classDef__ as ProcedureDeclaration & { procedures: ProcedureDeclaration[] };
                const setter = (classDef as any).procedures.find(
                    (p: any) => p.isProperty && p.propertyType === propType && p.name.name.toLowerCase() === name
                ) ?? (classDef as any).procedures.find(
                    (p: any) => p.isProperty && p.propertyType === fallbackType && p.name.name.toLowerCase() === name
                );
                if (setter) return ctx.callMethod(obj, setter, args);
                obj.__instanceEnv__.set(name, args[0]);
                return;
            }
            if (typeof obj === 'object' && obj !== null) {
                const keys = Object.keys(obj);
                const match = keys.find(k => k.toLowerCase() === name) ?? name;
                const val = obj[match];
                if (typeof val === 'function') { val.apply(obj, args); } else { obj[match] = args[0]; }
                return;
            }
        }
        ctx.throwError(VbaErrorCode.OBJECT_DOESNT_SUPPORT_PROPERTY, `Object doesn't support this property or method: '${procName}'`);
    }, [
        { name: 'Object' },
        { name: 'ProcName' },
        { name: 'CallType' },
        { name: 'Args', isParamArray: true },
    ]);
}

// ---------------------------------------------------------------------------
// Conversion functions — CByte, CInt, CLng, CStr, Hex, Oct, Val, etc.
// ---------------------------------------------------------------------------

export function registerConversionFunctions(ctx: StdlibCtx): void {
    ctx.reg('cbyte', (val: any) => {
        if (val === vbaNull) ctx.throwError(VbaErrorCode.INVALID_USE_OF_NULL, 'Invalid use of Null');
        if (val instanceof VbaBoolean) return val.valueOf() ? 255 : 0;
        const n = ctx.round(ctx.toVbaNumber(val));
        if (n < 0 || n > 255) ctx.throwError(VbaErrorCode.OVERFLOW, "Overflow");
        return n;
    }, [{ name: 'Expression' }]);
    ctx.reg('cint', (val: any) => {
        if (val === vbaNull) ctx.throwError(VbaErrorCode.INVALID_USE_OF_NULL, 'Invalid use of Null');
        const n = ctx.round(ctx.toVbaNumber(val));
        if (n < -32768 || n > 32767) ctx.throwError(VbaErrorCode.OVERFLOW, "Overflow");
        return n;
    }, [{ name: 'Expression' }]);
    ctx.reg('clng', (val: any) => {
        if (val === vbaNull) ctx.throwError(VbaErrorCode.INVALID_USE_OF_NULL, 'Invalid use of Null');
        const n = ctx.round(ctx.toVbaNumber(val));
        if (n < -2147483648 || n > 2147483647) ctx.throwError(VbaErrorCode.OVERFLOW, "Overflow");
        return n;
    }, [{ name: 'Expression' }]);
    ctx.reg('csng', (val: any) => {
        if (val === vbaNull) ctx.throwError(VbaErrorCode.INVALID_USE_OF_NULL, 'Invalid use of Null');
        const n = ctx.toVbaNumber(val);
        const f32 = Math.fround(n);
        if (!isFinite(f32) && isFinite(n)) ctx.throwError(VbaErrorCode.OVERFLOW, "Overflow");
        return f32;
    }, [{ name: 'Expression' }]);
    ctx.reg('cdbl', (val: any) => {
        if (val === vbaNull) ctx.throwError(VbaErrorCode.INVALID_USE_OF_NULL, 'Invalid use of Null');
        return ctx.toVbaNumber(val);
    }, [{ name: 'Expression' }]);
    ctx.reg('cdate', (val: any) => {
        if (val === vbaNull) ctx.throwError(VbaErrorCode.INVALID_USE_OF_NULL, 'Invalid use of Null');
        if (val === null || val === vbaEmpty) ctx.throwError(VbaErrorCode.TYPE_MISMATCH, "Type mismatch");
        if (val instanceof VbaDate) return val;
        if (typeof val === 'string' && !/^\d+(\.\d+)?$/.test(val)) {
            return new VbaDate(toVbaDate(parseVbaDate(val)));
        }
        return new VbaDate(ctx.toVbaNumber(val));
    }, [{ name: 'Expression' }]);
    ctx.reg('cvdate', (val: any) => {
        if (val === vbaNull) return vbaNull;
        return (ctx.envGet('cdate') as Function)(val);
    }, [{ name: 'Expression' }]);
    ctx.reg('cdec', (val: any) => {
        if (val === vbaNull) ctx.throwError(VbaErrorCode.INVALID_USE_OF_NULL, 'Invalid use of Null');
        if (val instanceof VbaDecimal) return val;
        if (val instanceof VbaCurrency) return new VbaDecimal(val.internal, 4);
        if (val instanceof VbaBoolean) return new VbaDecimal(BigInt(val.value), 0);
        if (typeof val === 'bigint') return new VbaDecimal(val, 0);
        if (typeof val === 'string') return VbaDecimal.fromString(val.trim());
        return VbaDecimal.fromNumber(ctx.toVbaNumber(val));
    }, [{ name: 'Expression' }]);
    ctx.reg('ccur', (val: any) => {
        if (val === vbaNull) ctx.throwError(VbaErrorCode.INVALID_USE_OF_NULL, 'Invalid use of Null');
        if (val instanceof VbaCurrency) return val;
        if (val instanceof VbaBoolean) return new VbaCurrency(BigInt(val.value) * 10000n);
        if (typeof val === 'bigint') return new VbaCurrency(val * 10000n);
        if (typeof val === 'string') {
            const trimmed = val.trim();
            if (!/^-?(\d+\.?\d*|\.\d+)$/.test(trimmed)) {
                ctx.throwError(VbaErrorCode.TYPE_MISMATCH, 'Type mismatch');
            }
            return new VbaCurrency(parseFixedPointString(trimmed, 4));
        }
        return VbaCurrency.fromNumber(ctx.toVbaNumber(val));
    }, [{ name: 'Expression' }]);
    const clnglngFunc = (val: any) => {
        if (val === vbaNull) ctx.throwError(VbaErrorCode.INVALID_USE_OF_NULL, 'Invalid use of Null');
        if (val === null) ctx.throwError(VbaErrorCode.TYPE_MISMATCH, "Type mismatch");
        if (typeof val === 'bigint') return val;
        if (typeof val === 'string') {
            const trimmed = val.trim();
            if (/^-?\d+$/.test(trimmed)) {
                try {
                    const n = BigInt(trimmed);
                    if (n < -9223372036854775808n || n > 9223372036854775807n) ctx.throwError(VbaErrorCode.OVERFLOW, "Overflow");
                    return n;
                } catch {
                    ctx.throwError(VbaErrorCode.TYPE_MISMATCH, "Type mismatch");
                }
            }
        }
        const num = ctx.toVbaNumber(val);
        if (num < -9223372036854775808 || num > 9223372036854775807) {
            if (Math.abs(num - 9223372036854775807) < 1025 || Math.abs(num + 9223372036854775808) < 1025) {
                // near-edge: let BigInt conversion handle it
            } else {
                ctx.throwError(VbaErrorCode.OVERFLOW, "Overflow");
            }
        }
        const n = BigInt(ctx.round(num));
        if (n < -9223372036854775808n || n > 9223372036854775807n) ctx.throwError(VbaErrorCode.OVERFLOW, "Overflow");
        return n;
    };
    ctx.reg('clnglng', clnglngFunc, [{ name: 'Expression' }]);
    ctx.envSet('clngptr', ctx.envGet('clnglng')); // clnglng と同じ関数オブジェクトのため __vbaParamSpec__ も引き継ぐ
    ctx.reg('cstr', (val: any) => {
        if (val === vbaNull) return '';
        try { return vbaToString(val); } catch (e: any) {
            if (e?.type === 'VbaError') ctx.throwError(e.number, e.message);
            throw e;
        }
    }, [{ name: 'Expression' }]);
    ctx.reg('cbool', (val: any) => {
        if (val === vbaNull) ctx.throwError(VbaErrorCode.INVALID_USE_OF_NULL, 'Invalid use of Null');
        return vbaToBoolean(val);
    }, [{ name: 'Expression' }]);
    ctx.reg('cvar', (val: any) => val, [{ name: 'Expression' }]);
    ctx.reg('cverr', (val: any) => {
        if (val instanceof VbaErrorValue) return val;
        const code = ctx.toVbaNumber(val);
        if (code < 0 || code > 65535) ctx.throwError(VbaErrorCode.INVALID_PROCEDURE_CALL, "Invalid procedure call or argument");
        return new VbaErrorValue(code);
    }, [{ name: 'ErrorNumber' }]);

    const hexFn = (n: any) => {
        if (n === vbaNull) return vbaNull;
        const rounded = vbaRound(ctx.toVbaNumber(n), 0);
        if (rounded >= 0) return rounded.toString(16).toUpperCase();
        // §6.1.2.3.1.17: -32767〜-1 → 16-bit (4 chars), -2147483648〜-32768 → 32-bit (8 chars)
        if (rounded >= -32767) return (rounded & 0xFFFF).toString(16).toUpperCase().padStart(4, '0');
        return (rounded >>> 0).toString(16).toUpperCase().padStart(8, '0');
    };
    ctx.reg('hex', hexFn, [{ name: 'Number' }], ['$']);
    const octFn = (n: any) => {
        if (n === vbaNull) return vbaNull;
        const rounded = vbaRound(ctx.toVbaNumber(n), 0);
        if (rounded >= 0) return rounded.toString(8);
        // §6.1.2.3.1.19: -32767〜-1 → 16-bit octal (6 chars), -2147483648〜-32768 → 32-bit (11 chars)
        if (rounded >= -32767) return (rounded & 0xFFFF).toString(8).padStart(6, '0');
        return (rounded >>> 0).toString(8).padStart(11, '0');
    };
    ctx.reg('oct', octFn, [{ name: 'Number' }], ['$']);
    ctx.reg('val', (s: any) => {
        if (s === vbaNull) ctx.throwError(VbaErrorCode.INVALID_USE_OF_NULL, 'Invalid use of Null');
        // Non-string values are Let-coerced to String before parsing (VBA spec).
        if (typeof s !== 'string') s = vbaToString(s);
        const cleaned = s.trim().replace(/ /g, '');
        if (cleaned.toLowerCase().startsWith('&h')) return parseInt(cleaned.slice(2), 16) || 0;
        if (cleaned.toLowerCase().startsWith('&o')) return parseInt(cleaned.slice(2), 8) || 0;
        const match = cleaned.match(/^[+-]?\d*(\.\d*)?([eE][+-]?\d+)?/);
        return match ? parseFloat(match[0]) || 0 : 0;
    }, [{ name: 'String' }]);

    // VB string constants (settable, not const, for compatibility)
    ctx.envSet('vbcrlf', "\r\n");
    ctx.envSet('vbcr', "\r");
    ctx.envSet('vblf', "\n");
    ctx.envSet('vbtab', "\t");
    ctx.envSet('vbnullstring', "");
    ctx.envSet('vbnullchar', "\0");
    ctx.envSet('vbok', 1);
    ctx.envSet('vbcancel', 2);
    ctx.envSet('vbabort', 3);
    ctx.envSet('vbretry', 4);
    ctx.envSet('vbignore', 5);
    ctx.envSet('vbyes', 6);
    ctx.envSet('vbno', 7);
    ctx.envSet('vbmethod', 1);
    ctx.envSet('vbget', 2);
    ctx.envSet('vblet', 4);
    ctx.envSet('vbset', 8);
    ctx.envSet('vbokonly', 0);
    ctx.envSet('vbokcancel', 1);
    ctx.envSet('vbabortretryignore', 2);
    ctx.envSet('vbyesnocancel', 3);
    ctx.envSet('vbyesno', 4);
    ctx.envSet('vbretrycancel', 5);
    ctx.envSet('vbcritical', 16);
    ctx.envSet('vbquestion', 32);
    ctx.envSet('vbexclamation', 48);
    ctx.envSet('vbinformation', 64);
    ctx.envSet('vbdefaultbutton1', 0);
    ctx.envSet('vbdefaultbutton2', 256);
    ctx.envSet('vbdefaultbutton3', 512);
    ctx.envSet('vbdefaultbutton4', 768);
    ctx.envSet('vbtextcompare', 1);
    ctx.envSet('vbbinarycompare', 0);
    ctx.envSet('vbuppercase', 1);
    ctx.envSet('vblowercase', 2);
    ctx.envSet('vbpropercase', 3);
    ctx.envSet('vbwide', 4);
    ctx.envSet('vbnarrow', 8);
    ctx.envSet('vbkatakana', 16);
    ctx.envSet('vbhiragana', 32);
}

// ---------------------------------------------------------------------------
// Math functions — Abs, Atn, Cos, Exp, Int, Fix, Log, Round, Rnd, etc.
// ---------------------------------------------------------------------------

export function registerMathFunctions(ctx: StdlibCtx): void {
    ctx.reg('abs', (val: any) => val === vbaNull ? vbaNull : Math.abs(ctx.toVbaNumber(val)), [{ name: 'Number' }]);
    ctx.reg('atn', (val: any) => val === vbaNull ? vbaNull : Math.atan(ctx.toVbaNumber(val)), [{ name: 'Number' }]);
    ctx.reg('cos', (val: any) => val === vbaNull ? vbaNull : Math.cos(ctx.toVbaNumber(val)), [{ name: 'Number' }]);
    ctx.reg('exp', (val: any) => {
        if (val === vbaNull) return vbaNull;
        const n = ctx.toVbaNumber(val);
        if (n > 709.782712893) ctx.throwError(VbaErrorCode.OVERFLOW, "Overflow");
        return Math.exp(n);
    }, [{ name: 'Number' }]);
    ctx.reg('int', (val: any) => val === vbaNull ? vbaNull : Math.floor(ctx.toVbaNumber(val)), [{ name: 'Number' }]);
    ctx.reg('fix', (val: any) => {
        if (val === vbaNull) return vbaNull;
        const n = ctx.toVbaNumber(val);
        return n >= 0 ? Math.floor(n) : Math.ceil(n);
    }, [{ name: 'Number' }]);
    ctx.reg('log', (val: any) => {
        if (val === vbaNull) return vbaNull;
        const n = ctx.toVbaNumber(val);
        if (n <= 0) ctx.throwError(VbaErrorCode.INVALID_PROCEDURE_CALL, "Invalid procedure call or argument");
        return Math.log(n);
    }, [{ name: 'Number' }]);
    ctx.reg('round', (val: any, digits: any = 0) => {
        if (val === vbaNull) return vbaNull;
        if (digits !== undefined && digits === vbaNull) ctx.throwError(VbaErrorCode.TYPE_MISMATCH, 'Type mismatch');
        return ctx.round(ctx.toVbaNumber(val), Number(digits));
    }, [
        { name: 'Number' },
        { name: 'NumDigitsAfterDecimal', optional: true },
    ]);
    ctx.reg('sgn', (val: any) => {
        if (val === vbaNull) return vbaNull;
        const n = ctx.toVbaNumber(val);
        return n > 0 ? 1 : n < 0 ? -1 : 0;
    }, [{ name: 'Number' }]);
    ctx.reg('sin', (val: any) => val === vbaNull ? vbaNull : Math.sin(ctx.toVbaNumber(val)), [{ name: 'Number' }]);
    ctx.reg('sqr', (val: any) => {
        if (val === vbaNull) return vbaNull;
        const n = ctx.toVbaNumber(val);
        if (n < 0) ctx.throwError(VbaErrorCode.INVALID_PROCEDURE_CALL, "Invalid procedure call or argument");
        return Math.sqrt(n);
    }, [{ name: 'Number' }]);
    ctx.reg('tan', (val: any) => val === vbaNull ? vbaNull : Math.tan(ctx.toVbaNumber(val)), [{ name: 'Number' }]);

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
    ctx.reg('rnd', rndFunc, [{ name: 'Number', optional: true }]);
    ctx.reg('randomize', (val?: any) => {
        rndInitialized = true;
        rndSeed = (val === undefined || val === null) ? (Date.now() % 4294967296) : (Math.round(Math.abs(Number(val)) * 1000) % 4294967296);
        lastRnd = rndSeed / 4294967296;
    }, [{ name: 'Number', optional: true }]);
}

// ---------------------------------------------------------------------------
// String functions — Asc, Chr, InStr, LCase, Left, Len, Mid, Right, etc.
// ---------------------------------------------------------------------------

export function registerStringFunctions(ctx: StdlibCtx): void {
    const ascFunc = (s: any) => {
        if (s === vbaNull) return vbaNull;
        const str = String(s ?? '');
        if (str.length === 0) ctx.throwError(VbaErrorCode.INVALID_PROCEDURE_CALL, "Invalid procedure call or argument");
        return str.charCodeAt(0);
    };
    ctx.reg('asc', ascFunc, [{ name: 'String' }]);
    ctx.reg('ascw', ascFunc, [{ name: 'String' }]);
    const chrFunc = (n: any) => {
        if (n === vbaNull) return vbaNull;
        const code = Number(n);
        if (code < 0 || code > 255) ctx.throwError(VbaErrorCode.INVALID_PROCEDURE_CALL, "Invalid procedure call or argument");
        return String.fromCharCode(code);
    };
    ctx.reg('chr', chrFunc, [{ name: 'CharCode' }], ['$']);
    const chrwFunc = (n: any) => {
        if (n === vbaNull) return vbaNull;
        const code = Number(n);
        if (code < -32768 || code > 65535) ctx.throwError(VbaErrorCode.INVALID_PROCEDURE_CALL, "Invalid procedure call or argument");
        return String.fromCharCode(code < 0 ? code + 65536 : code);
    };
    ctx.reg('chrw', chrwFunc, [{ name: 'CharCode' }], ['$']);
    // Byte-oriented variants (UTF-16LE model: 1 char = 2 bytes, same as MidB)
    ctx.reg('lenb', (s: any) => s === vbaNull ? vbaNull : String(s ?? '').length * 2, [{ name: 'String' }]);
    ctx.reg('ascb', (s: any) => {
        if (s === vbaNull) return vbaNull;
        const str = String(s ?? '');
        if (str.length === 0) ctx.throwError(VbaErrorCode.INVALID_PROCEDURE_CALL, "Invalid procedure call or argument");
        return str.charCodeAt(0) & 0xFF;
    }, [{ name: 'String' }], ['$']);
    ctx.reg('chrb', (n: any) => { if (n === vbaNull) return vbaNull; return String.fromCharCode(Number(n) & 0xFF); }, [{ name: 'CharCode' }], ['$']);
    // InStr: Start は先頭にある Optional 引数のため、引数の個数で意味が変わる
    const instrFunc = (...args: any[]) => {
        let start: any = 1, s1: any, s2: any, comp: any;
        if (args.length >= 4) [start, s1, s2, comp] = args;
        else if (args.length === 3) [start, s1, s2] = args;  // arg count determines form, not type
        else [s1, s2] = args;
        if (start === vbaNull) ctx.throwError(VbaErrorCode.INVALID_USE_OF_NULL, 'Invalid use of Null');
        if (comp === vbaNull) ctx.throwError(VbaErrorCode.INVALID_USE_OF_NULL, 'Invalid use of Null');
        if (Number(start) < 1) ctx.throwError(VbaErrorCode.INVALID_PROCEDURE_CALL, "Invalid procedure call or argument");
        if (s1 === vbaNull || s2 === vbaNull) return vbaNull;
        const str1 = String(s1 ?? ''), str2 = String(s2 ?? '');
        const isText = (comp === 1) || (comp === undefined && ctx.compMode === 'Text');
        const idx = isText ? str1.toLowerCase().indexOf(str2.toLowerCase(), start - 1) : str1.indexOf(str2, start - 1);
        return idx === -1 ? 0 : idx + 1;
    };
    ctx.regOvl('instr', instrFunc, [
        { params: [{ name: 'String1' }, { name: 'String2' }] },
        { params: [{ name: 'Start' }, { name: 'String1' }, { name: 'String2' }] },
        { params: [{ name: 'Start' }, { name: 'String1' }, { name: 'String2' }, { name: 'Compare' }] },
    ]);
    // InStrB: バイト単位での検索（VBA では 1 文字 = 2 バイト）
    const instrbFunc = (...args: any[]) => {
        let startByte = 1, s1, s2, comp;
        if (args.length >= 4) [startByte, s1, s2, comp] = args;
        else if (args.length === 3 && typeof args[0] === 'number') [startByte, s1, s2] = args;
        else [s1, s2] = args;
        if (s1 === vbaNull || s2 === vbaNull) return vbaNull;
        const str1 = String(s1 ?? ''), str2 = String(s2 ?? '');
        const startChar = Math.floor((Number(startByte) - 1) / 2) + 1;
        const isText = (comp === 1) || (comp === undefined && ctx.compMode === 'Text');
        const idx = isText ? str1.toLowerCase().indexOf(str2.toLowerCase(), startChar - 1) : str1.indexOf(str2, startChar - 1);
        return idx === -1 ? 0 : idx * 2 + 1;
    };
    ctx.regOvl('instrb', instrbFunc, [
        { params: [{ name: 'String1' }, { name: 'String2' }] },
        { params: [{ name: 'Start' }, { name: 'String1' }, { name: 'String2' }] },
        { params: [{ name: 'Start' }, { name: 'String1' }, { name: 'String2' }, { name: 'Compare' }] },
    ]);
    ctx.reg('instrrev', (s1: any, s2: any, start: any = -1, comp: any = undefined) => {
        if (start === vbaNull) ctx.throwError(VbaErrorCode.INVALID_USE_OF_NULL, 'Invalid use of Null');
        if (comp === vbaNull) ctx.throwError(VbaErrorCode.INVALID_USE_OF_NULL, 'Invalid use of Null');
        const startNum = Number(start);
        if (startNum !== -1 && startNum < 1) ctx.throwError(VbaErrorCode.INVALID_PROCEDURE_CALL, "Invalid procedure call or argument");
        if (s1 === vbaNull || s2 === vbaNull) return vbaNull;
        const str = String(s1 ?? ''), find = String(s2 ?? '');
        if (str === "") return 0;
        if (startNum !== -1 && startNum > str.length) return 0;
        const effStart = (startNum === -1) ? str.length : startNum;
        if (find === "") return effStart;
        const isText = (comp === 1) || (comp === undefined && ctx.compMode === 'Text');
        const idx = isText ? str.toLowerCase().lastIndexOf(find.toLowerCase(), effStart - 1) : str.lastIndexOf(find, effStart - 1);
        return idx === -1 ? 0 : idx + 1;
    }, [
        { name: 'StringCheck' },
        { name: 'StringMatch' },
        { name: 'Start', optional: true },
        { name: 'Compare', optional: true },
    ]);
    const lcaseFunc = (val: any) => val === vbaNull ? vbaNull : String(val ?? '').toLowerCase();
    ctx.reg('lcase', lcaseFunc, [{ name: 'String' }], ['$']);
    const strFunc = (val: any) => {
        if (val === vbaNull) ctx.throwError(VbaErrorCode.INVALID_USE_OF_NULL, 'Invalid use of Null');
        const n = ctx.toVbaNumber(val);
        return n >= 0 ? " " + n : String(n);
    };
    ctx.reg('str', strFunc, [{ name: 'Number' }], ['$']);
    const ucaseFunc = (val: any) => val === vbaNull ? vbaNull : String(val ?? '').toUpperCase();
    ctx.reg('ucase', ucaseFunc, [{ name: 'String' }], ['$']);
    const leftFunc = (val: any, len: any) => {
        if (val === vbaNull || len === vbaNull) return vbaNull;
        const l = Number(len);
        if (l < 0) ctx.throwError(VbaErrorCode.INVALID_PROCEDURE_CALL, "Invalid procedure call or argument");
        return String(val ?? '').substring(0, l);
    };
    ctx.reg('left', leftFunc, [{ name: 'String' }, { name: 'Length' }], ['$']);
    const rightFunc = (val: any, len: any) => {
        if (val === vbaNull || len === vbaNull) return vbaNull;
        const s = String(val ?? ''), l = Number(len);
        if (l < 0) ctx.throwError(VbaErrorCode.INVALID_PROCEDURE_CALL, "Invalid procedure call or argument");
        return s.substring(s.length - l);
    };
    ctx.reg('right', rightFunc, [{ name: 'String' }, { name: 'Length' }], ['$']);
    const midFunc = (val: any, start: any, len?: any) => {
        if (val === vbaNull || start === vbaNull || len === vbaNull) return vbaNull;
        const s = String(val ?? ''), st = Number(start);
        if (st < 1) ctx.throwError(VbaErrorCode.INVALID_PROCEDURE_CALL, "Invalid procedure call or argument");
        if (len !== undefined && Number(len) < 0) ctx.throwError(VbaErrorCode.INVALID_PROCEDURE_CALL, "Invalid procedure call or argument");
        return len !== undefined ? s.substring(st - 1, st - 1 + Number(len)) : s.substring(st - 1);
    };
    ctx.reg('mid', midFunc, [
        { name: 'String' },
        { name: 'Start' },
        { name: 'Length', optional: true },
    ], ['$']);
    ctx.reg('len', (val: any) => val === vbaNull ? vbaNull : String(val ?? '').length, [{ name: 'String' }]);
    const ltrimFunc = (val: any) => val === vbaNull ? vbaNull : String(val ?? '').trimStart();
    ctx.reg('ltrim', ltrimFunc, [{ name: 'String' }], ['$']);
    const rtrimFunc = (val: any) => val === vbaNull ? vbaNull : String(val ?? '').trimEnd();
    ctx.reg('rtrim', rtrimFunc, [{ name: 'String' }], ['$']);
    const trimFunc = (val: any) => val === vbaNull ? vbaNull : String(val ?? '').trim();
    ctx.reg('trim', trimFunc, [{ name: 'String' }], ['$']);
    const spaceFunc = (n: any) => {
        if (n === vbaNull) return vbaNull;
        const count = Number(n);
        if (count < 0) ctx.throwError(VbaErrorCode.INVALID_PROCEDURE_CALL, "Invalid procedure call or argument");
        return ' '.repeat(count);
    };
    ctx.reg('space', spaceFunc, [{ name: 'Number' }], ['$']);
    const stringFunc = (n: any, char: any) => {
        if (n === vbaNull || char === vbaNull) return vbaNull;
        const count = Number(n);
        if (count < 0) ctx.throwError(VbaErrorCode.INVALID_PROCEDURE_CALL, "Invalid procedure call or argument");
        let c: string;
        if (typeof char === 'number') {
            // §6.1.2.11.1.38: numbers > 255 use character Mod 256
            c = String.fromCharCode(Math.trunc(char) % 256);
        } else {
            const s = String(char ?? '');
            // Empty string character is invalid per spec
            if (s.length === 0) ctx.throwError(VbaErrorCode.INVALID_PROCEDURE_CALL, "Invalid procedure call or argument");
            c = s[0];
        }
        return c.repeat(count);
    };
    ctx.reg('string', stringFunc, [{ name: 'Number' }, { name: 'Character' }], ['$']);
    ctx.reg('split', (s: any, del: any = ' ', limit: any = -1, compare: any = undefined) => {
        if (s === vbaNull) return vbaNull;
        if (del === vbaNull || limit === vbaNull || compare === vbaNull) ctx.throwError(VbaErrorCode.INVALID_USE_OF_NULL, 'Invalid use of Null');
        const str = String(s ?? '');
        const delimiter = del === null || del === undefined ? ' ' : String(del);
        const n = (limit === null || limit === undefined) ? -1 : Number(limit);
        let result: string[];
        if (str === '' || n === 0) {
            result = [];
        } else if (delimiter === '') {
            result = [str];
        } else {
            const isText = (compare === 1) || (compare === undefined && ctx.compMode === 'Text');
            const cmpStr = isText ? str.toLowerCase() : str;
            const cmpDel = isText ? delimiter.toLowerCase() : delimiter;
            result = [];
            let pos = 0;
            for (;;) {
                if (n > 0 && result.length === n - 1) { result.push(str.substring(pos)); break; }
                const idx = cmpStr.indexOf(cmpDel, pos);
                if (idx === -1) { result.push(str.substring(pos)); break; }
                result.push(str.substring(pos, idx));
                pos = idx + delimiter.length;
            }
        }
        (result as any).vbaBase = 0;
        return result;
    }, [
        { name: 'Expression' },
        { name: 'Delimiter', optional: true },
        { name: 'Limit', optional: true },
        { name: 'Compare', optional: true },
    ]);
    ctx.reg('join', (arr: any, del: any = ' ') => {
        if (del === vbaNull) return vbaNull;
        if (!Array.isArray(arr)) ctx.throwError(VbaErrorCode.TYPE_MISMATCH, 'Type mismatch');
        const base = (arr as any).vbaBase || 0;
        const elems = arr.slice(base).map((el: any) => {
            if (el === vbaNull) ctx.throwError(VbaErrorCode.TYPE_MISMATCH, 'Type mismatch');
            return el === undefined || el === vbaEmpty ? '' : String(el);
        });
        return elems.join(String(del));
    }, [
        { name: 'SourceArray' },
        { name: 'Delimiter', optional: true },
    ]);
    ctx.reg('replace', (s: any, f: any, r: any, start: any = 1, count: any = -1, compare: any = undefined) => {
        if (s === vbaNull) ctx.throwError(VbaErrorCode.INVALID_USE_OF_NULL, 'Invalid use of Null');
        if (f === vbaNull || r === vbaNull) return vbaNull;
        const str = String(s ?? '');
        const find = String(f ?? '');
        const repl = String(r ?? '');
        const startNum = Number(start ?? 1);
        if (startNum < 1) ctx.throwError(VbaErrorCode.INVALID_PROCEDURE_CALL, 'Invalid procedure call or argument');
        const countNum = Number(count ?? -1);
        const isText = (compare === 1) || (compare === undefined && ctx.compMode === 'Text');
        // Slice from start position (1-based), operate, then return from that offset
        const prefix = str.substring(0, startNum - 1);
        const working = str.substring(startNum - 1);
        if (find === '') return working;
        const findLower = isText ? find.toLowerCase() : find;
        let result = '';
        let remaining = working;
        let replacements = 0;
        while (remaining.length > 0) {
            const searchIn = isText ? remaining.toLowerCase() : remaining;
            const idx = searchIn.indexOf(findLower);
            if (idx === -1 || (countNum >= 0 && replacements >= countNum)) {
                result += remaining;
                break;
            }
            result += remaining.substring(0, idx) + repl;
            remaining = remaining.substring(idx + find.length);
            replacements++;
        }
        // VBA Replace returns from start position (prefix is NOT included)
        void prefix;
        return result;
    }, [
        { name: 'Expression' },
        { name: 'Find' },
        { name: 'Replace' },
        { name: 'Start', optional: true },
        { name: 'Count', optional: true },
        { name: 'Compare', optional: true },
    ]);
    ctx.reg('strcomp', (s1: any, s2: any, comp?: number) => {
        if (s1 === vbaNull || s2 === vbaNull) return vbaNull;
        if (comp !== undefined && (comp as any) === vbaNull) ctx.throwError(VbaErrorCode.INVALID_USE_OF_NULL, 'Invalid use of Null');
        let str1 = String(s1 ?? ''), str2 = String(s2 ?? '');
        const isText = (comp === 1) || (comp === undefined && ctx.compMode === 'Text');
        if (isText) { str1 = str1.toLowerCase(); str2 = str2.toLowerCase(); }
        return str1 < str2 ? -1 : (str1 > str2 ? 1 : 0);
    }, [
        { name: 'String1' },
        { name: 'String2' },
        { name: 'Compare', optional: true },
    ]);
    ctx.reg('strconv', (s: any, conv: any) => {
        if (s === vbaNull) return vbaNull;
        let str = String(s ?? '');
        const c = Number(conv);
        const caseConv = c & 3;
        if (caseConv === 1) str = str.toUpperCase();
        else if (caseConv === 2) str = str.toLowerCase();
        else if (caseConv === 3) str = str.split(/\s+/).map(w => w.charAt(0).toUpperCase() + w.slice(1).toLowerCase()).join(' ');
        if (c & 16) {
            str = str.replace(/[ぁ-ゖ]/g, m => String.fromCharCode(m.charCodeAt(0) + 0x60));
        } else if (c & 32) {
            str = str.replace(/[ァ-ヶ]/g, m => String.fromCharCode(m.charCodeAt(0) - 0x60));
        }
        if (c & 4) {
            str = str.replace(/[!-~]/g, m => String.fromCharCode(m.charCodeAt(0) + 0xFEE0)).replace(/ /g, '　');
        } else if (c & 8) {
            str = str.replace(/[！-～]/g, m => String.fromCharCode(m.charCodeAt(0) - 0xFEE0)).replace(/　/g, ' ');
        }
        return str;
    }, [{ name: 'String' }, { name: 'Conversion' }, { name: 'LCID', optional: true }]);
    ctx.reg('strreverse', (s: any) => {
        if (s === vbaNull) ctx.throwError(VbaErrorCode.INVALID_USE_OF_NULL, 'Invalid use of Null');
        return String(s ?? '').split('').reverse().join('');
    }, [{ name: 'Expression' }]);
    ctx.reg('filter', (source: any, match: any, include: any = vbaTrue, compare: any = undefined) => {
        if (!Array.isArray(source)) ctx.throwError(VbaErrorCode.TYPE_MISMATCH, "Type mismatch");
        const srcDims = (source as any).__vbaDimensions__;
        if (srcDims && srcDims.length > 1) ctx.throwError(VbaErrorCode.TYPE_MISMATCH, "Type mismatch");
        if (match === vbaNull) ctx.throwError(VbaErrorCode.TYPE_MISMATCH, "Type mismatch");
        const find = String(match ?? '');
        const isInclude = ctx.isTrue(include);
        const isText = (compare === 1) || (compare === undefined && ctx.compMode === 'Text');
        const result = source.filter((s: any) => {
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
    ctx.reg('leftb', (val: any, len: any) => {
        if (val === vbaNull || len === vbaNull) return vbaNull;
        const s = String(val ?? '');
        return s.substring(0, Math.floor(Number(len) / 2));
    }, [{ name: 'String' }, { name: 'Length' }]);
    ctx.reg('rightb', (val: any, len: any) => {
        if (val === vbaNull || len === vbaNull) return vbaNull;
        const s = String(val ?? '');
        const charLen = Math.floor(Number(len) / 2);
        return s.substring(s.length - charLen);
    }, [{ name: 'String' }, { name: 'Length' }]);
    const midbFunc = (val: any, start: any, len?: any) => {
        if (val === vbaNull || start === vbaNull || len === vbaNull) return vbaNull;
        const s = String(val ?? '');
        const charStart = Math.floor((Number(start) + 1) / 2);
        if (len === undefined) return s.substring(charStart - 1);
        const charLen = Math.floor(Number(len) / 2);
        return s.substring(charStart - 1, charStart - 1 + charLen);
    };
    ctx.reg('midb', midbFunc, [
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
            if (fmtLower === 'yes/no' || fmtLower === 'on/off' || fmtLower === 'true/false') {
                const isTrue = val instanceof VbaBoolean ? val.value !== 0 : (Number(val) !== 0);
                if (fmtLower === 'yes/no') return isTrue ? 'Yes' : 'No';
                if (fmtLower === 'on/off') return isTrue ? 'On' : 'Off';
                return isTrue ? 'True' : 'False';
            }
            const numVal = (val instanceof VbaBoolean) ? val.value
                : (val instanceof VbaCurrency || val instanceof VbaDecimal) ? Number(val.toString())
                : val;
            if (typeof numVal === 'number') return formatNumber(numVal, fmt);
            return String(val);
        }
        if (dateNamedFormats.includes(fmtLower)) {
            const dateVal = val instanceof VbaDate ? fromVbaDate(val.value) : (typeof val === 'number' ? fromVbaDate(val) : new Date(String(val)));
            return formatDate(dateVal, fmt);
        }
        const effectiveVal = (val instanceof VbaBoolean) ? val.value
            : (val instanceof VbaCurrency || val instanceof VbaDecimal) ? Number(val.toString())
            : val;
        const isDatePattern = /y|m|d|h|n|s|am\/pm/i.test(fmt);
        if (typeof effectiveVal === 'string') {
            // If the format contains date/time symbols, try to parse the string as a date first
            if (isDatePattern && !/^[0#,.%]+$/.test(fmt)) {
                try {
                    const parsed = parseVbaDate(effectiveVal);
                    return formatDate(parsed, fmt);
                } catch { /* fall through to string formatting */ }
            }
            return formatString(effectiveVal, fmt);
        }
        if (effectiveVal instanceof VbaDate) return formatDate(fromVbaDate(effectiveVal.value), fmt);
        if (typeof effectiveVal === 'number') {
            if (isDatePattern && !/^[0#,.%]+$/.test(fmt)) return formatDate(fromVbaDate(effectiveVal), fmt);
            return formatNumber(effectiveVal, fmt);
        }
        return String(effectiveVal);
    };
    ctx.reg('format', formatFunc, [
        { name: 'Expression' },
        { name: 'Format', optional: true },
        { name: 'FirstDayOfWeek', optional: true },
        { name: 'FirstWeekOfYear', optional: true },
    ], ['$']);

    // Helper shared by FormatCurrency / FormatNumber / FormatPercent
    const fmtNumeric = (val: any, digits: any, _leadingDigit: any, _parens: any, groupDigits: any,
                        prefix: string, suffix: string, scale: number): string => {
        if (val === vbaNull) return '';
        let n: number;
        if (val instanceof VbaDate) n = val.value;
        else if (val instanceof VbaCurrency) n = Number(val.internal) / 10000;
        else if (val instanceof VbaDecimal) n = Number(val.mantissa) / Math.pow(10, val.scale);
        else n = ctx.toVbaNumber(val);
        n *= scale;
        const dec = digits === vbaMissing || Number(digits) < 0 ? 2 : Number(digits);
        const group = groupDigits === vbaMissing || Number(groupDigits) < 0 ? true : Number(groupDigits) !== 0;
        const neg = n < 0;
        const abs = Math.abs(n);
        let formatted = vbaRound(abs, dec).toFixed(dec);
        if (group) {
            const [intPart, fracPart] = formatted.split('.');
            const grouped = intPart.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
            formatted = fracPart !== undefined ? `${grouped}.${fracPart}` : grouped;
        }
        const result = `${prefix}${formatted}${suffix}`;
        return neg ? `-${result}` : result;
    };
    ctx.reg('formatcurrency', (val: any, digits: any = vbaMissing, lead: any = vbaMissing, parens: any = vbaMissing, group: any = vbaMissing) =>
        fmtNumeric(val, digits, lead, parens, group, '$', '', 1), [
        { name: 'Expression' }, { name: 'NumDigitsAfterDecimal', optional: true },
        { name: 'IncludeLeadingDigit', optional: true }, { name: 'UseParensForNegativeNumbers', optional: true },
        { name: 'GroupDigits', optional: true },
    ], ['$']);
    ctx.reg('formatnumber', (val: any, digits: any = vbaMissing, lead: any = vbaMissing, parens: any = vbaMissing, group: any = vbaMissing) =>
        fmtNumeric(val, digits, lead, parens, group, '', '', 1), [
        { name: 'Expression' }, { name: 'NumDigitsAfterDecimal', optional: true },
        { name: 'IncludeLeadingDigit', optional: true }, { name: 'UseParensForNegativeNumbers', optional: true },
        { name: 'GroupDigits', optional: true },
    ], ['$']);
    ctx.reg('formatpercent', (val: any, digits: any = vbaMissing, lead: any = vbaMissing, parens: any = vbaMissing, group: any = vbaMissing) =>
        fmtNumeric(val, digits, lead, parens, group, '', '%', 100), [
        { name: 'Expression' }, { name: 'NumDigitsAfterDecimal', optional: true },
        { name: 'IncludeLeadingDigit', optional: true }, { name: 'UseParensForNegativeNumbers', optional: true },
        { name: 'GroupDigits', optional: true },
    ], ['$']);
    ctx.reg('formatdatetime', (val: any, namedFmt: any = 0) => {
        if (val === vbaNull) return '';
        const d = (val instanceof VbaDate) ? fromVbaDate(val.value) : parseVbaDate(val);
        const fmt = Number(namedFmt ?? 0);
        const pad2 = (n: number) => String(n).padStart(2, '0');
        const mo = d.getMonth() + 1, dy = d.getDate(), yr = d.getFullYear();
        const hh = d.getHours(), mm = d.getMinutes(), ss = d.getSeconds();
        const h12 = hh % 12 || 12, ampm = hh < 12 ? 'AM' : 'PM';
        const dayNames = ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'];
        const monNames = ['January','February','March','April','May','June','July','August','September','October','November','December'];
        if (fmt === 1) return `${dayNames[d.getDay()]}, ${monNames[mo-1]} ${dy}, ${yr}`;
        if (fmt === 2) return `${mo}/${dy}/${yr}`;
        if (fmt === 3) return `${h12}:${pad2(mm)}:${pad2(ss)} ${ampm}`;
        if (fmt === 4) return `${pad2(hh)}:${pad2(mm)}`;
        // fmt === 0: general date — date only if time=0, else date+time
        const hasTime = hh !== 0 || mm !== 0 || ss !== 0;
        const datePart = `${mo}/${dy}/${yr}`;
        return hasTime ? `${datePart} ${h12}:${pad2(mm)}:${pad2(ss)} ${ampm}` : datePart;
    }, [{ name: 'Date' }, { name: 'NamedFormat', optional: true }], ['$']);
}

// ---------------------------------------------------------------------------
// Date/time stdlib functions — Year, Month, Day, DateAdd, DateDiff, etc.
// ---------------------------------------------------------------------------

export function registerStdlibDateTimeFunctions(ctx: StdlibCtx): void {
    ctx.reg('year',   (d: any) => d === vbaNull ? vbaNull : parseVbaDate(d).getFullYear(), [{ name: 'Date' }]);
    ctx.reg('month',  (d: any) => d === vbaNull ? vbaNull : parseVbaDate(d).getMonth() + 1, [{ name: 'Date' }]);
    ctx.reg('day',    (d: any) => d === vbaNull ? vbaNull : parseVbaDate(d).getDate(), [{ name: 'Date' }]);
    ctx.reg('hour',   (d: any) => d === vbaNull ? vbaNull : parseVbaDate(d).getHours(), [{ name: 'Time' }]);
    ctx.reg('minute', (d: any) => d === vbaNull ? vbaNull : parseVbaDate(d).getMinutes(), [{ name: 'Time' }]);
    ctx.reg('second', (d: any) => d === vbaNull ? vbaNull : parseVbaDate(d).getSeconds(), [{ name: 'Time' }]);
    ctx.reg('dateserial', (y: any, m: any, d: any) => {
        if (y === vbaNull || m === vbaNull || d === vbaNull) return vbaNull;
        let year = Number(y);
        // VBA spec §6.1.2.4.1.4: 0-29 → 2000-2029, 30-99 → 1930-1999
        if (year >= 0 && year <= 29) year += 2000;
        else if (year >= 30 && year <= 99) year += 1900;
        const date = new Date(year, Number(m) - 1, Number(d));
        date.setFullYear(year); // prevent JS legacy year offset for 0-99
        return new VbaDate(toVbaDate(date));
    }, [{ name: 'Year' }, { name: 'Month' }, { name: 'Day' }]);
    ctx.reg('timeserial', (h: any, n: any, s: any) => {
        if (h === vbaNull || n === vbaNull || s === vbaNull) return vbaNull;
        return new VbaDate(toVbaDate(new Date(1899, 11, 30, Number(h), Number(n), Number(s))));
    }, [{ name: 'Hour' }, { name: 'Minute' }, { name: 'Second' }]);
    ctx.reg('weekday', (d: any, firstdayofweek: any = 1) => {
        if (d === vbaNull) return vbaNull;
        if (firstdayofweek === vbaNull) ctx.throwError(VbaErrorCode.INVALID_PROCEDURE_CALL, 'Invalid procedure call or argument');
        const dayOfWeek = parseVbaDate(d).getDay(); // 0=Sun
        let fdow = Number(firstdayofweek ?? 1);
        if (fdow === 0) fdow = 1; // vbUseSystemDayOfWeek → treat as vbSunday
        const weekStart = fdow <= 1 ? 0 : fdow - 1; // convert VBA 1-based to JS 0-based
        return ((dayOfWeek - weekStart + 7) % 7) + 1;
    }, [{ name: 'Date' }, { name: 'FirstDayOfWeek', optional: true }]);
    ctx.reg('dateadd', (interval: any, number: any, date: any) => {
        if (date === vbaNull || number === vbaNull) return vbaNull;
        const d = parseVbaDate(date);
        const n = Number(number);
        const intv = String(interval).toLowerCase();
        if (intv === 'yyyy' || intv === 'q' || intv === 'm') {
            const oldDay = d.getDate();
            if (intv === 'yyyy') d.setFullYear(d.getFullYear() + n);
            else if (intv === 'q') d.setMonth(d.getMonth() + n * 3);
            else if (intv === 'm') d.setMonth(d.getMonth() + n);
            if (d.getDate() !== oldDay) d.setDate(0);
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
        } else {
            ctx.throwError(VbaErrorCode.INVALID_PROCEDURE_CALL, 'Invalid procedure call or argument');
        }
        return new VbaDate(toVbaDate(d));
    }, [{ name: 'Interval' }, { name: 'Number' }, { name: 'Date' }]);
    ctx.reg('datediff', (interval: any, date1: any, date2: any, firstdayofweek: any = 1, _firstweekofyear: any = 1) => {
        if (date1 === vbaNull || date2 === vbaNull) return vbaNull;
        const d1 = parseVbaDate(date1);
        const d2 = parseVbaDate(date2);
        const intv = String(interval).toLowerCase();
        const diffMs = d2.getTime() - d1.getTime();
        if (intv === 'yyyy') return d2.getFullYear() - d1.getFullYear();
        else if (intv === 'q') return (d2.getFullYear() - d1.getFullYear()) * 4 + Math.floor(d2.getMonth() / 3) - Math.floor(d1.getMonth() / 3);
        else if (intv === 'm') return (d2.getFullYear() - d1.getFullYear()) * 12 + d2.getMonth() - d1.getMonth();
        else if (intv === 'y' || intv === 'd' || intv === 'w') return Math.round(diffMs / 86400000);
        else if (intv === 'ww') {
            // Week count depends on firstdayofweek
            let fdow = Number(firstdayofweek ?? 1);
            if (fdow === 0) fdow = 1;
            const weekStart = fdow <= 1 ? 0 : fdow - 1;
            const day1 = new Date(d1); day1.setHours(0, 0, 0, 0);
            const day2 = new Date(d2); day2.setHours(0, 0, 0, 0);
            // Align both dates to their week boundary
            const offset1 = (day1.getDay() - weekStart + 7) % 7;
            const offset2 = (day2.getDay() - weekStart + 7) % 7;
            const week1Start = new Date(day1); week1Start.setDate(day1.getDate() - offset1);
            const week2Start = new Date(day2); week2Start.setDate(day2.getDate() - offset2);
            return Math.round((week2Start.getTime() - week1Start.getTime()) / 604800000);
        }
        else if (intv === 'h') return Math.round(diffMs / 3600000);
        else if (intv === 'n') return Math.round(diffMs / 60000);
        else if (intv === 's') return Math.round(diffMs / 1000);
        else ctx.throwError(VbaErrorCode.INVALID_PROCEDURE_CALL, 'Invalid procedure call or argument');
    }, [
        { name: 'Interval' }, { name: 'Date1' }, { name: 'Date2' },
        { name: 'FirstDayOfWeek', optional: true },
        { name: 'FirstWeekOfYear', optional: true },
    ]);
    ctx.reg('datepart', (interval: any, date: any, firstdayofweek: any = 1, _firstweekofyear: any = 1) => {
        if (date === vbaNull) return vbaNull;
        const d = parseVbaDate(date);
        const intv = String(interval).toLowerCase();
        // firstdayofweek: 1=Sunday(default), 2=Monday, ..., 7=Saturday; 0=system default(treat as 1)
        const fdow = Math.max(0, Number(firstdayofweek ?? 1));
        const weekStart = fdow <= 1 ? 0 : fdow - 1; // JS: 0=Sun,1=Mon,...,6=Sat
        if (intv === 'yyyy') return d.getFullYear();
        else if (intv === 'q') return Math.floor(d.getMonth() / 3) + 1;
        else if (intv === 'm') return d.getMonth() + 1;
        else if (intv === 'y') {
            const start = new Date(d.getFullYear(), 0, 0);
            const diff = d.getTime() - start.getTime();
            return Math.floor(diff / 86400000);
        }
        else if (intv === 'd') return d.getDate();
        else if (intv === 'w') {
            // 'w' returns weekday number relative to firstdayofweek
            const dayOfWeek = d.getDay(); // 0=Sun
            return ((dayOfWeek - weekStart + 7) % 7) + 1;
        }
        else if (intv === 'ww') {
            const jan1 = new Date(d.getFullYear(), 0, 1);
            const jan1Day = jan1.getDay(); // 0=Sun
            const dayOfYear = Math.floor((d.getTime() - jan1.getTime()) / 86400000);
            const offset = (jan1Day - weekStart + 7) % 7;
            return Math.floor((dayOfYear + offset) / 7) + 1;
        }
        else if (intv === 'h') return d.getHours();
        else if (intv === 'n') return d.getMinutes();
        else if (intv === 's') return d.getSeconds();
        else ctx.throwError(VbaErrorCode.INVALID_PROCEDURE_CALL, 'Invalid procedure call or argument');
    }, [
        { name: 'Interval' },
        { name: 'Date' },
        { name: 'FirstDayOfWeek', optional: true },
        { name: 'FirstWeekOfYear', optional: true },
    ]);
    ctx.reg('datevalue', (val: any) => {
        if (val === vbaNull) return vbaNull;
        const d = parseVbaDate(val);
        return new VbaDate(Math.floor(toVbaDate(d)));
    }, [{ name: 'Date' }]);
    ctx.reg('timevalue', (val: any) => {
        if (val === vbaNull) return vbaNull;
        const d = parseVbaDate(val);
        const serial = toVbaDate(d);
        return new VbaDate(serial - Math.floor(serial));
    }, [{ name: 'Time' }]);
    ctx.reg('monthname', (month: any, abbreviate: any = vbaFalse) => {
        if (month === vbaNull) return vbaNull;
        const m = Number(month);
        if (m < 1 || m > 12) ctx.throwError(VbaErrorCode.INVALID_PROCEDURE_CALL, "Invalid procedure call or argument");
        const names = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
        const abbrs = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
        return ctx.isTrue(abbreviate) ? abbrs[m - 1] : names[m - 1];
    }, [{ name: 'Month' }, { name: 'Abbreviate', optional: true }]);
    ctx.reg('weekdayname', (weekday: any, abbreviate: any = vbaFalse, firstdayofweek: any = 1) => {
        if (weekday === vbaNull) return vbaNull;
        const w = Number(weekday);
        let first = Number(firstdayofweek);
        if (first === 0) first = 1;
        if (w < 1 || w > 7) ctx.throwError(VbaErrorCode.INVALID_PROCEDURE_CALL, "Invalid procedure call or argument");
        const names = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
        const abbrs = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
        const idx = (w + first - 2) % 7;
        return ctx.isTrue(abbreviate) ? abbrs[idx] : names[idx];
    }, [
        { name: 'Weekday' },
        { name: 'Abbreviate', optional: true },
        { name: 'FirstDayOfWeek', optional: true },
    ]);
}

// ---------------------------------------------------------------------------
// Interaction functions — Shell, MsgBox, InputBox, AppActivate, etc.
// ---------------------------------------------------------------------------

export function registerInteractionFunctions(ctx: StdlibCtx): void {
    ctx.reg('shell', (cmd: any, style: any = 1) => { ctx.print(`[SHELL] ${cmd} (Style: ${style})`); return 1; }, [
        { name: 'PathName' },
        { name: 'WindowStyle', optional: true },
    ]);
    ctx.reg('msgbox', (msg: any, _buttons: any = 0, _title: any = "", _helpFile?: string, _context?: number) => {
        const title = _title ? ` ${_title}:` : '';
        ctx.print(`[MSGBOX]${title} ${msg}`);
        return 1;
    }, [
        { name: 'Prompt' },
        { name: 'Buttons', optional: true },
        { name: 'Title', optional: true },
        { name: 'HelpFile', optional: true },
        { name: 'Context', optional: true },
    ]);
    ctx.reg('inputbox', (prompt: any, _title: any = "", def: any = "", _xpos?: number, _ypos?: number, _helpFile?: string, _context?: number) => {
        ctx.print(`[INPUTBOX] ${prompt}`);
        return def;
    }, [
        { name: 'Prompt' },
        { name: 'Title', optional: true },
        { name: 'Default', optional: true },
        { name: 'XPos', optional: true },
        { name: 'YPos', optional: true },
        { name: 'HelpFile', optional: true },
        { name: 'Context', optional: true },
    ]);
    ctx.reg('appactivate', (title: string, _wait?: boolean) => { ctx.print(`[APPACTIVATE] ${title}`); }, [
        { name: 'Title' },
        { name: 'Wait', optional: true },
    ]);
    ctx.reg('sendkeys', (keys: string, _wait?: boolean) => { ctx.print(`[SENDKEYS] ${keys}`); }, [
        { name: 'String' },
        { name: 'Wait', optional: true },
    ]);
    ctx.reg('doevents', () => 0, []);
}

// ---------------------------------------------------------------------------
// Financial functions — FV, PV, PMT, NPer, Rate, SLN, SYD, DDB, IRR, etc.
// ---------------------------------------------------------------------------

export function registerFinancialFunctions(ctx: StdlibCtx): void {
    const getRateFactor = (rate: number, nper: number) => {
        if (rate === 0) return nper;
        return (Math.pow(1 + rate, nper) - 1) / rate;
    };
    ctx.reg('fv', (rate: any, nper: any, pmt: any, pv: any = 0, type: any = 0) => {
        const r = Number(rate), n = Number(nper), p = Number(pmt), v = Number(pv), t = Number(type);
        const factor = getRateFactor(r, n);
        return -(v * Math.pow(1 + r, n) + p * (1 + r * t) * factor);
    }, [
        { name: 'Rate' }, { name: 'NPer' }, { name: 'Pmt' },
        { name: 'PV', optional: true }, { name: 'Type', optional: true },
    ]);
    ctx.reg('pv', (rate: any, nper: any, pmt: any, fv: any = 0, type: any = 0) => {
        const r = Number(rate), n = Number(nper), p = Number(pmt), f = Number(fv), t = Number(type);
        if (r === 0) return -(f + p * n);
        const p1 = Math.pow(1 + r, n);
        return -(f + p * (1 + r * t) * ((p1 - 1) / r)) / p1;
    }, [
        { name: 'Rate' }, { name: 'NPer' }, { name: 'Pmt' },
        { name: 'FV', optional: true }, { name: 'Type', optional: true },
    ]);
    ctx.reg('pmt', (rate: any, nper: any, pv: any, fv: any = 0, type: any = 0) => {
        const r = Number(rate), n = Number(nper), v = Number(pv), f = Number(fv), t = Number(type);
        if (r === 0) return -(v + f) / n;
        const p1 = Math.pow(1 + r, n);
        return -(v * p1 + f) / ((1 + r * t) * ((p1 - 1) / r));
    }, [
        { name: 'Rate' }, { name: 'NPer' }, { name: 'PV' },
        { name: 'FV', optional: true }, { name: 'Type', optional: true },
    ]);
    ctx.reg('nper', (rate: any, pmt: any, pv: any, fv: any = 0, type: any = 0) => {
        const r = Number(rate), p = Number(pmt), v = Number(pv), f = Number(fv), t = Number(type);
        if (r === 0) return -(v + f) / p;
        const num = p * (1 + r * t) - f * r;
        const den = p * (1 + r * t) + v * r;
        return Math.log(num / den) / Math.log(1 + r);
    }, [
        { name: 'Rate' }, { name: 'Pmt' }, { name: 'PV' },
        { name: 'FV', optional: true }, { name: 'Type', optional: true },
    ]);
    ctx.reg('rate', (nper: any, pmt: any, pv: any, fv: any = 0, type: any = 0, guess: any = 0.1) => {
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
    ctx.reg('sln', (cost: any, salvage: any, life: any) => {
        return (Number(cost) - Number(salvage)) / Number(life);
    }, [{ name: 'Cost' }, { name: 'Salvage' }, { name: 'Life' }]);
    ctx.reg('syd', (cost: any, salvage: any, life: any, period: any) => {
        const c = Number(cost), s = Number(salvage), l = Number(life), p = Number(period);
        return ((c - s) * (l - p + 1) * 2) / (l * (l + 1));
    }, [{ name: 'Cost' }, { name: 'Salvage' }, { name: 'Life' }, { name: 'Period' }]);
    ctx.reg('ddb', (cost: any, salvage: any, life: any, period: any, factor: any = 2) => {
        const c = Number(cost), s = Number(salvage), l = Number(life), p = Number(period), f = Number(factor);
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
    ctx.reg('irr', (values: any, guess: any = 0.1) => {
        if (!Array.isArray(values)) ctx.throwError(VbaErrorCode.TYPE_MISMATCH, "Type mismatch");
        const v = values.map(Number);
        let r = Number(guess);
        for (let i = 0; i < 100; i++) {
            let npv = 0, dnpv = 0;
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
    ctx.reg('mirr', (values: any, finance_rate: any, reinvest_rate: any) => {
        if (!Array.isArray(values)) ctx.throwError(VbaErrorCode.TYPE_MISMATCH, "Type mismatch");
        const v = values.map(Number);
        const fr = Number(finance_rate), rr = Number(reinvest_rate);
        const n = v.length - 1;
        let npv_neg = 0, npv_pos = 0;
        for (let t = 0; t < v.length; t++) {
            if (v[t] < 0) npv_neg += v[t] / Math.pow(1 + fr, t);
            else npv_pos += v[t] / Math.pow(1 + rr, t);
        }
        const tv = npv_pos * Math.pow(1 + rr, n);
        return Math.pow(-tv / npv_neg, 1 / n) - 1;
    }, [{ name: 'ValueArray' }, { name: 'FinanceRate' }, { name: 'ReinvestRate' }]);
    ctx.reg('npv', (rate: any, values: any) => {
        if (!Array.isArray(values)) ctx.throwError(VbaErrorCode.TYPE_MISMATCH, "Type mismatch");
        const r = Number(rate);
        const base: number = (values as any).vbaBase ?? 0;
        let result = 0;
        let period = 1;
        for (let idx = base; idx < values.length; idx++, period++) {
            result += Number(values[idx]) / Math.pow(1 + r, period);
        }
        return result;
    }, [{ name: 'Rate' }, { name: 'ValueArray' }]);
    ctx.reg('ipmt', (rate: any, per: any, nper: any, pv: any, fv: any = 0, type: any = 0) => {
        const r = Number(rate), p = Number(per), n = Number(nper), v = Number(pv), f = Number(fv), t = Number(type);
        const pmt = Number(ctx.envGet('pmt')(r, n, v, f, t));
        let ipmt: number;
        if (p === 1) {
            ipmt = t === 1 ? 0 : -v * r;
        } else {
            const fv_prev = Number(ctx.envGet('fv')(r, p - 1, pmt, v, t));
            ipmt = -fv_prev * r;
        }
        return ipmt;
    }, [
        { name: 'Rate' }, { name: 'Per' }, { name: 'NPer' }, { name: 'PV' },
        { name: 'FV', optional: true }, { name: 'Type', optional: true },
    ]);
    ctx.reg('ppmt', (rate: any, per: any, nper: any, pv: any, fv: any = 0, type: any = 0) => {
        const r = Number(rate), p = Number(per), n = Number(nper), v = Number(pv), f = Number(fv), t = Number(type);
        const pmt = Number(ctx.envGet('pmt')(r, n, v, f, t));
        const ipmt = Number(ctx.envGet('ipmt')(r, p, n, v, f, t));
        return pmt - ipmt;
    }, [
        { name: 'Rate' }, { name: 'Per' }, { name: 'NPer' }, { name: 'PV' },
        { name: 'FV', optional: true }, { name: 'Type', optional: true },
    ]);
}

// ---------------------------------------------------------------------------
// Constants — vbSunday, vbMonday, VarType constants, Error$, Environ, etc.
// ---------------------------------------------------------------------------

export function registerConstants(ctx: StdlibCtx): void {
    const errorMessages: Record<number, string> = {
        5: "Invalid procedure call or argument", 6: "Overflow", 9: "Subscript out of range",
        11: "Division by zero", 13: "Type mismatch", 52: "Bad file name or number",
        53: "File not found", 58: "File already exists", 62: "Input past end of file",
        70: "Permission denied", 76: "Path not found", 91: "Object variable not set",
        94: "Invalid use of Null",
    };
    const errFunc = (n?: any) => errorMessages[n === undefined ? ctx.errNum() : Number(n)] || "Application-defined or object-defined error";
    ctx.reg('error', errFunc, [{ name: 'ErrorNumber', optional: true }], ['$']);
    ctx.envSetConst('vbsunday', 1);
    ctx.envSetConst('vbmonday', 2);
    ctx.envSetConst('vbtuesday', 3);
    ctx.envSetConst('vbwednesday', 4);
    ctx.envSetConst('vbthursday', 5);
    ctx.envSetConst('vbfriday', 6);
    ctx.envSetConst('vbsaturday', 7);
    ctx.envSetConst('vbusesystem', 0);
    ctx.envSetConst('vbbinarycompare', 0);
    ctx.envSetConst('vbtextcompare', 1);
    ctx.envSetConst('vbempty', 0); ctx.envSetConst('vbnull', 1); ctx.envSetConst('vbinteger', 2); ctx.envSetConst('vblong', 3); ctx.envSetConst('vbsingle', 4); ctx.envSetConst('vbdouble', 5); ctx.envSetConst('vbcurrency', 6); ctx.envSetConst('vbdate', 7); ctx.envSetConst('vbstring', 8); ctx.envSetConst('vbobject', 9); ctx.envSetConst('vberror', 10); ctx.envSetConst('vbboolean', 11); ctx.envSetConst('vbvariant', 12); ctx.envSetConst('vbdataobject', 13); ctx.envSetConst('vbdecimal', 14); ctx.envSetConst('vbbyte', 17); ctx.envSetConst('vblonglong', 20); ctx.envSetConst('vbuserdefinedtype', 36); ctx.envSetConst('vbarray', 8192);
    ctx.envSetConst('vbfirstjan1', 1); ctx.envSetConst('vbfirstfourdays', 2); ctx.envSetConst('vbfirstfullweek', 3);
    ctx.envSetConst('vbcrlf', "\r\n"); ctx.envSetConst('vbtab', "\t"); ctx.envSetConst('vbcr', "\r"); ctx.envSetConst('vblf', "\n"); ctx.envSetConst('vbnewline', "\n"); ctx.envSetConst('vbnullstring', ''); ctx.envSetConst('vbnullchar', '\0'); ctx.envSetConst('vbback', "\b"); ctx.envSetConst('vbformfeed', "\f");
    ctx.envSetConst('vbnormal', 0); ctx.envSetConst('vbreadonly', 1); ctx.envSetConst('vbhidden', 2); ctx.envSetConst('vbsystem', 4); ctx.envSetConst('vbvolume', 8); ctx.envSetConst('vbdirectory', 16); ctx.envSetConst('vbarchive', 32); ctx.envSetConst('vbalias', 64);
    ctx.envSetConst('vbobjecterror', -2147221504);
    ctx.envSetConst('true', vbaTrue); ctx.envSetConst('false', vbaFalse); ctx.envSetConst('empty', vbaEmpty); ctx.envSetConst('nothing', vbaNothing); ctx.envSetConst('null', vbaNull);

    ctx.reg('environ', (k: any) => ctx.getEnv(k), [{ name: 'EnvString' }], ['$']);
    ctx.reg('rgb', (r: any, g: any, b: any) => {
        const clamp = (n: number) => Math.min(255, Math.max(0, Math.round(n)));
        return clamp(Number(r)) + clamp(Number(g)) * 256 + clamp(Number(b)) * 65536;
    }, [{ name: 'Red' }, { name: 'Green' }, { name: 'Blue' }]);
    const qbColorTable = [0, 8388608, 32768, 8421376, 128, 8388736, 32896, 12632256,
                          8421504, 16711680, 65280, 16776960, 255, 16711935, 65535, 16777215];
    ctx.reg('qbcolor', (c: any) => {
        const idx = Math.round(Number(c));
        if (idx < 0 || idx > 15) ctx.throwError(VbaErrorCode.INVALID_PROCEDURE_CALL, "Invalid procedure call or argument");
        return qbColorTable[idx];
    }, [{ name: 'Color' }]);
    ctx.reg('nz', (val: any, valueifnull: any = vbaMissing) => {
        if (val === vbaNull || val === vbaEmpty || val === null || val === undefined) {
            return valueifnull === vbaMissing ? '' : valueifnull;
        }
        return val;
    }, [{ name: 'Value' }, { name: 'ValueIfNull', optional: true }]);

    const ptrFn = () => ctx.ptrNext();
    ctx.reg('varptr', ptrFn, [{ name: 'VarName' }]);
    ctx.reg('strptr', ptrFn, [{ name: 'VarName' }]);
    ctx.reg('objptr', ptrFn, [{ name: 'VarName' }]);
    ctx.reg('createobject', (id: string, _serverName?: string) => ctx.createObj(id), [
        { name: 'Class' },
        { name: 'ServerName', optional: true },
    ]);
    const getObjectFunc = (pathname?: string, classId?: string) => {
        if (pathname) {
            return { __vbaTypeName__: 'Object', Path: pathname, Name: pathname.split(/[\\\/]/).pop() };
        }
        if (classId) return ctx.createObj(classId);
        return vbaNothing;
    };
    ctx.reg('getobject', getObjectFunc, [
        { name: 'PathName', optional: true },
        { name: 'Class', optional: true },
    ]);
    ctx.reg('iif', (c: any, t: any, f: any) => ctx.isTrue(c) ? t : f, [
        { name: 'Expr' }, { name: 'TruePart' }, { name: 'FalsePart' },
    ]);
    ctx.reg('choose', (i: any, ...c: any[]) => {
        if (i === vbaNull) return vbaNull;
        const idx = Math.floor(Number(i));
        return (idx >= 1 && idx <= c.length) ? c[idx - 1] : vbaNull;
    }, [
        { name: 'Index' },
        { name: 'Choice', isParamArray: true },
    ]);
    ctx.reg('switch', (...args: any[]) => {
        if (args.length % 2 !== 0) ctx.throwError(VbaErrorCode.INVALID_PROCEDURE_CALL, 'Invalid procedure call or argument');
        for (let i = 0; i < args.length; i += 2) if (ctx.isTrue(args[i])) return args[i + 1];
        return vbaNull;
    }, [
        { name: 'VarExpr', isParamArray: true },
    ]);
    ctx.reg('array', (...args: any[]) => {
        // Pre-fill base slots so element i maps to JS index (base + i), matching Option Base.
        const base = ctx.arrayBase;
        const a: any[] = new Array(base).fill(undefined).concat(args);
        (a as any).vbaBase = base;
        return a;
    }, [
        { name: 'Arglist', isParamArray: true },
    ]);
    ctx.reg('lbound', (a: any, dim: any = 1) => {
        if (!Array.isArray(a)) ctx.throwError(VbaErrorCode.SUBSCRIPT_OUT_OF_RANGE, "Subscript out of range");
        const dimIndex = Number(dim) - 1;
        if ((a as any).__vbaDimensions__) {
            if (dimIndex < 0 || dimIndex >= (a as any).__vbaDimensions__.length) {
                ctx.throwError(VbaErrorCode.SUBSCRIPT_OUT_OF_RANGE, "Subscript out of range");
            }
            return (a as any).__vbaDimensions__[dimIndex].lower;
        }
        if (dimIndex > 0) ctx.throwError(VbaErrorCode.SUBSCRIPT_OUT_OF_RANGE, "Subscript out of range");
        return (a as any).vbaBase || 0;
    }, [{ name: 'ArrayName' }, { name: 'Dimension', optional: true }]);
    ctx.reg('ubound', (a: any, dim: any = 1) => {
        if (!Array.isArray(a)) ctx.throwError(VbaErrorCode.SUBSCRIPT_OUT_OF_RANGE, "Subscript out of range");
        const dimIndex = Number(dim) - 1;
        if ((a as any).__vbaDimensions__) {
            if (dimIndex < 0 || dimIndex >= (a as any).__vbaDimensions__.length) {
                ctx.throwError(VbaErrorCode.SUBSCRIPT_OUT_OF_RANGE, "Subscript out of range");
            }
            return (a as any).__vbaDimensions__[dimIndex].upper;
        }
        if (dimIndex > 0) ctx.throwError(VbaErrorCode.SUBSCRIPT_OUT_OF_RANGE, "Subscript out of range");
        // a.length already includes vbaBase filler slots (added by Array() for Option Base 1),
        // so UBound = a.length - 1 (not vbaBase + a.length - 1).
        return a.length - 1;
    }, [{ name: 'ArrayName' }, { name: 'Dimension', optional: true }]);
}

// ---------------------------------------------------------------------------
// Registry functions — SaveSetting, GetSetting, GetAllSettings, DeleteSetting
// ---------------------------------------------------------------------------

export function registerRegistryFunctions(ctx: StdlibCtx): void {
    ctx.reg('savesetting', (app: string, sec: string, key: string, val: any) => {
        if (!ctx.registry[app]) ctx.registry[app] = {};
        if (!ctx.registry[app][sec]) ctx.registry[app][sec] = {};
        ctx.registry[app][sec][key] = String(val);
    }, [
        { name: 'AppName' }, { name: 'Section' }, { name: 'Key' }, { name: 'Setting' },
    ]);
    ctx.reg('getsetting', (app: string, sec: string, key: string, def: any = "") => {
        return (ctx.registry[app]?.[sec]?.[key]) ?? String(def);
    }, [
        { name: 'AppName' }, { name: 'Section' }, { name: 'Key' }, { name: 'Default', optional: true },
    ]);
    ctx.reg('getallsettings', (app: string, sec: string) => {
        const s = ctx.registry[app]?.[sec];
        if (!s) return vbaEmpty;
        const res = Object.entries(s).map(([k, v]) => [k, v]);
        (res as any).vbaBase = 0;
        return res;
    }, [{ name: 'AppName' }, { name: 'Section' }]);
    ctx.reg('deletesetting', (app: string, sec?: string, key?: string) => {
        if (!sec) delete ctx.registry[app];
        else if (!key) delete ctx.registry[app]?.[sec];
        else delete ctx.registry[app]?.[sec]?.[key];
    }, [
        { name: 'AppName' },
        { name: 'Section', optional: true },
        { name: 'Key', optional: true },
    ]);
}
