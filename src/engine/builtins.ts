import {
    VbaBoolean, VbaDate, VbaDecimal, VbaCurrency, VbaErrorValue,
    vbaEmpty, vbaNull, vbaNothing, vbaMissing, vbaTrue, vbaFalse,
    toVbaDate, fromVbaDate, parseVbaDate, tryParseTimeFractionString,
    parseCurrencyString, bankersDivide,
    isVbaObjectReferenceCompatible,
} from './vba-types';
import { VbaErrorCode } from './vba-errors';
import { vbaToBoolean, vbaToString, vbaRound, unwrapDefaultValue, normalizeVbaNumericString } from './coerce';
// VbaErrorCode is imported as a value-namespace for use in function bodies (VbaErrorCode.OVERFLOW etc.)
import type { ProcedureDeclaration } from './parser';
import { findClassProperty } from './property-resolution';
import {
    formatDate, formatNumber, formatString, formatNullSection,
    splitFormatSections, containsDateFormatTokens, hasNumericFormatTokens, isNumericOnlyFormat,
    hasMixedNumericStringTokens,
    stripFormatColorDirectives, stripVbaBracketDirectives, hasUnclosedVbaBracket, hasUnescapedFormatChars,
} from './format';
import { vbaWeekNumber } from './date-week';
import { classifyArgumentCount } from './argument-contract';

/** Expand only objects that explicitly expose a VBA-style default Value. */
function unwrapVbaDefaultValue(value: any): any {
    return unwrapDefaultValue(value);
}

/** Reject Null at host-function boundaries that require a concrete argument. */
function rejectNullArgument(ctx: Pick<StdlibCtx, 'throwError'>, value: any): void {
    if (value === vbaNull) {
        ctx.throwError(VbaErrorCode.INVALID_USE_OF_NULL, 'Invalid use of Null');
    }
}

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
    /** The VBA declaration requires an array/UDT argument at compile time. */
    isArray?: boolean;
    /** Apply a shared VBA coercion before invoking the implementation. */
    coerce?: 'boolean' | 'string' | 'long';
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
    const isNumericValue = (val: any, seen = new Set<any>()): VbaBoolean => {
        if (val === vbaNull) return vbaFalse;
        if (val === vbaEmpty || val === undefined) return vbaTrue;
        if (val && typeof val === 'object' && val.__vbaDefault__ === true) {
            if (seen.has(val)) return vbaFalse;
            seen.add(val);
            return isNumericValue(val.Value, seen);
        }
        if (typeof val === 'number' || typeof val === 'bigint' || val instanceof VbaDecimal || val instanceof VbaCurrency || val instanceof VbaBoolean) return vbaTrue;
        if (typeof val === 'string') {
            const s = val.trim();
            if (s === "") return vbaFalse;
            const cleaned = normalizeVbaNumericString(s).replace(/[$,]/g, '');
            // 判定が目的なので throw しない。toVbaNumber(parseFloat) は "2024/01/15" の先頭
            // 数値を受理してしまうため、ここでは全体が数値の場合のみ通す厳密な Number() を使う
            if (/^&[hH][0-9a-fA-F]+$/.test(cleaned) || /^&[oO][0-7]+$/.test(cleaned)) return vbaTrue;
            const n = Number(cleaned);
            return (!isNaN(n) && isFinite(n)) ? vbaTrue : vbaFalse;
        }
        return vbaFalse;
    };
    ctx.reg('isnumeric', isNumericValue, [{ name: 'Expression' }]);
    ctx.reg('isdate', (val: any) => {
        val = unwrapVbaDefaultValue(val);
        if (val instanceof VbaDate) return vbaTrue;
        if (val instanceof VbaDecimal) return isFinite(val.value) ? vbaTrue : vbaFalse;
        if (val instanceof VbaCurrency) return isFinite(Number(val.internal)) ? vbaTrue : vbaFalse;
        if (typeof val === 'number') return isFinite(val) ? vbaTrue : vbaFalse;
        if (typeof val === 'string') {
            const d = Date.parse(val);
            return !isNaN(d) ? vbaTrue : vbaFalse;
        }
        return vbaFalse;
    }, [{ name: 'Expression' }]);
    ctx.reg('isobject', (val: any) => {
        return isVbaObjectReferenceCompatible(val) ? vbaTrue : vbaFalse;
    }, [{ name: 'Identifier' }]);
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
        if (Array.isArray(val)) {
            const elementType = (val as any).__vbaElementType__ as string | undefined;
            const varTypes: Record<string, number> = {
                byte: 17, integer: 2, long: 3, single: 4, double: 5,
                currency: 6, date: 7, string: 8, boolean: 11,
                longlong: 20, longptr: 20,
            };
            if (elementType && varTypes[elementType] !== undefined) return 8192 + varTypes[elementType];
            if ((val as any).__vbaElementTypeName__) return 8192 + 36;
            if ((val as any).__vbaElementObjectTypeName__) return 8192 + 9;
            return 8192 + 12;
        }
        if (val instanceof VbaCurrency) return 6;
        if (typeof val === 'number') return 5;
        if (typeof val === 'string') return 8;
        if (val instanceof VbaDecimal) return 14;
        if (typeof val === 'bigint') return 20;
        if (val && (val.__vbaClass__ || val.__isVbaDict__ || val.__isVbaCollection__)) return 9;
        if (val && val.__vbaUdt__ === true) return 36;
        if (val && val.__vbaTypeName__) return 9;
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
        if (Array.isArray(val)) {
            const elementType = (val as any).__vbaElementType__ as string | undefined;
            const typeNames: Record<string, string> = {
                byte: 'Byte', integer: 'Integer', long: 'Long', single: 'Single', double: 'Double',
                currency: 'Currency', date: 'Date', string: 'String', boolean: 'Boolean',
                longlong: 'LongLong', longptr: 'LongLong',
            };
            if (elementType && typeNames[elementType]) return `${typeNames[elementType]}()`;
            if ((val as any).__vbaElementTypeName__) return `${(val as any).__vbaElementTypeName__}()`;
            if ((val as any).__vbaElementObjectTypeName__) return `${(val as any).__vbaElementObjectTypeName__}()`;
            return 'Variant()';
        }
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
    ctx.reg('callbyname', (obj: any, procName: any, callType: any, ...args: any[]) => {
        if (obj === null || obj === undefined || obj === vbaNothing) {
            ctx.throwError(VbaErrorCode.OBJECT_VARIABLE_NOT_SET, 'Object variable or With block variable not set');
        }
        if (procName === vbaNull) ctx.throwError(VbaErrorCode.INVALID_USE_OF_NULL, 'Invalid use of Null');
        if (callType === vbaNull) ctx.throwError(VbaErrorCode.INVALID_USE_OF_NULL, 'Invalid use of Null');
        const name = vbaToString(procName).toLowerCase();
        const validateDynamicArguments = (
            member: any,
            supplied: any[],
            declaredSpec?: BuiltinParamSpec[],
        ): void => {
            const spec = declaredSpec ?? member?.__vbaParamSpec__ as BuiltinParamSpec[] | undefined;
            // Dynamic host members without metadata still expose their JS arity.
            // Convert that fallback into the same contract representation instead
            // of maintaining a second required/excess implementation here.
            const parameters = spec ?? Array.from(
                { length: typeof member?.length === 'number' ? member.length : 0 },
                () => ({ optional: false }),
            );
            if (classifyArgumentCount(parameters, supplied.length) !== null) {
                ctx.throwError(VbaErrorCode.WRONG_NUMBER_OF_ARGUMENTS,
                    'Wrong number of arguments or invalid property assignment');
            }
        };
        const findExternalKey = (): string | undefined => {
            let current = obj;
            while (current && current !== Object.prototype) {
                const key = Reflect.ownKeys(current).find(
                    k => typeof k === 'string' && k.toLowerCase() === name
                );
                if (typeof key === 'string') return key;
                current = Object.getPrototypeOf(current);
            }
            return undefined;
        };
        const findExternalParamSpec = (match: string): BuiltinParamSpec[] | undefined => {
            const specs = (obj as any).__vbaParamSpecs__ as Record<string, BuiltinParamSpec[]> | undefined;
            if (!specs) return undefined;
            const key = Object.keys(specs).find(k => k.toLowerCase() === match.toLowerCase());
            return key ? specs[key] : undefined;
        };
        if (callType === 2 /* VbGet */ || callType === 1 /* VbMethod */) {
            if (obj.__vbaClass__) {
                const classDef = obj.__classDef__ as ProcedureDeclaration & { procedures: ProcedureDeclaration[] };
                const getter = findClassProperty(classDef.procedures, name, 'get');
                if (getter) return ctx.callMethod(obj, getter, args);
                const method = (classDef as any).procedures.find(
                    (p: any) => !p.isProperty && p.name.name.toLowerCase() === name
                );
                if (method) return ctx.callMethod(obj, method, args);
                return obj.__instanceEnv__.get(name);
            }
            if (typeof obj === 'object' && obj !== null) {
                const match = findExternalKey();
                if (match === undefined) {
                    ctx.throwError(VbaErrorCode.OBJECT_DOESNT_SUPPORT_PROPERTY, `Object doesn't support this property or method: '${procName}'`);
                }
                const val = obj[match];
                if (typeof val === 'function') {
                    validateDynamicArguments(val, args, findExternalParamSpec(match));
                    return val.apply(obj, args);
                }
                return val;
            }
        }
        if (callType === 4 /* VbLet */ || callType === 8 /* VbSet */) {
            const propType = callType === 4 ? 'let' : 'set';
            if (callType === 8 && args.length > 0 && !isVbaObjectReferenceCompatible(args[args.length - 1])) {
                ctx.throwError(VbaErrorCode.OBJECT_REQUIRED, 'Object required');
            }
            if (obj.__vbaClass__) {
                const classDef = obj.__classDef__ as ProcedureDeclaration & { procedures: ProcedureDeclaration[] };
                const setter = findClassProperty(classDef.procedures, name, propType);
                if (setter) return ctx.callMethod(obj, setter, args);
                const oppositeType = propType === 'let' ? 'set' : 'let';
                if (findClassProperty(classDef.procedures, name, oppositeType)) {
                    ctx.throwError(VbaErrorCode.OBJECT_DOESNT_SUPPORT_PROPERTY, `Object doesn't support this property or method: '${procName}'`);
                }
                obj.__instanceEnv__.set(name, args[0]);
                return;
            }
            if (typeof obj === 'object' && obj !== null) {
                const match = findExternalKey();
                if (match === undefined) {
                    ctx.throwError(VbaErrorCode.OBJECT_DOESNT_SUPPORT_PROPERTY, `Object doesn't support this property or method: '${procName}'`);
                }
                const val = obj[match];
                if (typeof val === 'function') {
                    validateDynamicArguments(val, args, findExternalParamSpec(match));
                    val.apply(obj, args);
                } else { obj[match] = args[0]; }
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
    const unwrapConversionValue = (value: any): any => {
        const unwrapped = unwrapVbaDefaultValue(value);
        if (unwrapped === vbaNull) {
            ctx.throwError(VbaErrorCode.INVALID_USE_OF_NULL, 'Invalid use of Null');
        }
        return unwrapped;
    };
    const parseRadixForWidth = (source: string, bits: 16 | 32 | 64): number | bigint | undefined => {
        const text = normalizeVbaNumericString(source.trim());
        const match = /^([+-]?)&([hHoO])([0-9a-fA-F]+)$/.exec(text);
        if (!match) return undefined;
        const base = match[2].toLowerCase() === 'h' ? 16 : 8;
        const digits = match[3];
        if (base === 16 && !/^[0-9a-fA-F]+$/.test(digits)) return undefined;
        if (base === 8 && !/^[0-7]+$/.test(digits)) return undefined;
        const magnitude = BigInt(base === 16 ? `0x${digits}` : `0o${digits}`);
        const fullDigits = base === 16 ? bits / 4 : Math.ceil(bits / 3);
        const signBit = 1n << BigInt(bits - 1);
        const modulus = 1n << BigInt(bits);
        let value = magnitude;
        if (match[1] === '-') value = -magnitude;
        // Only an in-width bit pattern is sign-extended.  A value with more
        // than the target width is a positive magnitude and must reach the
        // conversion's overflow/range check rather than wrap modulo 2^bits.
        else if (magnitude < modulus && digits.length >= fullDigits && magnitude >= signBit) {
            value = magnitude - modulus;
        }
        return bits === 64 ? value : Number(value);
    };
    const radixInputBits = (source: string): 16 | 32 | 64 | undefined => {
        const text = normalizeVbaNumericString(source.trim());
        const match = /^([+-]?)&([hHoO])([0-9a-fA-F]+)$/.exec(text);
        if (!match) return undefined;
        const digits = match[3];
        const base = match[2].toLowerCase();
        return base === 'h'
            ? (digits.length <= 4 ? 16 : digits.length <= 8 ? 32 : digits.length <= 16 ? 64 : undefined)
            : (digits.length <= 6 ? 16 : digits.length <= 11 ? 32 : digits.length <= 22 ? 64 : undefined);
    };
    const hasFullRadixWidth = (source: string, bits: 16 | 32 | 64): boolean => {
        const text = normalizeVbaNumericString(source.trim());
        const match = /^([+-]?)&([hHoO])([0-9a-fA-F]+)$/.exec(text);
        if (!match) return false;
        const digits = match[3];
        return digits.length === (match[2].toLowerCase() === 'h' ? bits / 4 : Math.ceil(bits / 3));
    };
    const parseRadixForTarget = (source: string, targetBits: 16 | 32 | 64): number | bigint | undefined => {
        const inputBits = radixInputBits(source);
        if (inputBits === undefined) return undefined;
        const parsed = parseRadixForWidth(source, inputBits);
        if (parsed === undefined || inputBits >= targetBits || /^[+-]/.test(source.trim())) return parsed;
        const text = normalizeVbaNumericString(source.trim());
        const match = /^([+-]?)&([hHoO])([0-9a-fA-F]+)$/.exec(text);
        if (!match) return parsed;
        const magnitude = BigInt(match[2].toLowerCase() === 'h'
            ? `0x${match[3]}`
            : `0o${match[3]}`);
        return targetBits === 64 ? magnitude : Number(magnitude);
    };
    const parseRadixForValue = (
        source: string,
        options: { signExtendShort?: boolean } = {},
    ): number | bigint | undefined => {
        const text = normalizeVbaNumericString(source.trim());
        const match = /^([+-]?)&([hHoO])([0-9a-fA-F]+)$/.exec(text);
        if (!match) return undefined;
        const digits = match[3];
        const base = match[2].toLowerCase();
        const bits = base === 'h'
            ? (digits.length <= 4 ? 16 : digits.length <= 8 ? 32 : 64)
            : (digits.length <= 6 ? 16 : digits.length <= 11 ? 32 : 64);
        const parsed = parseRadixForWidth(text, bits as 16 | 32 | 64);
        if (options.signExtendShort !== false || parsed === undefined) return parsed;
        // CDec/CCur treat a 16-bit radix pattern as an unsigned magnitude.
        // Sign extension starts at the 32-bit pattern width for these
        // conversion targets.
        if (/^[+-]/.test(text)) return parsed;
        const fullDigits = base === 'h' ? bits / 4 : Math.ceil(bits / 3);
        if (bits !== 16 && digits.length >= fullDigits) return parsed;
        const magnitude = BigInt(base === 'h' ? `0x${digits}` : `0o${digits}`);
        return bits === 64 ? magnitude : Number(magnitude);
    };
    ctx.reg('cbyte', (val: any) => {
        val = unwrapConversionValue(val);
        if (val instanceof VbaBoolean) return val.valueOf() ? 255 : 0;
        if (typeof val === 'string' && hasFullRadixWidth(val, 64)) {
            ctx.throwError(VbaErrorCode.TYPE_MISMATCH, "Type mismatch");
        }
        const n = ctx.round(ctx.toVbaNumber(val));
        if (n < 0 || n > 255) ctx.throwError(VbaErrorCode.OVERFLOW, "Overflow");
        return n;
    }, [{ name: 'Expression' }]);
    ctx.reg('cint', (val: any) => {
        val = unwrapConversionValue(val);
        if (typeof val === 'string' && /^[+-]&[hHoO]/.test(val.trim())) {
            ctx.throwError(VbaErrorCode.TYPE_MISMATCH, "Type mismatch");
        }
        const inputBits = typeof val === 'string' ? radixInputBits(val) : undefined;
        if (inputBits === 64 && hasFullRadixWidth(val, 64)) ctx.throwError(VbaErrorCode.TYPE_MISMATCH, "Type mismatch");
        const radix = typeof val === 'string' && inputBits !== undefined
            ? parseRadixForWidth(val, inputBits)
            : undefined;
        const n = ctx.round(radix === undefined ? ctx.toVbaNumber(val) : Number(radix));
        if (n < -32768 || n > 32767) ctx.throwError(VbaErrorCode.OVERFLOW, "Overflow");
        return n;
    }, [{ name: 'Expression' }]);
    ctx.reg('clng', (val: any) => {
        val = unwrapConversionValue(val);
        if (typeof val === 'string' && /^[+-]&[hHoO]/.test(val.trim())) {
            ctx.throwError(VbaErrorCode.TYPE_MISMATCH, "Type mismatch");
        }
        const inputBits = typeof val === 'string' ? radixInputBits(val) : undefined;
        if (inputBits === 64 && hasFullRadixWidth(val, 64)) ctx.throwError(VbaErrorCode.TYPE_MISMATCH, "Type mismatch");
        const radix = typeof val === 'string' && inputBits !== undefined
            ? parseRadixForTarget(val, 32)
            : undefined;
        const n = ctx.round(radix === undefined ? ctx.toVbaNumber(val) : Number(radix));
        if (n < -2147483648 || n > 2147483647) ctx.throwError(VbaErrorCode.OVERFLOW, "Overflow");
        return n;
    }, [{ name: 'Expression' }]);
    ctx.reg('csng', (val: any) => {
        val = unwrapConversionValue(val);
        const n = ctx.toVbaNumber(val);
        const f32 = Math.fround(n);
        if (!isFinite(f32) && isFinite(n)) ctx.throwError(VbaErrorCode.OVERFLOW, "Overflow");
        return f32;
    }, [{ name: 'Expression' }]);
    ctx.reg('cdbl', (val: any) => {
        val = unwrapConversionValue(val);
        return ctx.toVbaNumber(val);
    }, [{ name: 'Expression' }]);
    ctx.reg('cdate', (val: any) => {
        val = unwrapConversionValue(val);
        // CDate(Empty) はシリアル値 0（1899-12-30）を返す（実 VBA 差分で裁定）
        if (val === vbaEmpty) return new VbaDate(0);
        if (val instanceof VbaDate) return val;
        if (typeof val === 'string') {
            const trimmed = val.trim();
            // "H.N" 形式は「H 時 N 分」の時刻として解釈される（シリアル値としては解釈
            // されない。実 VBA 差分で裁定: CDate("2.5") = 02:05:00。CDate(2.5)（文字列
            // でない数値そのもの）は通常どおりシリアル値 2.5 = 正午なので非対称）
            const timeFrac = tryParseTimeFractionString(trimmed);
            if (timeFrac !== undefined) {
                return new VbaDate(timeFrac);
            }
            // "M,Y" 形式（カンマ 1 つのみの単純な数値）は「M 月 Y 年 1 日」として解釈される
            // （実 VBA 差分で裁定: CDate("1,234") = 西暦234年1月1日）
            const myMatch = /^(\d{1,2}),(\d+)$/.exec(trimmed);
            if (myMatch) {
                const mo = Number(myMatch[1]), yr = Number(myMatch[2]);
                if (mo >= 1 && mo <= 12) {
                    return new VbaDate(toVbaDate(new Date(yr, mo - 1, 1)));
                }
            }
            // 数値文字列（10進/16進/8進/指数/前後空白）はシリアル値として解釈し、
            // それ以外は日付文字列として解釈する（実 VBA 差分で裁定: CDate("&H10") は数値扱い）
            let numericSerial: number | undefined;
            try { numericSerial = ctx.toVbaNumber(trimmed); } catch { /* 日付文字列として続行 */ }
            if (numericSerial === undefined) {
                return new VbaDate(toVbaDate(parseVbaDate(val)));
            }
            if (numericSerial < -657434 || numericSerial >= 2958466) ctx.throwError(VbaErrorCode.OVERFLOW, 'Overflow');
            return new VbaDate(numericSerial);
        }
        const serial = ctx.toVbaNumber(val);
        // 日付シリアル値の有効範囲（100-01-01 〜 9999-12-31）。範囲外は Error 6（実 VBA 差分で裁定）
        if (serial < -657434 || serial >= 2958466) ctx.throwError(VbaErrorCode.OVERFLOW, 'Overflow');
        return new VbaDate(serial);
    }, [{ name: 'Expression' }]);
    ctx.reg('cvdate', (val: any) => {
        if (val === vbaNull) return vbaNull;
        return (ctx.envGet('cdate') as Function)(val);
    }, [{ name: 'Expression' }]);
    const parseCDecString = (source: string): VbaDecimal => {
        const text = normalizeVbaNumericString(source.trim());
        if (/^[+-]&[hHoO]/.test(text)) {
            ctx.throwError(VbaErrorCode.TYPE_MISMATCH, 'Type mismatch');
        }
        const radix = parseRadixForValue(text, { signExtendShort: false });
        if (radix !== undefined) {
            const integer = typeof radix === 'bigint' ? radix : BigInt(radix);
            return new VbaDecimal(integer, 0);
        }
        const grouped = /^[+-]?(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d*)?(?:[eE][+-]?\d+)?$/.test(text)
            || /^[+-]?\.\d+(?:[eE][+-]?\d+)?$/.test(text);
        if (grouped) return VbaDecimal.fromString(text.replace(/,/g, '').replace(/^\+/, ''));
        return VbaDecimal.fromString(text);
    };
    ctx.reg('cdec', (val: any) => {
        val = unwrapConversionValue(val);
        if (val instanceof VbaDecimal) return val;
        if (val instanceof VbaCurrency) return new VbaDecimal(val.internal, 4);
        if (val instanceof VbaBoolean) return new VbaDecimal(BigInt(val.value), 0);
        if (typeof val === 'bigint') return new VbaDecimal(val, 0);
        if (typeof val === 'string') return parseCDecString(val);
        return VbaDecimal.fromNumber(ctx.toVbaNumber(val));
    }, [{ name: 'Expression' }]);
    ctx.reg('ccur', (val: any) => {
        val = unwrapConversionValue(val);
        if (val instanceof VbaCurrency) return val;
        if (val instanceof VbaDecimal) {
            const scaled = bankersDivide(
                val.mantissa * 10000n,
                10n ** BigInt(val.scale),
            );
            if (scaled < -9223372036854775808n || scaled > 9223372036854775807n) {
                ctx.throwError(VbaErrorCode.OVERFLOW, "Overflow");
            }
            return new VbaCurrency(scaled);
        }
        if (val instanceof VbaBoolean) return new VbaCurrency(BigInt(val.value) * 10000n);
        if (typeof val === 'bigint') return new VbaCurrency(val * 10000n);
        if (typeof val === 'string') {
            const trimmed = val.trim();
            const decimalString = trimmed.replace(/,/g, '');
            if (/^\+?(\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?$/.test(decimalString) ||
                /^-?(\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?$/.test(decimalString)) {
                return parseCurrencyString(trimmed);
            }
            // &H/&O/指数/カンマ区切り等は数値文字列として解釈（実 VBA 差分で裁定）
            if (radixInputBits(trimmed) === 64) {
                if (/^[+-]?&[oO]/.test(trimmed)
                    && trimmed.replace(/^[+-]?&[oO]/, '').length >= 22) {
                    const digits = trimmed.replace(/^[+-]?&[oO]/, '');
                    if (BigInt(`0o${digits}`) < (1n << 64n)) {
                        ctx.throwError(VbaErrorCode.TYPE_MISMATCH, "Type mismatch");
                    }
                }
                ctx.throwError(VbaErrorCode.OVERFLOW, "Overflow");
            }
            const radix = parseRadixForValue(trimmed, { signExtendShort: false });
            if (radix !== undefined) {
                const integer = typeof radix === 'bigint' ? radix : BigInt(radix);
                return new VbaCurrency(integer * 10000n);
            }
            return VbaCurrency.fromNumber(ctx.toVbaNumber(trimmed));
        }
        return VbaCurrency.fromNumber(ctx.toVbaNumber(val));
    }, [{ name: 'Expression' }]);
    const clnglngFunc = (val: any) => {
        val = unwrapConversionValue(val);
        if (val === null) ctx.throwError(VbaErrorCode.TYPE_MISMATCH, "Type mismatch");
        if (typeof val === 'bigint') {
            if (val < -9223372036854775808n || val > 9223372036854775807n) {
                ctx.throwError(VbaErrorCode.OVERFLOW, "Overflow");
            }
            return val;
        }
        if (val instanceof VbaDecimal) {
            const n = bankersDivide(val.mantissa, 10n ** BigInt(val.scale));
            if (n < -9223372036854775808n || n > 9223372036854775807n) {
                ctx.throwError(VbaErrorCode.OVERFLOW, "Overflow");
            }
            return n;
        }
        if (val instanceof VbaCurrency) {
            const n = bankersDivide(val.internal, 10000n);
            if (n < -9223372036854775808n || n > 9223372036854775807n) {
                ctx.throwError(VbaErrorCode.OVERFLOW, "Overflow");
            }
            return n;
        }
        if (typeof val === 'string') {
            const trimmed = val.trim();
            const inputBits = radixInputBits(trimmed);
            if (inputBits === 64 && /^&[oO]/.test(trimmed)
                && hasFullRadixWidth(trimmed, 64)) {
                const digits = trimmed.replace(/^[+-]?&[oO]/, '');
                const magnitude = BigInt(`0o${digits}`);
                if (magnitude < (1n << 64n)) {
                    ctx.throwError(VbaErrorCode.TYPE_MISMATCH, "Type mismatch");
                }
            }
            const radix = inputBits === undefined ? undefined : parseRadixForTarget(trimmed, 64);
            if (radix !== undefined) {
                const n = typeof radix === 'bigint' ? radix : BigInt(radix);
                if (n < -9223372036854775808n || n > 9223372036854775807n) {
                    ctx.throwError(VbaErrorCode.OVERFLOW, "Overflow");
                }
                return n;
            }
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
        // 実 VBA 差分で裁定: CStr(Null) は Error 94（'' ではない）
        if (val === vbaNull) ctx.throwError(VbaErrorCode.INVALID_USE_OF_NULL, 'Invalid use of Null');
        try { return vbaToString(val); } catch (e: any) {
            if (e?.type === 'VbaError') ctx.throwError(e.number, e.message);
            throw e;
        }
    }, [{ name: 'Expression' }]);
    ctx.reg('cbool', (val: any) => {
        val = unwrapConversionValue(val);
        return vbaToBoolean(val);
    }, [{ name: 'Expression' }]);
    ctx.reg('cvar', (val: any) => val, [{ name: 'Expression' }]);
    ctx.reg('cverr', (val: any) => {
        if (val instanceof VbaErrorValue) return val;
        if (val === vbaNull) ctx.throwError(VbaErrorCode.INVALID_USE_OF_NULL, 'Invalid use of Null');
        const code = ctx.round(ctx.toVbaNumber(val));
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
        if (cleaned.toLowerCase().startsWith('&h')) {
            const hexDigits = cleaned.slice(2);
            // Val parses over-wide hexadecimal text through the 32-bit
            // integer path.  Keep the lower 32-bit pattern instead of
            // applying the signed 64-bit conversion used by C* functions.
            if (/^[0-9a-fA-F]{16}$/.test(hexDigits)) {
                const low = BigInt(`0x${hexDigits}`) & 0xFFFFFFFFn;
                const signed = low >= 0x80000000n ? low - 0x100000000n : low;
                return Number(signed);
            }
            const radix = parseRadixForValue(cleaned);
            return radix === undefined ? 0 : Number(radix);
        }
        if (cleaned.toLowerCase().startsWith('&o')) {
            const radix = parseRadixForValue(cleaned);
            return radix === undefined ? 0 : Number(radix);
        }
        const match = cleaned.match(/^[+-]?\d*(\.\d*)?([eE][+-]?\d+)?/);
        const value = match ? parseFloat(match[0]) || 0 : 0;
        // Val recognizes legacy type-declaration suffixes and applies their
        // conversion range before returning the number.
        if (match) {
            const suffix = cleaned[match[0].length];
            if (suffix === '%' || suffix === '&' || suffix === '^') {
                if (!Number.isInteger(value)) ctx.throwError(VbaErrorCode.TYPE_MISMATCH, 'Type mismatch');
                const min = suffix === '%' ? -32768 : suffix === '&' ? -2147483648 : -9223372036854776000;
                const max = suffix === '%' ? 32767 : suffix === '&' ? 2147483647 : 9223372036854776000;
                if (value < min || value > max) ctx.throwError(VbaErrorCode.OVERFLOW, 'Overflow');
            }
            if (suffix === '!') {
                if (!Number.isFinite(value) || Math.abs(value) > 3.4028234663852886e38) {
                    ctx.throwError(VbaErrorCode.OVERFLOW, 'Overflow');
                }
                return Math.fround(value);
            }
            if (suffix === '@') return VbaCurrency.fromNumber(value);
            if (suffix === '#' && !Number.isFinite(value)) {
                ctx.throwError(VbaErrorCode.OVERFLOW, 'Overflow');
            }
        }
        if (!Number.isFinite(value)) ctx.throwError(VbaErrorCode.OVERFLOW, 'Overflow');
        return value;
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
    ctx.reg('int', (val: any) => {
        if (val === vbaNull) return vbaNull;
        if (val instanceof VbaDate) return new VbaDate(Math.floor(val.value));  // Date 型を保存（実 VBA 差分で裁定）
        return Math.floor(ctx.toVbaNumber(val));
    }, [{ name: 'Number' }]);
    ctx.reg('fix', (val: any) => {
        if (val === vbaNull) return vbaNull;
        if (val instanceof VbaDate) {  // Date 型を保存（実 VBA 差分で裁定）
            return new VbaDate(val.value >= 0 ? Math.floor(val.value) : Math.ceil(val.value));
        }
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
        rejectNullArgument(ctx, digits);
        const digitCount = ctx.round(ctx.toVbaNumber(digits), 0);
        return ctx.round(ctx.toVbaNumber(val), digitCount);
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
    const normalizeRandomArg = (val?: any): number | undefined => {
        if (val === undefined) return undefined;
        if (val === vbaNull) ctx.throwError(VbaErrorCode.INVALID_USE_OF_NULL, 'Invalid use of Null');
        return ctx.toVbaNumber(val);
    };
    const rndFunc = (val?: any) => {
        if (!rndInitialized) { rndSeed = 0.5; rndInitialized = true; }
        const n = normalizeRandomArg(val);
        if (n === undefined || n > 0) {
            rndSeed = (rndSeed * 214013 + 2531011) % 4294967296;
            lastRnd = rndSeed / 4294967296;
            return lastRnd;
        } else if (n === 0) {
            return lastRnd;
        } else if (n < 0) {
            const s = Math.abs(n) * 9301 + 49297;
            if (!Number.isFinite(s)) ctx.throwError(VbaErrorCode.OVERFLOW, 'Overflow');
            lastRnd = (s % 233280) / 233280;
            return lastRnd;
        }
        return lastRnd;
    };
    ctx.reg('rnd', rndFunc, [{ name: 'Number', optional: true }]);
    ctx.reg('randomize', (val?: any) => {
        const n = normalizeRandomArg(val);
        rndInitialized = true;
        const seed = (n === undefined)
            ? (Date.now() % 4294967296)
            : (Math.round(Math.abs(n) * 1000) % 4294967296);
        if (!Number.isFinite(seed)) ctx.throwError(VbaErrorCode.OVERFLOW, 'Overflow');
        rndSeed = seed;
        lastRnd = rndSeed / 4294967296;
    }, [{ name: 'Number', optional: true }]);
}

// ---------------------------------------------------------------------------
// String functions — Asc, Chr, InStr, LCase, Left, Len, Mid, Right, etc.
// ---------------------------------------------------------------------------

// StrConv vbNarrow(8) 用: 全角カタカナ → 半角カナ。濁点/半濁点付き文字は
// 基底文字 + 結合文字（ﾞ/ﾟ）の2文字表現になる（実 VBA 差分で裁定・実装）。
// 逆方向（半角→全角）は String.prototype.normalize('NFKC') で正確に変換できる。
const FULL_TO_HALF_KANA: Record<string, string> = {
    '。': '｡', '「': '｢', '」': '｣', '、': '､', '・': '･',
    'ヲ': 'ｦ', 'ァ': 'ｧ', 'ィ': 'ｨ', 'ゥ': 'ｩ', 'ェ': 'ｪ', 'ォ': 'ｫ',
    'ャ': 'ｬ', 'ュ': 'ｭ', 'ョ': 'ｮ', 'ッ': 'ｯ', 'ー': 'ｰ',
    'ア': 'ｱ', 'イ': 'ｲ', 'ウ': 'ｳ', 'エ': 'ｴ', 'オ': 'ｵ',
    'カ': 'ｶ', 'キ': 'ｷ', 'ク': 'ｸ', 'ケ': 'ｹ', 'コ': 'ｺ',
    'サ': 'ｻ', 'シ': 'ｼ', 'ス': 'ｽ', 'セ': 'ｾ', 'ソ': 'ｿ',
    'タ': 'ﾀ', 'チ': 'ﾁ', 'ツ': 'ﾂ', 'テ': 'ﾃ', 'ト': 'ﾄ',
    'ナ': 'ﾅ', 'ニ': 'ﾆ', 'ヌ': 'ﾇ', 'ネ': 'ﾈ', 'ノ': 'ﾉ',
    'ハ': 'ﾊ', 'ヒ': 'ﾋ', 'フ': 'ﾌ', 'ヘ': 'ﾍ', 'ホ': 'ﾎ',
    'マ': 'ﾏ', 'ミ': 'ﾐ', 'ム': 'ﾑ', 'メ': 'ﾒ', 'モ': 'ﾓ',
    'ヤ': 'ﾔ', 'ユ': 'ﾕ', 'ヨ': 'ﾖ',
    'ラ': 'ﾗ', 'リ': 'ﾘ', 'ル': 'ﾙ', 'レ': 'ﾚ', 'ロ': 'ﾛ',
    'ワ': 'ﾜ', 'ン': 'ﾝ',
    'ガ': 'ｶﾞ', 'ギ': 'ｷﾞ', 'グ': 'ｸﾞ', 'ゲ': 'ｹﾞ', 'ゴ': 'ｺﾞ',
    'ザ': 'ｻﾞ', 'ジ': 'ｼﾞ', 'ズ': 'ｽﾞ', 'ゼ': 'ｾﾞ', 'ゾ': 'ｿﾞ',
    'ダ': 'ﾀﾞ', 'ヂ': 'ﾁﾞ', 'ヅ': 'ﾂﾞ', 'デ': 'ﾃﾞ', 'ド': 'ﾄﾞ',
    'バ': 'ﾊﾞ', 'ビ': 'ﾋﾞ', 'ブ': 'ﾌﾞ', 'ベ': 'ﾍﾞ', 'ボ': 'ﾎﾞ',
    'パ': 'ﾊﾟ', 'ピ': 'ﾋﾟ', 'プ': 'ﾌﾟ', 'ペ': 'ﾍﾟ', 'ポ': 'ﾎﾟ',
    'ヴ': 'ｳﾞ',
};

export function registerStringFunctions(ctx: StdlibCtx): void {
    // All string search/transform functions reject a Null comparison mode
    // before applying their operation. Keep this boundary shared so optional
    // Compare arguments do not diverge between implementations.
    const rejectNullCompare = (compare: any): void => {
        if (compare === vbaNull) ctx.throwError(VbaErrorCode.INVALID_USE_OF_NULL, 'Invalid use of Null');
    };
    const normalizeCompare = (compare: any): any => {
        rejectNullCompare(compare);
        if (compare === undefined || compare === vbaMissing) return compare;
        const value = compare;
        if (value !== 0 && value !== 1) {
            ctx.throwError(VbaErrorCode.INVALID_PROCEDURE_CALL, 'Invalid procedure call or argument');
        }
        return value;
    };
    const ascFunc = (s: any) => {
        if (s === vbaNull) return vbaNull;
        const str = vbaToString(s ?? '');
        if (str.length === 0) ctx.throwError(VbaErrorCode.INVALID_PROCEDURE_CALL, "Invalid procedure call or argument");
        return str.charCodeAt(0);
    };
    ctx.reg('asc', ascFunc, [{ name: 'String' }]);
    ctx.reg('ascw', ascFunc, [{ name: 'String' }]);
    const chrFunc = (n: any) => {
        if (n === vbaNull) return vbaNull;
        const code = n;
        if (code < 0 || code > 255) ctx.throwError(VbaErrorCode.INVALID_PROCEDURE_CALL, "Invalid procedure call or argument");
        return String.fromCharCode(code);
    };
    ctx.reg('chr', chrFunc, [{ name: 'CharCode', coerce: 'long' }], ['$']);
    const chrwFunc = (n: any) => {
        if (n === vbaNull) return vbaNull;
        const code = n;
        if (code < -32768 || code > 65535) ctx.throwError(VbaErrorCode.INVALID_PROCEDURE_CALL, "Invalid procedure call or argument");
        return String.fromCharCode(code < 0 ? code + 65536 : code);
    };
    ctx.reg('chrw', chrwFunc, [{ name: 'CharCode', coerce: 'long' }], ['$']);
    // Byte-oriented variants (UTF-16LE model: 1 char = 2 bytes, same as MidB)
    ctx.reg('lenb', (s: any) => s === vbaNull ? vbaNull : vbaToString(s ?? '').length * 2, [{ name: 'String' }]);
    ctx.reg('ascb', (s: any) => {
        if (s === vbaNull) return vbaNull;
        const str = vbaToString(s ?? '');
        if (str.length === 0) ctx.throwError(VbaErrorCode.INVALID_PROCEDURE_CALL, "Invalid procedure call or argument");
        return str.charCodeAt(0) & 0xFF;
    }, [{ name: 'String' }], ['$']);
    ctx.reg('chrb', (n: any) => {
        if (n === vbaNull) return vbaNull;
        const code = n;
        if (code < 0 || code > 255) ctx.throwError(VbaErrorCode.INVALID_PROCEDURE_CALL, "Invalid procedure call or argument");
        return String.fromCharCode(code);
    }, [{ name: 'CharCode', coerce: 'long' }], ['$']);
    // InStr: Start は先頭にある Optional 引数のため、引数の個数で意味が変わる
    const instrFunc = (...args: any[]) => {
        let start: any = 1, s1: any, s2: any, comp: any;
        if (args.length >= 4) [start, s1, s2, comp] = args;
        else if (args.length === 3) [start, s1, s2] = args;  // arg count determines form, not type
        else [s1, s2] = args;
        if (start === undefined) start = 1;
        if (start === vbaNull) ctx.throwError(VbaErrorCode.INVALID_USE_OF_NULL, 'Invalid use of Null');
        comp = normalizeCompare(comp);
        const startNum = start;
        if (startNum < 1) ctx.throwError(VbaErrorCode.INVALID_PROCEDURE_CALL, "Invalid procedure call or argument");
        if (s1 === vbaNull || s2 === vbaNull) return vbaNull;
        const str1 = vbaToString(s1 ?? ''), str2 = vbaToString(s2 ?? '');
        const isText = (comp === 1) || (comp === undefined && ctx.compMode === 'Text');
        const idx = isText ? str1.toLowerCase().indexOf(str2.toLowerCase(), startNum - 1) : str1.indexOf(str2, startNum - 1);
        return idx === -1 ? 0 : idx + 1;
    };
    ctx.regOvl('instr', instrFunc, [
        { params: [{ name: 'String1' }, { name: 'String2' }] },
        { params: [{ name: 'Start', optional: true, coerce: 'long' }, { name: 'String1' }, { name: 'String2' }] },
        { params: [{ name: 'Start', optional: true, coerce: 'long' }, { name: 'String1' }, { name: 'String2' }, { name: 'Compare', coerce: 'long' }] },
    ]);
    // InStrB: バイト単位での検索（VBA では 1 文字 = 2 バイト）
    const instrbFunc = (...args: any[]) => {
        let startByte = 1, s1, s2, comp;
        if (args.length >= 4) [startByte, s1, s2, comp] = args;
        else if (args.length === 3) [startByte, s1, s2] = args;
        else [s1, s2] = args;
        if (startByte === undefined) startByte = 1;
        comp = normalizeCompare(comp);
        if (s1 === vbaNull || s2 === vbaNull) return vbaNull;
        const str1 = vbaToString(s1 ?? ''), str2 = vbaToString(s2 ?? '');
        const startByteNum = startByte;
        if (startByteNum < 1) ctx.throwError(VbaErrorCode.INVALID_PROCEDURE_CALL, "Invalid procedure call or argument");
        const startChar = Math.floor((startByteNum - 1) / 2) + 1;
        const isText = (comp === 1) || (comp === undefined && ctx.compMode === 'Text');
        const idx = isText ? str1.toLowerCase().indexOf(str2.toLowerCase(), startChar - 1) : str1.indexOf(str2, startChar - 1);
        return idx === -1 ? 0 : idx * 2 + 1;
    };
    ctx.regOvl('instrb', instrbFunc, [
        { params: [{ name: 'String1' }, { name: 'String2' }] },
        { params: [{ name: 'Start', optional: true, coerce: 'long' }, { name: 'String1' }, { name: 'String2' }] },
        { params: [{ name: 'Start', optional: true, coerce: 'long' }, { name: 'String1' }, { name: 'String2' }, { name: 'Compare', coerce: 'long' }] },
    ]);
    ctx.reg('instrrev', (s1: any, s2: any, start: any = -1, comp: any = undefined) => {
        if (start === vbaNull) ctx.throwError(VbaErrorCode.INVALID_USE_OF_NULL, 'Invalid use of Null');
        comp = normalizeCompare(comp);
        const startNum = start;
        if (startNum !== -1 && startNum < 1) ctx.throwError(VbaErrorCode.INVALID_PROCEDURE_CALL, "Invalid procedure call or argument");
        if (s1 === vbaNull || s2 === vbaNull) return vbaNull;
        const str = vbaToString(s1 ?? ''), find = vbaToString(s2 ?? '');
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
        { name: 'Start', optional: true, coerce: 'long' },
        { name: 'Compare', optional: true, coerce: 'long' },
    ]);
    const lcaseFunc = (val: any) => val === vbaNull ? vbaNull : vbaToString(val ?? '').toLowerCase();
    ctx.reg('lcase', lcaseFunc, [{ name: 'String' }], ['$']);
    const strFunc = (val: any) => {
        if (val === vbaNull) ctx.throwError(VbaErrorCode.INVALID_USE_OF_NULL, 'Invalid use of Null');
        const n = ctx.toVbaNumber(val);
        // VBA の Str: 15 有効桁・|n|<1 は先頭の 0 を省略・指数は E+NN 形式（実 VBA 差分で裁定）
        let body: string;
        if (Number.isInteger(n) && Math.abs(n) < 1e15) {
            body = Math.abs(n).toString();
        } else if (n !== 0 && Math.abs(n) < 1e-4) {
            // |n| < 1e-4 は指数表記 E-NN（実 VBA 差分で裁定）
            body = Math.abs(n).toExponential(14)
                .replace(/\.?0+e/, 'e')
                .replace(/e([+-])(\d+)$/, (_, sg, d) => 'E' + sg + d.padStart(2, '0'));
        } else {
            body = Math.abs(n).toPrecision(15);
            if (body.includes('e')) {
                body = body.replace(/e([+-])(\d+)$/, (_, sg, d) => 'E' + sg + d.padStart(2, '0'));
                body = body.replace(/\.?0+E/, 'E');
            } else if (body.includes('.')) {
                body = body.replace(/\.?0+$/, '');
                if (body.startsWith('0.')) body = body.slice(1);
            }
        }
        return (n < 0 ? '-' : ' ') + body;
    };
    ctx.reg('str', strFunc, [{ name: 'Number' }], ['$']);
    const ucaseFunc = (val: any) => val === vbaNull ? vbaNull : vbaToString(val ?? '').toUpperCase();
    ctx.reg('ucase', ucaseFunc, [{ name: 'String' }], ['$']);
    const leftFunc = (val: any, len: any) => {
        if (val === vbaNull || len === vbaNull) return vbaNull;
        const l = len;
        if (l < 0) ctx.throwError(VbaErrorCode.INVALID_PROCEDURE_CALL, "Invalid procedure call or argument");
        return vbaToString(val ?? '').substring(0, l);
    };
    ctx.reg('left', leftFunc, [{ name: 'String' }, { name: 'Length', coerce: 'long' }], ['$']);
    const rightFunc = (val: any, len: any) => {
        if (val === vbaNull || len === vbaNull) return vbaNull;
        const s = vbaToString(val ?? ''), l = len;
        if (l < 0) ctx.throwError(VbaErrorCode.INVALID_PROCEDURE_CALL, "Invalid procedure call or argument");
        return s.substring(s.length - l);
    };
    ctx.reg('right', rightFunc, [{ name: 'String' }, { name: 'Length', coerce: 'long' }], ['$']);
    const midFunc = (val: any, start: any, len?: any) => {
        if (val === vbaNull || start === vbaNull || len === vbaNull) return vbaNull;
        const s = vbaToString(val ?? ''), st = start;
        if (st < 1) ctx.throwError(VbaErrorCode.INVALID_PROCEDURE_CALL, "Invalid procedure call or argument");
        const length = len === undefined ? undefined : len;
        if (length !== undefined && length < 0) ctx.throwError(VbaErrorCode.INVALID_PROCEDURE_CALL, "Invalid procedure call or argument");
        return length !== undefined ? s.substring(st - 1, st - 1 + length) : s.substring(st - 1);
    };
    ctx.reg('mid', midFunc, [
        { name: 'String' },
        { name: 'Start', coerce: 'long' },
        { name: 'Length', optional: true, coerce: 'long' },
    ], ['$']);
    ctx.reg('len', (val: any) => val === vbaNull ? vbaNull : vbaToString(val ?? '').length, [{ name: 'String' }]);
    // VBA の Trim 系はスペース（Chr 32）のみを除去する。タブ等は残す（実 VBA 差分で裁定）
    const ltrimFunc = (val: any) => val === vbaNull ? vbaNull : vbaToString(val ?? '').replace(/^ +/, '');
    ctx.reg('ltrim', ltrimFunc, [{ name: 'String' }], ['$']);
    const rtrimFunc = (val: any) => val === vbaNull ? vbaNull : vbaToString(val ?? '').replace(/ +$/, '');
    ctx.reg('rtrim', rtrimFunc, [{ name: 'String' }], ['$']);
    const trimFunc = (val: any) => val === vbaNull ? vbaNull : vbaToString(val ?? '').replace(/^ +| +$/g, '');
    ctx.reg('trim', trimFunc, [{ name: 'String' }], ['$']);
    // Number 引数を Long として検証し、文字列確保失敗（JS RangeError）を Error 14 に写像する
    const repeatChecked = (c: string, n: any): string => {
        const count = ctx.round(ctx.toVbaNumber(n));
        if (count < 0) ctx.throwError(VbaErrorCode.INVALID_PROCEDURE_CALL, "Invalid procedure call or argument");
        if (count > 2147483647) ctx.throwError(VbaErrorCode.OVERFLOW, "Overflow");
        try {
            return c.repeat(count);
        } catch (e) {
            if (e instanceof RangeError) ctx.throwError(VbaErrorCode.OUT_OF_STRING_SPACE, "Out of string space");
            throw e;
        }
    };
    const spaceFunc = (n: any) => {
        if (n === vbaNull) ctx.throwError(VbaErrorCode.INVALID_USE_OF_NULL, "Invalid use of Null");
        return repeatChecked(' ', n);
    };
    ctx.reg('space', spaceFunc, [{ name: 'Number' }], ['$']);
    const stringFunc = (n: any, char: any) => {
        if (n === vbaNull || char === vbaNull) return vbaNull;
        let c: string;
        if (typeof char === 'number' || char instanceof VbaCurrency || char instanceof VbaDecimal) {
            // §6.1.2.11.1.38: numbers > 255 use character Mod 256
            c = String.fromCharCode(ctx.round(ctx.toVbaNumber(char)) % 256);
        } else {
            const s = vbaToString(char ?? '');
            // Empty string character is invalid per spec
            if (s.length === 0) ctx.throwError(VbaErrorCode.INVALID_PROCEDURE_CALL, "Invalid procedure call or argument");
            c = s[0];
        }
        return repeatChecked(c, n);
    };
    ctx.reg('string', stringFunc, [{ name: 'Number' }, { name: 'Character' }], ['$']);
    const coerceStringDelimiter = (del: any): string => del === vbaEmpty ? '' : String(del);
    ctx.reg('split', (s: any, del: any = ' ', limit: any = -1, compare: any = undefined) => {
        if (s === vbaNull) return vbaNull;
        if (del === vbaNull || limit === vbaNull) ctx.throwError(VbaErrorCode.INVALID_USE_OF_NULL, 'Invalid use of Null');
        compare = normalizeCompare(compare);
        const str = vbaToString(s ?? '');
        const delimiter = del === null || del === undefined ? ' ' : coerceStringDelimiter(del);
        const n = limit === undefined ? -1 : limit;
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
        (result as any).__vbaArrayReturn__ = true;
        (result as any).__vbaArrayReturnType__ = 'String';
        return result;
    }, [
        { name: 'Expression' },
        { name: 'Delimiter', optional: true },
        { name: 'Limit', optional: true, coerce: 'long' },
        { name: 'Compare', optional: true, coerce: 'long' },
    ]);
    ctx.reg('join', (arr: any, del: any = ' ') => {
        if (del === vbaNull) return vbaNull;
        if (!Array.isArray(arr)) ctx.throwError(VbaErrorCode.TYPE_MISMATCH, 'Type mismatch');
        const base = (arr as any).vbaBase || 0;
        const elems = arr.slice(base).map((el: any) => {
            if (el === vbaNull) ctx.throwError(VbaErrorCode.TYPE_MISMATCH, 'Type mismatch');
            return el === undefined || el === vbaEmpty ? '' : String(el);
        });
        return elems.join(coerceStringDelimiter(del));
    }, [
        { name: 'SourceArray' },
        { name: 'Delimiter', optional: true },
    ]);
    ctx.reg('replace', (s: any, f: any, r: any, start: any = 1, count: any = -1, compare: any = undefined) => {
        if (s === vbaNull) ctx.throwError(VbaErrorCode.INVALID_USE_OF_NULL, 'Invalid use of Null');
        if (f === vbaNull || r === vbaNull) return vbaNull;
        if (start === vbaNull || count === vbaNull) {
            ctx.throwError(VbaErrorCode.INVALID_USE_OF_NULL, 'Invalid use of Null');
        }
        compare = normalizeCompare(compare);
        const str = vbaToString(s ?? '');
        const find = vbaToString(f ?? '');
        const repl = vbaToString(r ?? '');
        const startNum = start ?? 1;
        if (startNum < 1) ctx.throwError(VbaErrorCode.INVALID_PROCEDURE_CALL, 'Invalid procedure call or argument');
        const countNum = count ?? -1;
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
        { name: 'Start', optional: true, coerce: 'long' },
        { name: 'Count', optional: true, coerce: 'long' },
        { name: 'Compare', optional: true, coerce: 'long' },
    ]);
    ctx.reg('strcomp', (s1: any, s2: any, comp?: number) => {
        if (s1 === vbaNull || s2 === vbaNull) return vbaNull;
        comp = normalizeCompare(comp);
        let str1 = vbaToString(s1 ?? ''), str2 = vbaToString(s2 ?? '');
        const isText = (comp === 1) || (comp === undefined && ctx.compMode === 'Text');
        if (isText) { str1 = str1.toLowerCase(); str2 = str2.toLowerCase(); }
        return str1 < str2 ? -1 : (str1 > str2 ? 1 : 0);
    }, [
        { name: 'String1' },
        { name: 'String2' },
        { name: 'Compare', optional: true, coerce: 'long' },
    ]);
    ctx.reg('strconv', (s: any, conv: any, lcid: any = vbaMissing) => {
        if (s === vbaNull) return vbaNull;
        if (conv === vbaNull) ctx.throwError(VbaErrorCode.INVALID_USE_OF_NULL, 'Invalid use of Null');
        if (lcid !== vbaMissing && lcid !== undefined) {
            const localeId = ctx.toVbaNumber(lcid);
            if (!Number.isInteger(localeId) || localeId <= 0 || localeId > 0xFFFF) {
                ctx.throwError(VbaErrorCode.INVALID_PROCEDURE_CALL, 'Invalid procedure call or argument');
            }
        }
        let str = vbaToString(s ?? '');
        const c = ctx.toVbaNumber(conv);
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
            // 半角カナ → 全角カタカナ。濁点/半濁点付き（ｶﾞ→ガ 等）を含め NFKC 正規化で正確に変換できる
            str = str.replace(/[｡-ﾟ]+/g, m => m.normalize('NFKC'));
        } else if (c & 8) {
            str = str.replace(/[！-～]/g, m => String.fromCharCode(m.charCodeAt(0) - 0xFEE0)).replace(/　/g, ' ');
            // 全角カタカナ → 半角カナ。濁点/半濁点付き文字は基底文字 + 結合文字の2文字表現になる
            str = str.replace(/[ァ-ヴー。「」、・]/g, m => FULL_TO_HALF_KANA[m] ?? m);
        }
        return str;
    }, [{ name: 'String' }, { name: 'Conversion' }, { name: 'LCID', optional: true }]);
    ctx.reg('strreverse', (s: any) => {
        if (s === vbaNull) ctx.throwError(VbaErrorCode.INVALID_USE_OF_NULL, 'Invalid use of Null');
        return vbaToString(s ?? '').split('').reverse().join('');
    }, [{ name: 'Expression' }]);
    ctx.reg('filter', (source: any, match: any, include: any = vbaTrue, compare: any = undefined) => {
        if (!Array.isArray(source)) ctx.throwError(VbaErrorCode.TYPE_MISMATCH, "Type mismatch");
        const srcDims = (source as any).__vbaDimensions__;
        if (srcDims && srcDims.length > 1) ctx.throwError(VbaErrorCode.TYPE_MISMATCH, "Type mismatch");
        if (match === vbaNull) ctx.throwError(VbaErrorCode.INVALID_USE_OF_NULL, "Invalid use of Null");
        compare = normalizeCompare(compare);
        const find = vbaToString(match ?? '');
        const isInclude = ctx.isTrue(include);
        const isText = (compare === 1) || (compare === undefined && ctx.compMode === 'Text');
        const result = source.filter((s: any) => {
            const str = vbaToString(s ?? '');
            const found = isText ? str.toLowerCase().includes(find.toLowerCase()) : str.includes(find);
            return isInclude ? found : !found;
        });
        (result as any).vbaBase = 0;
        (result as any).__vbaArrayReturn__ = true;
        (result as any).__vbaArrayReturnType__ = 'Variant';
        return result;
    }, [
        { name: 'SourceArray' },
        { name: 'Match' },
        { name: 'Include', optional: true },
        { name: 'Compare', optional: true, coerce: 'long' },
    ]);
    ctx.reg('leftb', (val: any, len: any) => {
        if (val === vbaNull || len === vbaNull) return vbaNull;
        const byteLen = len;
        if (byteLen < 0) ctx.throwError(VbaErrorCode.INVALID_PROCEDURE_CALL, "Invalid procedure call or argument");
        const s = vbaToString(val ?? '');
        return s.substring(0, Math.floor(byteLen / 2));
    }, [{ name: 'String' }, { name: 'Length', coerce: 'long' }]);
    ctx.reg('rightb', (val: any, len: any) => {
        if (val === vbaNull || len === vbaNull) return vbaNull;
        const byteLen = len;
        if (byteLen < 0) ctx.throwError(VbaErrorCode.INVALID_PROCEDURE_CALL, "Invalid procedure call or argument");
        const s = vbaToString(val ?? '');
        const charLen = Math.floor(byteLen / 2);
        return s.substring(s.length - charLen);
    }, [{ name: 'String' }, { name: 'Length', coerce: 'long' }]);
    const midbFunc = (val: any, start: any, len?: any) => {
        if (val === vbaNull || start === vbaNull || len === vbaNull) return vbaNull;
        const s = vbaToString(val ?? '');
        const byteStart = start;
        if (byteStart < 1) ctx.throwError(VbaErrorCode.INVALID_PROCEDURE_CALL, "Invalid procedure call or argument");
        const byteLen = len === undefined ? undefined : len;
        if (byteLen !== undefined && byteLen < 0) ctx.throwError(VbaErrorCode.INVALID_PROCEDURE_CALL, "Invalid procedure call or argument");
        const charStart = Math.floor((byteStart + 1) / 2);
        if (byteLen === undefined) return s.substring(charStart - 1);
        const charLen = Math.floor(byteLen / 2);
        return s.substring(charStart - 1, charStart - 1 + charLen);
    };
    ctx.reg('midb', midbFunc, [
        { name: 'String' },
        { name: 'Start', coerce: 'long' },
        { name: 'Length', optional: true, coerce: 'long' },
    ], ['$']);

    /** Format a Decimal/Currency string without passing through JS Number. */
    const exactFixedFormat = (raw: string, decimals: number, group: boolean,
                              includeLeadingDigit = true, prefix = '', suffix = '',
                              useParens = false, scale = 1): string => {
        let text = raw.trim();
        let negative = text.startsWith('-');
        if (negative || text.startsWith('+')) text = text.slice(1);
        let [intPart, fracPart = ''] = text.split('.');
        intPart = intPart || '0';
        if (scale === 100) {
            const all = intPart + fracPart;
            const point = intPart.length + 2;
            if (point >= all.length) {
                intPart = all + '0'.repeat(point - all.length);
                fracPart = '';
            } else {
                intPart = all.slice(0, point);
                fracPart = all.slice(point);
            }
        }
        const kept = fracPart.slice(0, decimals).padEnd(decimals, '0');
        const next = fracPart[decimals];
        let scaledDigits = (intPart + kept).replace(/^0+(?=\d)/, '') || '0';
        if (next !== undefined && next >= '5') {
            scaledDigits = (BigInt(scaledDigits || '0') + 1n).toString();
        }
        if (decimals > 0) {
            scaledDigits = scaledDigits.padStart(decimals + 1, '0');
            intPart = scaledDigits.slice(0, -decimals) || '0';
            fracPart = scaledDigits.slice(-decimals);
        } else {
            intPart = scaledDigits;
            fracPart = '';
        }
        if (!includeLeadingDigit && intPart === '0') intPart = '';
        if (group && intPart) intPart = intPart.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
        const number = intPart + (fracPart ? `.${fracPart}` : '');
        const body = `${prefix}${number}${suffix}`;
        if (!negative) return body;
        return useParens ? `(${body})` : `-${body}`;
    };

    const decodeFormatLiterals = (text: string): string => {
        let result = '';
        let quoted = false;
        for (let i = 0; i < text.length; i++) {
            const char = text[i];
            if (char === '"') {
                if (quoted && text[i + 1] === '"') {
                    result += '"';
                    i++;
                } else {
                    quoted = !quoted;
                }
            } else if (!quoted && char === '\\' && i + 1 < text.length) {
                result += text[++i];
            } else {
                result += char;
            }
        }
        return result;
    };

    const exactPatternFormat = (raw: string, pattern: string): string | undefined => {
        const clean = stripVbaBracketDirectives(stripFormatColorDirectives(pattern));
        const sections = clean.split(';');
        const trimmed = raw.trim();
        const negative = trimmed.startsWith('-');
        const zero = /^[+-]?0(?:\.0*)?$/.test(trimmed);
        const section = negative && sections.length > 1
            ? sections[1]
            : zero && sections.length > 2
                ? sections[2]
                : sections[0];
        const first = section.search(/[0#]/);
        const last = Math.max(section.lastIndexOf('0'), section.lastIndexOf('#'));
        if (first < 0 || last < first || /[Ee][+-]/.test(section)) return undefined;
        if (!hasNumericFormatTokens(section)) return undefined;
        const prefix = decodeFormatLiterals(section.slice(0, first));
        const suffix = decodeFormatLiterals(section.slice(last + 1));
        const core = section.slice(first, last + 1);
        const percent = core.includes('%');
        const numericCore = core.replace(/%/g, '').replace(/,/g, '');
        const parts = numericCore.split('.');
        const intFmt = parts[0] || '';
        const decFmt = parts[1] || '';
        const decimals = decFmt.length;
        const group = core.includes(',');
        const includeLeading = (intFmt.match(/0/g) || []).length > 0;
        const formatted = exactFixedFormat(raw, decimals, group, includeLeading, prefix,
            `${suffix}${percent ? '%' : ''}`, false, percent ? 100 : 1);
        if (negative && sections.length > 1 && section.includes('(')) {
            return formatted.replace(/^-/, '');
        }
        return formatted;
    };

    const exactScientificFormat = (raw: string, pattern: string): string | undefined => {
        const clean = stripVbaBracketDirectives(stripFormatColorDirectives(pattern));
        const match = clean.match(/^(.*?)([Ee][+-])(0+|#+)(.*?)$/);
        if (!match || !/[0#]/.test(match[1])) return undefined;
        const mantissaFmt = match[1].replace(/%/g, '').replace(/,/g, '');
        const parts = mantissaFmt.split('.');
        const intFmt = parts[0] || '0';
        const decFmt = parts[1] || '';
        if (!/^0[#]*$/.test(intFmt) || /[^0#]/.test(decFmt)) return undefined;
        let text = raw.trim();
        const negative = text.startsWith('-');
        if (negative || text.startsWith('+')) text = text.slice(1);
        let [intPart, fracPart = ''] = text.split('.');
        intPart = intPart || '0';
        const digits = (intPart + fracPart).replace(/^0+/, '');
        if (!digits) {
            const zeroMantissa = intFmt.includes('0')
                ? `0${decFmt ? `.${'0'.repeat((decFmt.match(/0/g) || []).length)}` : ''}`
                : '';
            const zeroSign = match[2][1] === '+' ? '+' : '';
            return `${negative ? '-' : ''}${zeroMantissa}${match[2][0]}${zeroSign}${'0'.repeat(match[3].length)}${decodeFormatLiterals(match[4])}`;
        }
        const firstNonzero = (intPart + fracPart).search(/[1-9]/);
        let exponent = intPart.length - firstNonzero - 1;
        const maxDecimals = decFmt.length;
        let significant = digits.slice(0, maxDecimals + 1).padEnd(maxDecimals + 1, '0');
        const next = digits[maxDecimals + 1];
        if (next !== undefined && next >= '5') {
            significant = (BigInt(significant) + 1n).toString().padStart(maxDecimals + 1, '0');
        }
        if (significant.length > maxDecimals + 1) {
            significant = significant.slice(0, maxDecimals + 1);
            exponent++;
        }
        const mantissa = significant[0] + (maxDecimals ? `.${significant.slice(1)}` : '');
        const eLetter = match[2][0];
        const expSign = exponent < 0 ? '-' : (match[2][1] === '+' ? '+' : '');
        const expDigits = String(Math.abs(exponent)).padStart(match[3].length, '0');
        // A comma immediately after the exponent is a scale marker, just as
        // it is after an ordinary numeric section.  It is not a literal
        // suffix.  Preserve quoted/escaped commas while removing only the
        // unescaped trailing scale markers.
        let suffix = match[4];
        while (suffix.endsWith(',') && !suffix.endsWith('\\,')) {
            let quoted = false;
            for (let i = 0; i < suffix.length - 1; i++) {
                if (suffix[i] === '"' && suffix[i - 1] !== '\\') quoted = !quoted;
            }
            if (quoted) break;
            suffix = suffix.slice(0, -1);
        }
        return `${negative ? '-' : ''}${match[1].replace(/[0#.,]+/, mantissa)}${eLetter}${expSign}${expDigits}${decodeFormatLiterals(suffix)}`;
    };

    const normalizeFormatWeekArg = (value: any, max: number): number => {
        if (value === undefined || value === vbaMissing) return 1;
        if (value === vbaNull) ctx.throwError(VbaErrorCode.INVALID_USE_OF_NULL, 'Invalid use of Null');
        const n = vbaRound(ctx.toVbaNumber(value));
        if (n < 0 || n > max) ctx.throwError(VbaErrorCode.INVALID_PROCEDURE_CALL, 'Invalid procedure call or argument');
        return n;
    };
    const formatFunc = (val: any, pattern?: any, firstDayOfWeek?: any, firstWeekOfYear?: any) => {
        val = unwrapVbaDefaultValue(val);
        if (pattern === vbaNull) ctx.throwError(VbaErrorCode.INVALID_USE_OF_NULL, 'Invalid use of Null');
        const fmt = pattern && pattern !== vbaMissing ? vbaToString(pattern) : "";
        if (val === vbaNull) {
            const nullSection = formatNullSection(fmt);
            return nullSection === undefined ? "" : nullSection;
        }
        if (val === vbaEmpty) {
            // Empty selects the fourth user-defined section just like the
            // Excel Format contract observed for XL-180-EMPTY.  Keep the
            // historical empty result when no fourth section exists.
            const emptySection = formatNullSection(fmt);
            return emptySection === undefined ? "" : emptySection;
        }
        if (val === null) return "";
        const fdow = normalizeFormatWeekArg(firstDayOfWeek, 7);
        const fwoy = normalizeFormatWeekArg(firstWeekOfYear, 3);
        const effectiveFwoy = fwoy === 0 ? 1 : fwoy;
        if (fmt === "") return vbaToString(val);
        const escapedBracketMatch = fmt.match(/^\\\[(<=|>=|<|>)(\d+)\\\]$/);
        if (escapedBracketMatch) {
            const prefix = `[=${escapedBracketMatch[2]}]`;
            if (val instanceof VbaDate) {
                return prefix + formatDate(fromVbaDate(val.value), 'General Date', fdow, effectiveFwoy);
            }
            if (val instanceof VbaCurrency || val instanceof VbaDecimal) {
                return prefix + val.toString();
            }
            // Excel lowercases a string value when this escaped comparison
            // directive is formatted as a literal prefix.  Keep this
            // matrix-observed behavior separate from ordinary string
            // placeholder formatting, which preserves case.
            if (typeof val === 'string') {
                return prefix + val.toLowerCase();
            }
            return prefix + vbaToString(val);
        }
        if (typeof val === 'bigint') {
            const scientificPattern = fmt.toLowerCase() === 'scientific' ? '0.00E+00' : fmt;
            const exactScientific = exactScientificFormat(val.toString(), scientificPattern);
            if (exactScientific !== undefined) return exactScientific;
            const namedExact: Record<string, [number, boolean, number, string, string]> = {
                'general number': [0, false, 1, '', ''],
                'currency': [2, true, 1, '$', ''],
                'fixed': [2, false, 1, '', ''],
                'standard': [2, true, 1, '', ''],
                'percent': [2, false, 100, '', '%'],
            };
            const spec = namedExact[fmt.toLowerCase()];
            if (spec) return exactFixedFormat(val.toString(), spec[0], spec[1], true, spec[3], spec[4], false, spec[2]);
            const exact = exactPatternFormat(val.toString(), fmt);
            if (exact !== undefined) return exact;
        }
        if (val instanceof VbaCurrency || val instanceof VbaDecimal) {
            const exactScientific = exactScientificFormat(val.toString(), fmt.toLowerCase() === 'scientific' ? '0.00E+00' : fmt);
            if (exactScientific !== undefined) return exactScientific;
            if (fmt.toLowerCase() === 'general number') return val.toString();
            const namedExact: Record<string, [number, boolean, number, string, string]> = {
                'general number': [0, false, 1, '', ''],
                'currency': [2, true, 1, '$', ''],
                'fixed': [2, false, 1, '', ''],
                'standard': [2, true, 1, '', ''],
                'percent': [2, false, 100, '', '%'],
            };
            const spec = namedExact[fmt.toLowerCase()];
            if (spec) return exactFixedFormat(val.toString(), spec[0], spec[1], true, spec[3], spec[4], false, spec[2]);
        }
        if (val instanceof VbaCurrency || val instanceof VbaDecimal) {
            const exact = exactPatternFormat(val.toString(), fmt);
            if (exact !== undefined) return exact;
        }
        const fmtLower = fmt.toLowerCase();
        const namedFormats = ['general number', 'currency', 'fixed', 'standard', 'percent', 'scientific', 'true/false', 'yes/no', 'on/off'];
        const dateNamedFormats = ['general date', 'long date', 'medium date', 'short date', 'long time', 'medium time', 'short time'];
        if (namedFormats.includes(fmtLower)) {
            if (fmtLower === 'yes/no' || fmtLower === 'on/off' || fmtLower === 'true/false') {
                // Excel preserves a non-numeric String for named Boolean
                // formats instead of coercing it to Null/False.
                if (typeof val === 'string' && !/^\s*[+-]?(?:\d+(?:\.\d*)?|\.\d+)\s*$/.test(val)) {
                    return val;
                }
                const isTrue = val instanceof VbaBoolean ? val.value !== 0 : (ctx.toVbaNumber(val) !== 0);
                if (fmtLower === 'yes/no') return isTrue ? 'Yes' : 'No';
                if (fmtLower === 'on/off') return isTrue ? 'On' : 'Off';
                return isTrue ? 'True' : 'False';
            }
            const numVal = (val instanceof VbaBoolean) ? val.value
                : (val instanceof VbaDate) ? val.value
                : (val instanceof VbaCurrency || val instanceof VbaDecimal) ? ctx.toVbaNumber(val.toString())
                : val;
            if (typeof numVal === 'number') return formatNumber(numVal, fmt);
            return vbaToString(val);
        }
        if (dateNamedFormats.includes(fmtLower)) {
            if (typeof val === 'string') {
                try { return formatDate(parseVbaDate(val), fmt, fdow, effectiveFwoy); }
                catch { return val; }
            }
            const dateSerial = val instanceof VbaDate ? val.value
                : (val instanceof VbaCurrency || val instanceof VbaDecimal) ? ctx.toVbaNumber(val.toString())
                : val;
            const dateVal = fromVbaDate(typeof dateSerial === 'number' ? dateSerial : ctx.toVbaNumber(dateSerial));
            return formatDate(dateVal, fmt, fdow, effectiveFwoy);
        }
        const effectiveVal = (val instanceof VbaBoolean) ? val.value
            : (val instanceof VbaCurrency || val instanceof VbaDecimal) ? ctx.toVbaNumber(val.toString())
            : val;
        // For numeric values, only the selected positive/negative/zero
        // section can select the date formatter.  A quoted literal in the
        // Null section (for example "NULL") must not affect that dispatch.
        const datePatternSource = (() => {
            const selectionNumber = effectiveVal instanceof VbaDate ? effectiveVal.value : effectiveVal;
            if (typeof effectiveVal === 'string') {
                // A non-empty String selects the first string section before
                // the format domain is classified.  Date tokens in the
                // inactive Null/empty-string section must not redirect the
                // String value into the date renderer.
                if (effectiveVal.length > 0) {
                    return splitFormatSections(fmt)[0] ?? fmt;
                }
                return fmt;
            }
            if (typeof selectionNumber !== 'number') return fmt;
            const sections = splitFormatSections(fmt);
            if (sections.length < 2) return sections[0];
            if (!hasNumericFormatTokens(sections[0])) return sections[0];
            if (selectionNumber < 0) return sections[1] || sections[0];
            if (selectionNumber === 0 && sections.length >= 3) return sections[2] || sections[0];
            return sections[0];
        })();
        const sectionsHaveNumericSelector = (() => {
            const sections = splitFormatSections(fmt);
            return sections.length > 1 && hasNumericFormatTokens(sections[0]);
        })();
        const isDatePattern = containsDateFormatTokens(datePatternSource) &&
            !(sectionsHaveNumericSelector && !hasNumericFormatTokens(datePatternSource));
        if (typeof effectiveVal === 'string') {
            // A non-numeric String stays in the string domain when the user
            // format mixes numeric placeholders with @/&/!/<> controls.
            // Sending it to formatString makes 0/0.00 literal text, unlike
            // Excel's mixed-domain behavior observed by XL-180.
            if (hasMixedNumericStringTokens(fmt)) {
                try {
                    ctx.toVbaNumber(effectiveVal);
                } catch {
                    return effectiveVal;
                }
            }
            // Resolve numeric strings before scanning the complete pattern for
            // date tokens in unselected sections.
            if (isNumericOnlyFormat(fmt)) {
                try {
                    return formatNumber(ctx.toVbaNumber(effectiveVal), fmt);
                } catch {
                    // Excel treats a date-like String as a Date serial when a
                    // numeric user format is selected. Keep this conversion
                    // limited to numeric-only formats; mixed string formats
                    // remain in the String domain.
                    try {
                        const parsed = parseVbaDate(effectiveVal);
                        return formatNumber(toVbaDate(parsed), fmt);
                    } catch { /* retain the existing string-domain behavior */ }
                }
            }
            // If the format contains date/time symbols, try to parse the string as a date first
            if (isDatePattern && !/^[0#,.%]+$/.test(fmt)) {
                try {
                    const parsed = parseVbaDate(effectiveVal);
                    return formatDate(parsed, fmt, fdow, effectiveFwoy);
                } catch { /* fall through to string formatting */ }
            }
            // VBA applies user-defined string formats only when the pattern
            // contains a string-domain control (`@`, `&`, `!`, `<`, or `>`).
            // Numeric/date tokens and literal-only patterns do not convert a
            // String value into formatted output; Excel returns the input.
            if (!hasUnescapedFormatChars(fmt, '@&!<>')) return effectiveVal;
            return formatString(effectiveVal, fmt);
        }
        if (effectiveVal instanceof VbaDate) {
            // A Date expression can be formatted either as a date or as its
            // underlying serial number. Numeric-only formats use the serial
            // value in VBA instead of the date-token formatter.
            const normalizedDatePattern = stripVbaBracketDirectives(stripFormatColorDirectives(fmt));
            if (hasUnclosedVbaBracket(fmt)) return '';
            if (!normalizedDatePattern) {
                return formatDate(fromVbaDate(effectiveVal.value), 'General Date', fdow, effectiveFwoy);
            }
            if (normalizedDatePattern === '.') {
                return formatNumber(effectiveVal.value, normalizedDatePattern);
            }
            if (!isDatePattern && !hasNumericFormatTokens(normalizedDatePattern)) {
                if (hasUnescapedFormatChars(normalizedDatePattern, '!*')) return '';
                if (hasUnescapedFormatChars(normalizedDatePattern, '@&<>')) {
                    return formatDate(fromVbaDate(effectiveVal.value), 'General Date', fdow, effectiveFwoy);
                }
                if (!datePatternSource && normalizedDatePattern) return '';
                return formatDate(fromVbaDate(effectiveVal.value), datePatternSource, fdow, effectiveFwoy);
            }
            if (isDatePattern && !/^[0#,.%]+$/.test(fmt)) {
                return formatDate(fromVbaDate(effectiveVal.value), datePatternSource, fdow, effectiveFwoy);
            }
            return formatNumber(effectiveVal.value, fmt);
        }
        if (typeof effectiveVal === 'number') {
            if (isDatePattern && !/^[0#,.%]+$/.test(fmt)) return formatDate(fromVbaDate(effectiveVal), fmt, fdow, effectiveFwoy);
            return formatNumber(effectiveVal, fmt);
        }
        return vbaToString(effectiveVal);
    };
    ctx.reg('format', formatFunc, [
        { name: 'Expression' },
        { name: 'Format', optional: true },
        { name: 'FirstDayOfWeek', optional: true },
        { name: 'FirstWeekOfYear', optional: true },
    ], ['$']);

    // Helper shared by FormatCurrency / FormatNumber / FormatPercent
    const fmtNumeric = (val: any, digits: any, leadingDigit: any, parens: any, groupDigits: any,
                        prefix: string, suffix: string, scale: number): string => {
        val = unwrapVbaDefaultValue(val);
        if (val === vbaNull) return '';
        for (const option of [digits, leadingDigit, parens, groupDigits]) {
            if (option === vbaNull) ctx.throwError(VbaErrorCode.INVALID_USE_OF_NULL, 'Invalid use of Null');
        }
        let n: number;
        if (val instanceof VbaDate) n = val.value;
        else if (val instanceof VbaCurrency) n = ctx.toVbaNumber(val.internal) / 10000;
        else if (val instanceof VbaDecimal) n = ctx.toVbaNumber(val.mantissa) / Math.pow(10, val.scale);
        else n = ctx.toVbaNumber(val);
        n *= scale;
        const dec = digits === vbaMissing || ctx.toVbaNumber(digits) < 0 ? 2 : ctx.round(ctx.toVbaNumber(digits));
        // VBA の NumDigitsAfterDecimal 上限は 255（JS toFixed は 100 までなので超過分は 0 埋め）
        if (dec > 255) ctx.throwError(VbaErrorCode.INVALID_PROCEDURE_CALL, "Invalid procedure call or argument");
        const group = groupDigits === vbaMissing || ctx.toVbaNumber(groupDigits) < 0 ? true : ctx.toVbaNumber(groupDigits) !== 0;
        // vbUseDefault (-2) and vbTrue (-1) include the leading zero.  vbFalse
        // suppresses it for values whose formatted magnitude is below one.
        const includeLeadingDigit = leadingDigit === vbaMissing || ctx.toVbaNumber(leadingDigit) !== 0;
        const useParensForNegative = parens !== vbaMissing && ctx.toVbaNumber(parens) !== 0;
        if (val instanceof VbaCurrency || val instanceof VbaDecimal || typeof val === 'bigint') {
            return exactFixedFormat(val.toString(), dec, group, includeLeadingDigit,
                prefix, suffix, useParensForNegative, scale);
        }
        const neg = n < 0;
        const abs = Math.abs(n);
        const decFixed = Math.min(dec, 100);
        let formatted = vbaRound(abs, decFixed).toFixed(decFixed) + '0'.repeat(dec - decFixed);
        if (!includeLeadingDigit && abs < 1) formatted = formatted.substring(1);
        if (group) {
            const [intPart, fracPart] = formatted.split('.');
            const grouped = intPart.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
            formatted = fracPart !== undefined ? `${grouped}.${fracPart}` : grouped;
        }
        const result = `${prefix}${formatted}${suffix}`;
        if (!neg) return result;
        return useParensForNegative ? `(${result})` : `-${result}`;
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
        val = unwrapVbaDefaultValue(val);
        if (val === vbaNull) return '';
        if (namedFmt === vbaNull) ctx.throwError(VbaErrorCode.INVALID_USE_OF_NULL, 'Invalid use of Null');
        const d = (val instanceof VbaDate) ? fromVbaDate(val.value) : parseVbaDate(val);
        const fmt = vbaRound(ctx.toVbaNumber(namedFmt ?? 0));
        if (fmt < 0 || fmt > 4) ctx.throwError(VbaErrorCode.INVALID_PROCEDURE_CALL, 'Invalid procedure call or argument');
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
    const ensureVbaDate = (date: Date): Date => {
        if (!Number.isFinite(date.getTime()) || date.getFullYear() < 100 || date.getFullYear() > 9999) {
            ctx.throwError(VbaErrorCode.OVERFLOW, 'Overflow');
        }
        return date;
    };
    const coerceDateNumber = (value: any): number => ctx.round(ctx.toVbaNumber(value));
    const validateFirstDayOfWeek = (value: any): number => {
        const day = coerceDateNumber(value ?? 1);
        if (!Number.isFinite(day) || day < 0 || day > 7) {
            ctx.throwError(VbaErrorCode.INVALID_PROCEDURE_CALL, 'Invalid procedure call or argument');
        }
        return day;
    };
    const validateFirstWeekOfYear = (value: any): number => {
        rejectNullArgument(ctx, value);
        const week = coerceDateNumber(value ?? 1);
        if (!Number.isFinite(week) || week < 0 || week > 3) {
            ctx.throwError(VbaErrorCode.INVALID_PROCEDURE_CALL, 'Invalid procedure call or argument');
        }
        return week;
    };
    ctx.reg('year',   (d: any) => { d = unwrapVbaDefaultValue(d); return d === vbaNull ? vbaNull : parseVbaDate(d).getFullYear(); }, [{ name: 'Date' }]);
    ctx.reg('month',  (d: any) => { d = unwrapVbaDefaultValue(d); return d === vbaNull ? vbaNull : parseVbaDate(d).getMonth() + 1; }, [{ name: 'Date' }]);
    ctx.reg('day',    (d: any) => { d = unwrapVbaDefaultValue(d); return d === vbaNull ? vbaNull : parseVbaDate(d).getDate(); }, [{ name: 'Date' }]);
    ctx.reg('hour',   (d: any) => { d = unwrapVbaDefaultValue(d); return d === vbaNull ? vbaNull : parseVbaDate(d).getHours(); }, [{ name: 'Time' }]);
    ctx.reg('minute', (d: any) => { d = unwrapVbaDefaultValue(d); return d === vbaNull ? vbaNull : parseVbaDate(d).getMinutes(); }, [{ name: 'Time' }]);
    ctx.reg('second', (d: any) => { d = unwrapVbaDefaultValue(d); return d === vbaNull ? vbaNull : parseVbaDate(d).getSeconds(); }, [{ name: 'Time' }]);
    ctx.reg('dateserial', (y: any, m: any, d: any) => {
        if (y === vbaNull || m === vbaNull || d === vbaNull) return vbaNull;
        const rawYear = coerceDateNumber(y);
        let year = rawYear;
        // VBA spec §6.1.2.4.1.4: 0-29 → 2000-2029, 30-99 → 1930-1999
        if (year >= 0 && year <= 29) year += 2000;
        else if (year >= 30 && year <= 99) year += 1900;
        const mm = coerceDateNumber(m) - 1, dd = coerceDateNumber(d);
        // 月/日は範囲外なら繰り上げ・繰り下げされる（実 VBA 差分で裁定: Month=13→翌年1月）。
        if (rawYear >= 0 && rawYear <= 99) {
            // 0-99 年の特殊レンジのみ: JS の Date コンストラクターが 0-99 年を 1900 年台へ
            // 自動オフセットしてしまうため、安全な基準年でいったん繰り上げ後の年差分だけを
            // 計算し、目的の年へ適用する（基準年の閏年有無で 2/29 絡みの端数がズレうる
            // 稀なエッジケースだが、0-99 年指定自体が非常に稀なので許容する）。
            const baseYear = 2000;
            const date = new Date(baseYear, mm, dd);
            ensureVbaDate(date);
            const yearOffset = date.getFullYear() - baseYear;
            date.setFullYear(year + yearOffset);
            return new VbaDate(toVbaDate(date));
        }
        // 通常の年（100 以上）は JS の Date コンストラクターがそのまま正しく
        // 繰り上げ・繰り下げてくれる（閏年判定も年に対して正確に行われる）
        const date = new Date(year, mm, dd);
        ensureVbaDate(date);
        return new VbaDate(toVbaDate(date));
    }, [{ name: 'Year' }, { name: 'Month' }, { name: 'Day' }]);
    ctx.reg('timeserial', (h: any, n: any, s: any) => {
        if (h === vbaNull || n === vbaNull || s === vbaNull) return vbaNull;
        const hour = coerceDateNumber(h), minute = coerceDateNumber(n), second = coerceDateNumber(s);
        if (![hour, minute, second].every(Number.isFinite)) ctx.throwError(VbaErrorCode.OVERFLOW, 'Overflow');
        const date = new Date(1899, 11, 30, hour, minute, second);
        ensureVbaDate(date);
        return new VbaDate(toVbaDate(date));
    }, [{ name: 'Hour' }, { name: 'Minute' }, { name: 'Second' }]);
    ctx.reg('weekday', (d: any, firstdayofweek: any = 1) => {
        d = unwrapVbaDefaultValue(d);
        if (d === vbaNull) return vbaNull;
        if (firstdayofweek === vbaNull) ctx.throwError(VbaErrorCode.INVALID_PROCEDURE_CALL, 'Invalid procedure call or argument');
        const dayOfWeek = parseVbaDate(d).getDay(); // 0=Sun
        let fdow = validateFirstDayOfWeek(firstdayofweek);
        if (fdow === 0) fdow = 1; // vbUseSystemDayOfWeek → treat as vbSunday
        const weekStart = fdow <= 1 ? 0 : fdow - 1; // convert VBA 1-based to JS 0-based
        return ((dayOfWeek - weekStart + 7) % 7) + 1;
    }, [{ name: 'Date' }, { name: 'FirstDayOfWeek', optional: true }]);
    ctx.reg('dateadd', (interval: any, number: any, date: any) => {
        date = unwrapVbaDefaultValue(date);
        if (date === vbaNull || number === vbaNull) return vbaNull;
        rejectNullArgument(ctx, interval);
        const d = parseVbaDate(date);
        const n = coerceDateNumber(number);
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
        ensureVbaDate(d);
        return new VbaDate(toVbaDate(d));
    }, [{ name: 'Interval' }, { name: 'Number' }, { name: 'Date' }]);
    ctx.reg('datediff', (interval: any, date1: any, date2: any, firstdayofweek: any = 1, _firstweekofyear: any = 1) => {
        date1 = unwrapVbaDefaultValue(date1);
        date2 = unwrapVbaDefaultValue(date2);
        if (date1 === vbaNull || date2 === vbaNull) return vbaNull;
        rejectNullArgument(ctx, interval);
        const d1 = parseVbaDate(date1);
        const d2 = parseVbaDate(date2);
        const intv = String(interval).toLowerCase();
        const calendarDayDiff = (): number => {
            const day1 = Date.UTC(d1.getFullYear(), d1.getMonth(), d1.getDate());
            const day2 = Date.UTC(d2.getFullYear(), d2.getMonth(), d2.getDate());
            return Math.round((day2 - day1) / 86400000);
        };
        let fdow = validateFirstDayOfWeek(firstdayofweek);
        validateFirstWeekOfYear(_firstweekofyear);
        if (intv === 'yyyy') return d2.getFullYear() - d1.getFullYear();
        else if (intv === 'q') return (d2.getFullYear() - d1.getFullYear()) * 4 + Math.floor(d2.getMonth() / 3) - Math.floor(d1.getMonth() / 3);
        else if (intv === 'm') return (d2.getFullYear() - d1.getFullYear()) * 12 + d2.getMonth() - d1.getMonth();
        else if (intv === 'y' || intv === 'd') return calendarDayDiff();
        else if (intv === 'w') {
            // Weekday counts are calendar-day counts: DateDiff excludes date1
            // and includes date2, so time-of-day must not round a six-day span
            // up to a seventh day. This is distinct from the `ww` boundary
            // calculation below.
            const days = calendarDayDiff();
            return days >= 0 ? Math.floor(days / 7) : -Math.floor(-days / 7);
        }
        else if (intv === 'ww') {
            // Week count depends on firstdayofweek
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
        else if (intv === 'h' || intv === 'n' || intv === 's') {
            // DateDiff counts boundaries of the requested unit, rather than
            // rounding the elapsed duration.  Project both endpoints onto
            // the unit boundary before taking their signed difference.
            const boundary = (date: Date): Date => {
                const value = new Date(date);
                if (intv === 'h') value.setMinutes(0, 0, 0);
                else if (intv === 'n') value.setSeconds(0, 0);
                else value.setMilliseconds(0);
                return value;
            };
            const unitMs = intv === 'h' ? 3600000 : intv === 'n' ? 60000 : 1000;
            return Math.round((boundary(d2).getTime() - boundary(d1).getTime()) / unitMs);
        }
        else ctx.throwError(VbaErrorCode.INVALID_PROCEDURE_CALL, 'Invalid procedure call or argument');
    }, [
        { name: 'Interval' }, { name: 'Date1' }, { name: 'Date2' },
        { name: 'FirstDayOfWeek', optional: true },
        { name: 'FirstWeekOfYear', optional: true },
    ]);
    ctx.reg('datepart', (interval: any, date: any, firstdayofweek: any = 1, firstweekofyear: any = 1) => {
        date = unwrapVbaDefaultValue(date);
        if (date === vbaNull) return vbaNull;
        rejectNullArgument(ctx, interval);
        const d = parseVbaDate(date);
        const intv = String(interval).toLowerCase();
        // firstdayofweek: 1=Sunday(default), 2=Monday, ..., 7=Saturday; 0=system default(treat as 1)
        const fdow = validateFirstDayOfWeek(firstdayofweek);
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
            const fwoy = validateFirstWeekOfYear(firstweekofyear);
            return vbaWeekNumber(d, fdow, fwoy === 0 ? 1 : fwoy);
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
        val = unwrapVbaDefaultValue(val);
        if (val === vbaNull) return vbaNull;
        const d = parseVbaDate(val);
        if (typeof val === 'string') {
            // JavaScript Date normalizes invalid calendar dates (for example,
            // 2024/02/30 → 2024/03/01). DateValue must reject that input.
            const text = val.trim();
            const parts = /^(\d{4})[\/-](\d{1,2})[\/-](\d{1,2})(?:\s|$)/.exec(text);
            if (parts) {
                const year = Number(parts[1]), month = Number(parts[2]), day = Number(parts[3]);
                if (d.getFullYear() !== year || d.getMonth() + 1 !== month || d.getDate() !== day) {
                    ctx.throwError(VbaErrorCode.TYPE_MISMATCH, 'Type mismatch: invalid date');
                }
            }
        }
        return new VbaDate(Math.floor(toVbaDate(d)));
    }, [{ name: 'Date' }]);
    ctx.reg('timevalue', (val: any) => {
        val = unwrapVbaDefaultValue(val);
        if (val === vbaNull) return vbaNull;
        const d = parseVbaDate(val);
        const serial = toVbaDate(d);
        return new VbaDate(serial - Math.floor(serial));
    }, [{ name: 'Time' }]);
    ctx.reg('monthname', (month: any, abbreviate: any = vbaFalse) => {
        if (month === vbaNull) return vbaNull;
        const m = coerceDateNumber(month);
        if (m < 1 || m > 12) ctx.throwError(VbaErrorCode.INVALID_PROCEDURE_CALL, "Invalid procedure call or argument");
        const names = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
        const abbrs = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
        return ctx.isTrue(abbreviate) ? abbrs[m - 1] : names[m - 1];
    }, [{ name: 'Month' }, { name: 'Abbreviate', optional: true }]);
    ctx.reg('weekdayname', (weekday: any, abbreviate: any = vbaFalse, firstdayofweek: any = 1) => {
        if (weekday === vbaNull) return vbaNull;
        const w = coerceDateNumber(weekday);
        let first = validateFirstDayOfWeek(firstdayofweek);
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
    ctx.reg('shell', (cmd: any, style: any = 1) => {
        rejectNullArgument(ctx, style);
        ctx.print(`[SHELL] ${vbaToString(cmd ?? '')} (Style: ${ctx.toVbaNumber(style)})`);
        return 1;
    }, [
        { name: 'PathName' },
        { name: 'WindowStyle', optional: true },
    ]);
    ctx.reg('msgbox', (msg: any, _buttons: any = 0, _title: any = "", _helpFile?: string, _context?: number) => {
        rejectNullArgument(ctx, _buttons);
        rejectNullArgument(ctx, _title);
        rejectNullArgument(ctx, _helpFile);
        rejectNullArgument(ctx, _context);
        const titleStr = vbaToString(_title ?? '');
        const title = titleStr ? ` ${titleStr}:` : '';
        ctx.print(`[MSGBOX]${title} ${vbaToString(msg ?? '')}`);
        return 1;
    }, [
        { name: 'Prompt' },
        { name: 'Buttons', optional: true },
        { name: 'Title', optional: true },
        { name: 'HelpFile', optional: true },
        { name: 'Context', optional: true },
    ]);
    ctx.reg('inputbox', (prompt: any, _title: any = "", def: any = "", _xpos?: number, _ypos?: number, _helpFile?: string, _context?: number) => {
        rejectNullArgument(ctx, _title);
        rejectNullArgument(ctx, def);
        rejectNullArgument(ctx, _xpos);
        rejectNullArgument(ctx, _ypos);
        rejectNullArgument(ctx, _helpFile);
        rejectNullArgument(ctx, _context);
        ctx.print(`[INPUTBOX] ${vbaToString(prompt ?? '')}`);
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
    ctx.reg('appactivate', (title: any, _wait?: boolean) => {
        rejectNullArgument(ctx, _wait);
        ctx.print(`[APPACTIVATE] ${vbaToString(title ?? '')}`);
    }, [
        { name: 'Title' },
        { name: 'Wait', optional: true },
    ]);
    ctx.reg('sendkeys', (keys: any, _wait?: boolean) => {
        rejectNullArgument(ctx, _wait);
        ctx.print(`[SENDKEYS] ${vbaToString(keys ?? '')}`);
    }, [
        { name: 'String' },
        { name: 'Wait', optional: true },
    ]);
    ctx.reg('doevents', () => 0, []);
}

// ---------------------------------------------------------------------------
// Financial functions — FV, PV, PMT, NPer, Rate, SLN, SYD, DDB, IRR, etc.
// ---------------------------------------------------------------------------

export function registerFinancialFunctions(ctx: StdlibCtx): void {
    const invalidFinancialArg = (): never => {
        ctx.throwError(VbaErrorCode.INVALID_PROCEDURE_CALL, 'Invalid procedure call or argument');
    };
    const finiteResult = (value: number): number => {
        if (!Number.isFinite(value)) ctx.throwError(VbaErrorCode.OVERFLOW, 'Overflow');
        return value;
    };
    const scaledLinearCombination = (
        terms: Array<[value: number, factor: number]>,
        divisor = 1,
    ): number => {
        // A zero-valued term must not hide an invalid rate factor.  Excel
        // reports overflow for FV/PMT-style calculations as soon as an
        // intermediate factor is non-finite, even when every cash-flow value
        // is zero (BUG-00539).
        if (terms.some(([, factor]) => !Number.isFinite(factor))) {
            return finiteResult(Number.NaN);
        }
        const scale = Math.max(...terms.map(([value]) => Math.abs(value)));
        if (scale === 0) return 0;
        if (!Number.isFinite(scale) || !Number.isFinite(divisor) || divisor === 0) {
            return finiteResult(Number.NaN);
        }
        const normalized = terms.reduce((sum, [value, factor]) => sum + (value / scale) * factor, 0);
        return finiteResult(scale * (normalized / divisor));
    };
    const toNum = (val: any): number => {
        if (val === vbaNull) ctx.throwError(VbaErrorCode.INVALID_USE_OF_NULL, 'Invalid use of Null');
        return ctx.toVbaNumber(val);
    };
    const toPaymentType = (val: any): number => {
        const type = toNum(val);
        if (type !== 0 && type !== 1) invalidFinancialArg();
        return type;
    };
    const getRateFactor = (rate: number, nper: number) => {
        if (rate === 0) return nper;
        return (Math.pow(1 + rate, nper) - 1) / rate;
    };
    const getRateFactorAndDerivative = (rate: number, nper: number) => {
        if (rate === 0) {
            return { factor: nper, derivative: nper * (nper - 1) / 2 };
        }
        const p1 = Math.pow(1 + rate, nper);
        return {
            factor: (p1 - 1) / rate,
            derivative: (nper * Math.pow(1 + rate, nper - 1) * rate - (p1 - 1)) / (rate * rate),
        };
    };
    const toCashFlowValues = (values: any): number[] => {
        if (!Array.isArray(values)) ctx.throwError(VbaErrorCode.TYPE_MISMATCH, "Type mismatch");
        const base: number = (values as any).vbaBase ?? 0;
        return Array.from({ length: values.length - base }, (_, i) => toNum(values[base + i]));
    };
    const requireMixedCashFlows = (values: number[]): number[] => {
        if (values.length < 2 || !values.some((value) => value < 0) || !values.some((value) => value > 0)) {
            invalidFinancialArg();
        }
        return values;
    };
    ctx.reg('fv', (rate: any, nper: any, pmt: any, pv: any = 0, type: any = 0) => {
        const r = toNum(rate), n = toNum(nper), p = toNum(pmt), v = toNum(pv), t = toPaymentType(type);
        const factor = getRateFactor(r, n);
        return -scaledLinearCombination([
            [v, Math.pow(1 + r, n)],
            [p, (1 + r * t) * factor],
        ]);
    }, [
        { name: 'Rate' }, { name: 'NPer' }, { name: 'Pmt' },
        { name: 'PV', optional: true }, { name: 'Type', optional: true },
    ]);
    ctx.reg('pv', (rate: any, nper: any, pmt: any, fv: any = 0, type: any = 0) => {
        const r = toNum(rate), n = toNum(nper), p = toNum(pmt), f = toNum(fv), t = toPaymentType(type);
        if (r === 0) return finiteResult(-(f + p * n));
        const p1 = Math.pow(1 + r, n);
        if (!Number.isFinite(p1) || p1 === 0) invalidFinancialArg();
        return -scaledLinearCombination([
            [f, 1],
            [p, (1 + r * t) * ((p1 - 1) / r)],
        ], p1);
    }, [
        { name: 'Rate' }, { name: 'NPer' }, { name: 'Pmt' },
        { name: 'FV', optional: true }, { name: 'Type', optional: true },
    ]);
    ctx.reg('pmt', (rate: any, nper: any, pv: any, fv: any = 0, type: any = 0) => {
        const r = toNum(rate), n = toNum(nper), v = toNum(pv), f = toNum(fv), t = toPaymentType(type);
        if (!Number.isFinite(n) || n <= 0) {
            ctx.throwError(VbaErrorCode.INVALID_PROCEDURE_CALL, 'Invalid procedure call or argument');
        }
        if (r === 0) return finiteResult(-(v + f) / n);
        const p1 = Math.pow(1 + r, n);
        return -scaledLinearCombination([
            [v, p1],
            [f, 1],
        ], (1 + r * t) * ((p1 - 1) / r));
    }, [
        { name: 'Rate' }, { name: 'NPer' }, { name: 'PV' },
        { name: 'FV', optional: true }, { name: 'Type', optional: true },
    ]);
    ctx.reg('nper', (rate: any, pmt: any, pv: any, fv: any = 0, type: any = 0) => {
        const r = toNum(rate), p = toNum(pmt), v = toNum(pv), f = toNum(fv), t = toPaymentType(type);
        // Excel classifies the maximum-magnitude, opposite-sign cancellation
        // boundary as overflow before returning the mathematically finite
        // logarithmic quotient (BUG-00538 / XL-217).
        if (Math.abs(p) >= 1e308 && Math.abs(v) >= 1e308 && p * v < 0) {
            ctx.throwError(VbaErrorCode.OVERFLOW, 'Overflow');
        }
        if (r === 0) {
            if (p === 0) invalidFinancialArg();
            const result = -(v + f) / p;
            return Number.isFinite(result) ? result : invalidFinancialArg();
        }
        const numerator = p * (1 + r * t) - f * r;
        const denominator = p * (1 + r * t) + v * r;
        const ratio = numerator / denominator;
        if (1 + r <= 0 || !Number.isFinite(ratio) || ratio <= 0) invalidFinancialArg();
        const result = Math.log(ratio) / Math.log(1 + r);
        return Number.isFinite(result) ? result : invalidFinancialArg();
    }, [
        { name: 'Rate' }, { name: 'Pmt' }, { name: 'PV' },
        { name: 'FV', optional: true }, { name: 'Type', optional: true },
    ]);
    ctx.reg('rate', (nper: any, pmt: any, pv: any, fv: any = 0, type: any = 0, guess: any = 0.1) => {
        let r = toNum(guess);
        const n = toNum(nper), p = toNum(pmt), v = toNum(pv), f = toNum(fv), t = toPaymentType(type);
        if (!Number.isFinite(n) || n <= 0 || !Number.isFinite(r) || 1 + r <= 0) invalidFinancialArg();
        for (let i = 0; i < 20; i++) {
            const p1 = Math.pow(1 + r, n);
            const { factor, derivative } = getRateFactorAndDerivative(r, n);
            const f_r = v * p1 + p * (1 + r * t) * factor + f;
            const df_r = v * n * Math.pow(1 + r, n - 1) + p * (t * factor + (1 + r * t) * derivative);
            if (!Number.isFinite(f_r) || !Number.isFinite(df_r) || df_r === 0) invalidFinancialArg();
            if (Math.abs(f_r) < 1e-10) return r;
            const newR = r - f_r / df_r;
            if (!Number.isFinite(newR) || 1 + newR <= 0) invalidFinancialArg();
            if (Math.abs(newR - r) < 1e-10) return newR;
            r = newR;
        }
        invalidFinancialArg();
    }, [
        { name: 'NPer' }, { name: 'Pmt' }, { name: 'PV' },
        { name: 'FV', optional: true }, { name: 'Type', optional: true }, { name: 'Guess', optional: true },
    ]);
    ctx.reg('sln', (cost: any, salvage: any, life: any) => {
        const c = toNum(cost), s = toNum(salvage);
        const l = toNum(life);
        if (!Number.isFinite(l) || l <= 0) invalidFinancialArg();
        // Divide before subtracting so opposite-signed values near the
        // Double limit do not overflow an otherwise finite VBA result.
        return finiteResult(c / l - s / l);
    }, [{ name: 'Cost' }, { name: 'Salvage' }, { name: 'Life' }]);
    ctx.reg('syd', (cost: any, salvage: any, life: any, period: any) => {
        const c = toNum(cost), s = toNum(salvage), l = toNum(life), p = toNum(period);
        if (!Number.isFinite(l) || !Number.isFinite(p) || l <= 0 || p <= 0 || p > l || s < 0) invalidFinancialArg();
        // Apply the denominator before the numerator multiplier so a
        // mathematically finite result does not overflow an intermediate.
        const factor = ((l - p + 1) * 2) / (l * (l + 1));
        return finiteResult((c - s) * factor);
    }, [{ name: 'Cost' }, { name: 'Salvage' }, { name: 'Life' }, { name: 'Period' }]);
    ctx.reg('ddb', (cost: any, salvage: any, life: any, period: any, factor: any = 2) => {
        const c = toNum(cost), s = toNum(salvage), l = toNum(life), p = toNum(period), f = toNum(factor);
        if (!Number.isFinite(l) || !Number.isFinite(p) || !Number.isFinite(f) || l <= 0 || p <= 0 || p > l || f <= 0) invalidFinancialArg();
        let book = c;
        let dep = 0;
        for (let i = 1; i <= p; i++) {
            const remaining = book - s;
            const scheduled = book * (f / l);
            if (!Number.isFinite(remaining) || !Number.isFinite(scheduled)) invalidFinancialArg();
            dep = Math.min(scheduled, Math.max(0, remaining));
            if (!Number.isFinite(dep)) invalidFinancialArg();
            book -= dep;
            if (!Number.isFinite(book)) invalidFinancialArg();
        }
        return finiteResult(dep);
    }, [
        { name: 'Cost' }, { name: 'Salvage' }, { name: 'Life' }, { name: 'Period' },
        { name: 'Factor', optional: true },
    ]);
    ctx.reg('irr', (values: any, guess: any = 0.1) => {
        const v = requireMixedCashFlows(toCashFlowValues(values));
        let r = toNum(guess);
        for (let i = 0; i < 100; i++) {
            let npv = 0, dnpv = 0;
            for (let t = 0; t < v.length; t++) {
                const p1 = Math.pow(1 + r, t);
                npv += v[t] / p1;
                if (t > 0) dnpv -= t * v[t] / (p1 * (1 + r));
            }
            const newR = r - npv / dnpv;
            if (!Number.isFinite(newR)) {
                ctx.throwError(VbaErrorCode.INVALID_PROCEDURE_CALL, 'Invalid procedure call or argument');
            }
            if (Math.abs(newR - r) < 1e-10) return newR;
            r = newR;
        }
        if (!Number.isFinite(r)) {
            ctx.throwError(VbaErrorCode.INVALID_PROCEDURE_CALL, 'Invalid procedure call or argument');
        }
        return r;
    }, [{ name: 'ValueArray', isArray: true }, { name: 'Guess', optional: true }]);
    ctx.reg('mirr', (values: any, finance_rate: any, reinvest_rate: any) => {
        const v = requireMixedCashFlows(toCashFlowValues(values));
        const fr = toNum(finance_rate), rr = toNum(reinvest_rate);
        // Excel rejects the exact zero discount denominator when a later
        // negative cash flow would be divided by (1 + FinanceRate)^t.
        if (fr === -1 && v.some((value, index) => index > 0 && value < 0)) {
            invalidFinancialArg();
        }
        // A zero reinvestment growth factor has no defined future-value
        // denominator. Report VBA's argument error before the calculation
        // reaches the generic non-finite-result path.
        if (1 + rr === 0) invalidFinancialArg();
        const n = v.length - 1;
        let npv_neg = 0, npv_pos = 0;
        for (let t = 0; t < v.length; t++) {
            if (v[t] < 0) npv_neg += v[t] / Math.pow(1 + fr, t);
            else npv_pos += v[t] / Math.pow(1 + rr, t);
        }
        const tv = npv_pos * Math.pow(1 + rr, n);
        const result = Math.pow(-tv / npv_neg, 1 / n) - 1;
        if (!Number.isFinite(result)) {
            ctx.throwError(VbaErrorCode.INVALID_PROCEDURE_CALL, 'Invalid procedure call or argument');
        }
        return finiteResult(result);
    }, [{ name: 'ValueArray', isArray: true }, { name: 'FinanceRate' }, { name: 'ReinvestRate' }]);
    ctx.reg('npv', (rate: any, values: any) => {
        if (!Array.isArray(values)) ctx.throwError(VbaErrorCode.TYPE_MISMATCH, "Type mismatch");
        const r = toNum(rate);
        if (!Number.isFinite(r) || 1 + r <= 0) invalidFinancialArg();
        const v = requireMixedCashFlows(toCashFlowValues(values));
        let result = 0;
        let period = 1;
        for (const value of v) {
            result += value / Math.pow(1 + r, period++);
        }
        return finiteResult(result);
    }, [{ name: 'Rate' }, { name: 'ValueArray', isArray: true }]);
    ctx.reg('ipmt', (rate: any, per: any, nper: any, pv: any, fv: any = 0, type: any = 0) => {
        const r = toNum(rate), p = toNum(per), n = toNum(nper), v = toNum(pv), f = toNum(fv), t = toPaymentType(type);
        if (!Number.isInteger(p) || p < 1 || p > n) {
            ctx.throwError(VbaErrorCode.INVALID_PROCEDURE_CALL, 'Invalid procedure call or argument');
        }
        const pmt = ctx.toVbaNumber(ctx.envGet('pmt')(r, n, v, f, t));
        let ipmt: number;
        if (p === 1) {
            ipmt = t === 1 ? 0 : -v * r;
        } else if (t === 1) {
            const fvBeforePreviousPayment = ctx.toVbaNumber(ctx.envGet('fv')(r, p - 2, pmt, v, t));
            ipmt = (fvBeforePreviousPayment - pmt) * r;
        } else {
            const fv_prev = ctx.toVbaNumber(ctx.envGet('fv')(r, p - 1, pmt, v, t));
            ipmt = fv_prev * r;
        }
        return finiteResult(ipmt);
    }, [
        { name: 'Rate' }, { name: 'Per' }, { name: 'NPer' }, { name: 'PV' },
        { name: 'FV', optional: true }, { name: 'Type', optional: true },
    ]);
    ctx.reg('ppmt', (rate: any, per: any, nper: any, pv: any, fv: any = 0, type: any = 0) => {
        const r = toNum(rate), p = toNum(per), n = toNum(nper), v = toNum(pv), f = toNum(fv), t = toPaymentType(type);
        const pmt = ctx.toVbaNumber(ctx.envGet('pmt')(r, n, v, f, t));
        const ipmt = ctx.toVbaNumber(ctx.envGet('ipmt')(r, p, n, v, f, t));
        return finiteResult(pmt - ipmt);
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
    const errFunc = (n?: any) => {
        if (n === vbaNull) ctx.throwError(VbaErrorCode.INVALID_USE_OF_NULL, 'Invalid use of Null');
        const code = n === undefined ? ctx.errNum() : ctx.round(ctx.toVbaNumber(n));
        return errorMessages[code] || "Application-defined or object-defined error";
    };
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

    ctx.reg('environ', (k: any) => {
        rejectNullArgument(ctx, k);
        return ctx.getEnv(k);
    }, [{ name: 'EnvString' }], ['$']);
    ctx.reg('rgb', (r: any, g: any, b: any) => {
        if (r === vbaNull || g === vbaNull || b === vbaNull) ctx.throwError(VbaErrorCode.INVALID_USE_OF_NULL, 'Invalid use of Null');
        const clamp = (n: number) => Math.min(255, Math.max(0, ctx.round(n)));
        return clamp(ctx.toVbaNumber(r)) + clamp(ctx.toVbaNumber(g)) * 256 + clamp(ctx.toVbaNumber(b)) * 65536;
    }, [{ name: 'Red' }, { name: 'Green' }, { name: 'Blue' }]);
    const qbColorTable = [0, 8388608, 32768, 8421376, 128, 8388736, 32896, 12632256,
                          8421504, 16711680, 65280, 16776960, 255, 16711935, 65535, 16777215];
    ctx.reg('qbcolor', (c: any) => {
        if (c === vbaNull) ctx.throwError(VbaErrorCode.INVALID_USE_OF_NULL, 'Invalid use of Null');
        const idx = ctx.round(ctx.toVbaNumber(c));
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
    ctx.reg('createobject', (id: any, _serverName?: string) => ctx.createObj(vbaToString(id ?? '')), [
        { name: 'Class' },
        { name: 'ServerName', optional: true },
    ]);
    const getObjectFunc = (pathname?: any, classId?: any) => {
        const pathStr = pathname === undefined || pathname === vbaMissing ? '' : vbaToString(pathname);
        if (pathStr) {
            return { __vbaTypeName__: 'Object', Path: pathStr, Name: pathStr.split(/[\\\/]/).pop() };
        }
        if (classId !== undefined && classId !== vbaMissing) return ctx.createObj(vbaToString(classId));
        return vbaNothing;
    };
    ctx.reg('getobject', getObjectFunc, [
        { name: 'PathName', optional: true },
        { name: 'Class', optional: true },
    ]);
    ctx.reg('iif', (c: any, t: any, f: any) => ctx.isTrue(c) ? t : f, [
        { name: 'Expr' }, { name: 'TruePart' }, { name: 'FalsePart' },
    ]);
    const coerceIndex = (value: any): number => ctx.round(ctx.toVbaNumber(value));
    ctx.reg('choose', (i: any, ...c: any[]) => {
        if (i === vbaNull) return vbaNull;
        const idx = coerceIndex(i);
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
        (a as any).__vbaArrayReturn__ = true;
        (a as any).__vbaArrayReturnType__ = 'Variant';
        return a;
    }, [
        { name: 'Arglist', isParamArray: true },
    ]);
    ctx.reg('lbound', (a: any, dim: any = 1) => {
        if (!Array.isArray(a)) ctx.throwError(VbaErrorCode.SUBSCRIPT_OUT_OF_RANGE, "Subscript out of range");
        if ((a as any).__vbaErased__) ctx.throwError(VbaErrorCode.SUBSCRIPT_OUT_OF_RANGE, "Subscript out of range");
        if (dim === vbaNull) ctx.throwError(VbaErrorCode.INVALID_USE_OF_NULL, "Invalid use of Null");
        const dimIndex = coerceIndex(dim) - 1;
        if ((a as any).__vbaDimensions__) {
            if (dimIndex < 0 || dimIndex >= (a as any).__vbaDimensions__.length) {
                ctx.throwError(VbaErrorCode.SUBSCRIPT_OUT_OF_RANGE, "Subscript out of range");
            }
            return (a as any).__vbaDimensions__[dimIndex].lower;
        }
        if (dimIndex < 0 || dimIndex > 0) ctx.throwError(VbaErrorCode.SUBSCRIPT_OUT_OF_RANGE, "Subscript out of range");
        return (a as any).vbaBase || 0;
    }, [{ name: 'ArrayName' }, { name: 'Dimension', optional: true }]);
    ctx.reg('ubound', (a: any, dim: any = 1) => {
        if (!Array.isArray(a)) ctx.throwError(VbaErrorCode.SUBSCRIPT_OUT_OF_RANGE, "Subscript out of range");
        if ((a as any).__vbaErased__) ctx.throwError(VbaErrorCode.SUBSCRIPT_OUT_OF_RANGE, "Subscript out of range");
        if (dim === vbaNull) ctx.throwError(VbaErrorCode.INVALID_USE_OF_NULL, "Invalid use of Null");
        const dimIndex = coerceIndex(dim) - 1;
        if ((a as any).__vbaDimensions__) {
            if (dimIndex < 0 || dimIndex >= (a as any).__vbaDimensions__.length) {
                ctx.throwError(VbaErrorCode.SUBSCRIPT_OUT_OF_RANGE, "Subscript out of range");
            }
            return (a as any).__vbaDimensions__[dimIndex].upper;
        }
        if (dimIndex < 0 || dimIndex > 0) ctx.throwError(VbaErrorCode.SUBSCRIPT_OUT_OF_RANGE, "Subscript out of range");
        // a.length already includes vbaBase filler slots (added by Array() for Option Base 1),
        // so UBound = a.length - 1 (not vbaBase + a.length - 1).
        return a.length - 1;
    }, [{ name: 'ArrayName' }, { name: 'Dimension', optional: true }]);
}

// ---------------------------------------------------------------------------
// Registry functions — SaveSetting, GetSetting, GetAllSettings, DeleteSetting
// ---------------------------------------------------------------------------

export function registerRegistryFunctions(ctx: StdlibCtx): void {
    ctx.reg('savesetting', (app: any, sec: any, key: any, val: any) => {
        const appName = app as string;
        const section = sec as string;
        const settingKey = key as string;
        if (!ctx.registry[appName]) ctx.registry[appName] = {};
        if (!ctx.registry[appName][section]) ctx.registry[appName][section] = {};
        ctx.registry[appName][section][settingKey] = val as string;
    }, [
        { name: 'AppName', coerce: 'string' }, { name: 'Section', coerce: 'string' },
        { name: 'Key', coerce: 'string' }, { name: 'Setting', coerce: 'string' },
    ]);
    ctx.reg('getsetting', (app: any, sec: any, key: any, def: any = "") => {
        return (ctx.registry[vbaToString(app)]?.[vbaToString(sec)]?.[vbaToString(key)]) ?? vbaToString(def);
    }, [
        { name: 'AppName', coerce: 'string' }, { name: 'Section', coerce: 'string' },
        { name: 'Key', coerce: 'string' }, { name: 'Default', optional: true },
    ]);
    ctx.reg('getallsettings', (app: any, sec: any) => {
        const s = ctx.registry[app as string]?.[sec as string];
        if (!s) return vbaEmpty;
        const res = Object.entries(s).map(([k, v]) => [k, v]);
        (res as any).__vbaDimensions__ = [
            { lower: 0, upper: res.length - 1 },
            { lower: 0, upper: 1 },
        ];
        (res as any).vbaBase = 0;
        return res;
    }, [{ name: 'AppName', coerce: 'string' }, { name: 'Section', coerce: 'string' }]);
    ctx.reg('deletesetting', (app: any, sec?: any, key?: any) => {
        const appName = app as string;
        const section = sec as string | undefined;
        const settingKey = key as string | undefined;
        if (!ctx.registry[appName]) ctx.throwError(VbaErrorCode.INVALID_PROCEDURE_CALL, 'Invalid procedure call or argument');
        if (section === undefined) { delete ctx.registry[appName]; return; }
        if (!ctx.registry[appName][section]) ctx.throwError(VbaErrorCode.INVALID_PROCEDURE_CALL, 'Invalid procedure call or argument');
        if (settingKey === undefined) { delete ctx.registry[appName][section]; return; }
        if (!(settingKey in ctx.registry[appName][section])) ctx.throwError(VbaErrorCode.INVALID_PROCEDURE_CALL, 'Invalid procedure call or argument');
        delete ctx.registry[appName][section][settingKey];
    }, [
        { name: 'AppName', coerce: 'string' },
        { name: 'Section', optional: true, coerce: 'string' },
        { name: 'Key', optional: true, coerce: 'string' },
    ]);
}
