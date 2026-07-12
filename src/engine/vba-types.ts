/**
 * VBA runtime value types and sentinel constants.
 * Extracted here to avoid circular imports between evaluator.ts and coerce.ts.
 */
import { VbaErrorCode, throwVbaError } from './vba-errors';

/**
 * VBA Boolean wrapper. Only two instances exist: vbaTrue (-1) and vbaFalse (0).
 * Use VbaBoolean.from(n) to obtain a singleton from a number.
 */
export class VbaBoolean {
    public readonly __isVbaBoolean__ = true;
    private constructor(public readonly value: -1 | 0) {}
    valueOf() { return this.value; }
    toString() { return this.value === -1 ? 'True' : 'False'; }

    static from(n: number): VbaBoolean {
        return n !== 0 ? vbaTrue : vbaFalse;
    }

    /** @internal */
    static _createSingleton(value: -1 | 0): VbaBoolean {
        return new VbaBoolean(value);
    }
}

// VBA date serial: days since local midnight 1899-12-30
const VBA_EPOCH = new Date(1899, 11, 30);
const MS_PER_DAY = 86400000;

export const toVbaDate = (d: Date): number =>
    (d.getTime() - VBA_EPOCH.getTime()) / MS_PER_DAY;

export const fromVbaDate = (serial: number): Date => {
    const ms = Math.round(serial * MS_PER_DAY);
    return new Date(VBA_EPOCH.getTime() + ms);
};

export const parseVbaDate = (val: any): Date => {
    if (val === null || val === undefined) throwVbaError(VbaErrorCode.TYPE_MISMATCH);
    if (val instanceof VbaDate || (val && val.__isVbaDate__)) return fromVbaDate(val.value);
    if (typeof val === 'number') return fromVbaDate(val);

    let str = String(val);
    if (/^\d{1,2}:\d{1,2}(:\d{1,2})?$/.test(str)) {
        str = "1899/12/30 " + str;
    }
    const d = new Date(str);
    if (isNaN(d.getTime())) throwVbaError(VbaErrorCode.TYPE_MISMATCH, `Type mismatch: '${val}'`);
    return new Date(d.getFullYear(), d.getMonth(), d.getDate(),
        d.getHours(), d.getMinutes(), d.getSeconds(), d.getMilliseconds());
};

export class VbaDate {
    public __isVbaDate__ = true;
    constructor(public value: number) {}
    valueOf() { return this.value; }
    toString() {
        const d = fromVbaDate(this.value);
        const datePart = [
            String(d.getFullYear()),
            String(d.getMonth() + 1).padStart(2, '0'),
            String(d.getDate()).padStart(2, '0')
        ].join("/");
        const hasFraction = Math.abs(this.value - Math.round(this.value)) > 1e-10;
        if (!hasFraction) return datePart;
        const timePart = [
            String(d.getHours()).padStart(2, '0'),
            String(d.getMinutes()).padStart(2, '0'),
            String(d.getSeconds()).padStart(2, '0')
        ].join(':');
        return Math.round(this.value) === 0 ? timePart : datePart + ' ' + timePart;
    }
}

/** Expand JS exponential notation "1e-7" to plain decimal "0.0000001". */
function expandExponential(s: string): string {
    const match = s.match(/^(-?)(\d+)(?:\.(\d*))?[eE]([+-]?\d+)$/);
    if (!match) return s;
    const sign = match[1];
    const integer = match[2];
    const frac = match[3] ?? '';
    const exp = parseInt(match[4]);
    const digits = integer + frac;
    const decimalPos = integer.length + exp;
    if (decimalPos >= digits.length) {
        return sign + digits.padEnd(decimalPos, '0');
    } else if (decimalPos <= 0) {
        return sign + '0.' + '0'.repeat(-decimalPos) + digits;
    } else {
        return sign + digits.slice(0, decimalPos) + '.' + digits.slice(decimalPos);
    }
}

/**
 * VBA Decimal: 96-bit fixed-point, scale 0-28.
 * Internal: mantissa (signed bigint, |mantissa| ≤ 2^96-1) × 10^(-scale).
 */
export class VbaDecimal {
    readonly mantissa: bigint;
    readonly scale: number;

    static readonly MAX_MANTISSA = 79228162514264337593543950335n; // 2^96 - 1
    static readonly MAX_SCALE = 28;

    constructor(mantissa: bigint, scale: number) {
        if (scale < 0 || scale > VbaDecimal.MAX_SCALE) {
            throwVbaError(VbaErrorCode.OVERFLOW, 'Overflow');
        }
        const abs = mantissa < 0n ? -mantissa : mantissa;
        if (abs > VbaDecimal.MAX_MANTISSA) {
            throwVbaError(VbaErrorCode.OVERFLOW, 'Overflow');
        }
        this.mantissa = mantissa;
        this.scale = scale;
    }

    /** Approximate float value — backward compat; loses precision for >15-digit values. */
    get value(): number {
        return Number(this.mantissa) / Math.pow(10, this.scale);
    }

    toString(): string {
        const neg = this.mantissa < 0n;
        const abs = neg ? -this.mantissa : this.mantissa;
        if (this.scale === 0) return (neg ? '-' : '') + String(abs);
        const factor = 10n ** BigInt(this.scale);
        const intPart = abs / factor;
        const fracStr = String(abs % factor).padStart(this.scale, '0').replace(/0+$/, '');
        return (neg ? '-' : '') + String(intPart) + (fracStr ? '.' + fracStr : '');
    }

    /** Create from a JS number using its shortest string representation. */
    static fromNumber(val: number): VbaDecimal {
        if (!isFinite(val)) throwVbaError(VbaErrorCode.OVERFLOW, 'Overflow');
        return VbaDecimal._parse(String(val));
    }

    /** Create from a decimal string, preserving up to 28 decimal digits. */
    static fromString(s: string): VbaDecimal {
        return VbaDecimal._parse(s);
    }

    private static _parse(s: string): VbaDecimal {
        let ns = s.trim();
        if (/[eE]/.test(ns)) ns = expandExponential(ns);
        if (!/^-?(\d+\.?\d*|\.\d+)$/.test(ns)) {
            throwVbaError(VbaErrorCode.TYPE_MISMATCH, `Type mismatch: '${s}'`);
        }
        const neg = ns.startsWith('-');
        const abs = neg ? ns.slice(1) : ns;
        const dotIdx = abs.indexOf('.');
        const intStr = dotIdx === -1 ? abs : abs.slice(0, dotIdx);
        const rawFrac = dotIdx === -1 ? '' : abs.slice(dotIdx + 1);

        let scale: number;
        let absMantissa: bigint;

        if (rawFrac.length <= VbaDecimal.MAX_SCALE) {
            scale = rawFrac.length;
            absMantissa = BigInt(intStr || '0') * (10n ** BigInt(scale)) + BigInt(rawFrac || '0');
        } else {
            scale = VbaDecimal.MAX_SCALE;
            const truncFrac = rawFrac.slice(0, scale);
            const nextDigit = parseInt(rawFrac[scale], 10);
            const base = BigInt(intStr || '0') * (10n ** BigInt(scale)) + BigInt(truncFrac);
            if (nextDigit > 5) {
                absMantissa = base + 1n;
            } else if (nextDigit < 5) {
                absMantissa = base;
            } else {
                const hasMore = rawFrac.slice(scale + 1).split('').some(d => d !== '0');
                absMantissa = hasMore ? base + 1n : (base % 2n === 0n ? base : base + 1n);
            }
        }

        if (absMantissa > VbaDecimal.MAX_MANTISSA) {
            throwVbaError(VbaErrorCode.OVERFLOW, 'Overflow');
        }
        return new VbaDecimal(neg ? -absMantissa : absMantissa, scale);
    }
}

/**
 * Parse a decimal string to a BigInt scaled by 10^scale (with banker's rounding on excess digits).
 * Does not use float arithmetic internally.
 */
export function parseFixedPointString(s: string, scale: number): bigint {
    const neg = s.startsWith('-');
    const abs = neg ? s.slice(1) : s;
    const dotIdx = abs.indexOf('.');
    const intStr = dotIdx === -1 ? abs : abs.slice(0, dotIdx);
    const rawFrac = dotIdx === -1 ? '' : abs.slice(dotIdx + 1);

    let mantissa: bigint;
    if (rawFrac.length <= scale) {
        const frac = rawFrac.padEnd(scale, '0');
        mantissa = BigInt(intStr || '0') * (10n ** BigInt(scale)) + BigInt(frac || '0');
    } else {
        const truncFrac = rawFrac.slice(0, scale);
        const nextDigit = parseInt(rawFrac[scale], 10);
        const base = BigInt(intStr || '0') * (10n ** BigInt(scale)) + BigInt(truncFrac);
        if (nextDigit > 5) {
            mantissa = base + 1n;
        } else if (nextDigit < 5) {
            mantissa = base;
        } else {
            const hasMore = rawFrac.slice(scale + 1).split('').some(d => d !== '0');
            if (hasMore) {
                mantissa = base + 1n;
            } else {
                // Exactly half: banker's round to even (check the last kept digit)
                mantissa = base % 2n === 0n ? base : base + 1n;
            }
        }
    }
    return neg ? -mantissa : mantissa;
}

/**
 * BigInt integer division with banker's rounding (round half to even).
 */
export function bankersDivide(n: bigint, d: bigint): bigint {
    const q = n / d;
    const r = n - q * d;
    const absR2 = (r < 0n ? -r : r) * 2n;
    const absD = d < 0n ? -d : d;
    if (absR2 < absD) return q;
    if (absR2 > absD) return r < 0n ? q - 1n : q + 1n;
    // Exactly half: round to even
    if (q % 2n === 0n) return q;
    return r < 0n ? q - 1n : q + 1n;
}

/**
 * VBA Currency: 64-bit fixed-point, scale 10^-4.
 * Internal representation: integer × 10^-4 (e.g., 1.5 → 15000n).
 * No valueOf() to avoid silent float coercion.
 */
export class VbaCurrency {
    readonly internal: bigint;

    static readonly MIN = -9223372036854775808n;
    static readonly MAX =  9223372036854775807n;

    constructor(internal: bigint) {
        if (internal < VbaCurrency.MIN || internal > VbaCurrency.MAX) {
            throwVbaError(VbaErrorCode.OVERFLOW, 'Overflow');
        }
        this.internal = internal;
    }

    static fromNumber(val: number): VbaCurrency {
        // Apply banker's round to 4 decimal places via float, then stringify to avoid float mantissa errors
        const factor = 10000;
        const scaled = val * factor;
        const i = Math.floor(scaled);
        const f = scaled - i;
        let rounded: number;
        if (Math.abs(f - 0.5) < 1e-10) {
            rounded = (i % 2 === 0 ? i : i + 1) / factor;
        } else {
            rounded = Math.round(scaled) / factor;
        }
        // Use toFixed to get exact 4-decimal string, bypassing any remaining float representation issues
        const s = rounded.toFixed(4);
        return new VbaCurrency(parseFixedPointString(s, 4));
    }

    toString(): string {
        const neg = this.internal < 0n;
        const abs = neg ? -this.internal : this.internal;
        const intPart = abs / 10000n;
        const frac = String(abs % 10000n).padStart(4, '0').replace(/0+$/, '');
        return frac === '' ? `${neg ? '-' : ''}${intPart}` : `${neg ? '-' : ''}${intPart}.${frac}`;
    }
}

export class VbaErrorValue {
    constructor(public code: number) {}
    valueOf() { return this.code; }
    toString() { return `Error ${this.code}`; }
}

/** VBA Empty (uninitialized Variant). Internally null. */
export const vbaEmpty = null;

/** VBA Null — distinct from Empty. Propagates through expressions. */
export const vbaNull = Symbol('vbaNull');

/** VBA Nothing — unset object reference. */
export const vbaNothing = Symbol('vbaNothing');

/** Sentinel for missing optional argument. */
export const vbaMissing = Symbol('vbaMissing');

/**
 * Represents a VBA project or module name used as a namespace qualifier.
 * Exists as a pre-defined value in the environment so that bare uses
 * (e.g. `VarType(VBA)` or `VarType(Module1)`) can be detected and rejected.
 */
export class VbaNamespaceRef {
    constructor(
        public readonly name: string,
        public readonly kind: 'project' | 'module'
    ) {}
}

export const vbaTrue: VbaBoolean = VbaBoolean._createSingleton(-1);
export const vbaFalse: VbaBoolean = VbaBoolean._createSingleton(0);

export type VbaBooleanType = VbaBoolean;
export type VbaNumericType = 'Byte' | 'Integer' | 'Long' | 'Single' | 'Double' | 'Currency' | 'LongLong' | 'LongPtr';
export type VbaVarType = VbaNumericType | 'String' | 'Boolean' | 'Date' | 'Variant' | 'Object';

// ============================================================
// VBA オブジェクトインターフェース
//
// JS/TS で VBA オブジェクトとして振る舞うクラス（モック・外部スタブ等）が
// 実装すべきインターフェース群。evaluator.ts が各マーカーを参照して挙動を決定する。
// ============================================================

/**
 * VBA 型名を宣言するオブジェクト。
 * `TypeOf x Is Range` / `TypeName(x)` で正しい型名を返すために必要。
 * VBA 側から型判定されうるすべてのオブジェクト（モック・外部 COM スタブ等）が実装する。
 *
 * @example
 * class MockRange implements VbaType {
 *   readonly __vbaTypeName__ = 'Range';
 * }
 */
export interface VbaType {
    readonly __vbaTypeName__: string;
}

/**
 * デフォルトプロパティを持つオブジェクト。
 * Excel の Range.Value のように「オブジェクトそのものを値文脈で使う」場合に実装する。
 *
 * evaluator は読み書き両方とも `Value` プロパティを経由する:
 * - **読み取り**: `x = obj` で `resolveObjectMemberKey(obj, 'value')` を使って `Value` getter を呼ぶ
 * - **書き込み**: `obj = x` で `resolveObjectMemberKey(obj, 'value')` を使って `Value` setter に代入
 *
 * @example
 * class MockRange implements VbaDefaultProperty {
 *   readonly __vbaDefault__ = true as const;
 *   private _v: any = 0;
 *   get Value() { return this._v; }
 *   set Value(v: any) { this._v = v; }
 * }
 * // VBA: x = Range("A1")   → Value getter の値が x に入る
 * // VBA: Range("A1") = 100 → Value setter が呼ばれる
 */
export interface VbaDefaultProperty {
    readonly __vbaDefault__: true;
    Value: any;
}

/**
 * For Each でイテレートできるコレクション。
 * `For Each item In col` で要素を列挙するために Symbol.iterator を実装する。
 * または items プロパティ（any[]）で代替できる。
 *
 * @example
 * class MockCollection implements VbaIterable {
 *   private _items: any[] = [];
 *   [Symbol.iterator]() { return this._items[Symbol.iterator](); }
 * }
 * // VBA: For Each item In col → _items の各要素を列挙
 */
export interface VbaIterable {
    [Symbol.iterator](): Iterator<any>;
}

/**
 * `CreateObject(progId)` で生成される COM オブジェクトのモック。
 * `__progId__` に COM ProgID（`"Word.Application"` 等）を宣言することで、
 * `registerComObject` の自動別名登録が機能する。
 *
 * `registerComObject("Word.Application", factory)` は factory() を一度呼び出し、
 * 返り値の `__progId__` を読んで同じ factory を別名でも登録する。
 * これにより `CreateObject("Word.Application")` と `New Word.Application` の両方が動く。
 *
 * @example
 * class MockWordApplication implements VbaComObject {
 *   readonly __progId__ = 'Word.Application';
 * }
 * evaluator.registerComObject('Word.Application', () => new MockWordApplication());
 * // VBA: Set app = CreateObject("Word.Application")  → MockWordApplication
 * // VBA: Dim app As New Word.Application             → MockWordApplication（自動別名）
 */
export interface VbaComObject {
    readonly __progId__: string;
}

export interface AutoInstancePlaceholder {
    readonly __isAutoInstance__: true;
    readonly __className__: string;
}

export function createAutoInstancePlaceholder(className: string): AutoInstancePlaceholder {
    return { __isAutoInstance__: true, __className__: className };
}

export function isAutoInstancePlaceholder(v: any): v is AutoInstancePlaceholder {
    return v != null && typeof v === 'object' && v.__isAutoInstance__ === true;
}
