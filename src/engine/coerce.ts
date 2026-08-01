/**
 * VBA value coercion functions.
 *
 * These implement the conversion rules from MS-VBAL:
 *   §6.1.2.2  — Numeric value coercion
 *   §6.1.2.3.1.1 — Boolean value coercion (CBool rules)
 *   §6.1.2.4  — String value coercion (CStr rules)
 *
 * All functions throw plain VbaError objects { type, number, message }.
 * Callers that need line-number-enriched errors should catch and re-throw
 * using Evaluator.throwVbaError.
 */
import {
    VbaBoolean, vbaTrue, vbaFalse,
    vbaEmpty, vbaNull, vbaNothing,
    VbaDate, VbaDecimal, VbaCurrency, VbaErrorValue,
} from './vba-types';
import { VbaErrorCode, throwVbaError } from './vba-errors';

/** Expand explicit VBA default Value chains and reject circular chains. */
export function unwrapDefaultValue(value: any): any {
    const seen = new Set<any>();
    let current = value;
    while (current && typeof current === 'object' && current.__vbaDefault__ === true) {
        if (seen.has(current)) throwVbaError(VbaErrorCode.TYPE_MISMATCH);
        seen.add(current);
        current = current.Value;
    }
    return current;
}

// ---------------------------------------------------------------------------
// §6.1.2.2 Numeric value coercion
// ---------------------------------------------------------------------------

/** Normalize VBA's D/d exponent marker before JavaScript numeric parsing. */
export function normalizeVbaNumericString(value: string): string {
    return value.replace(
        /^([+-]?(?:\d+(?:\.\d*)?|\.\d+))[dD]([+-]?\d+)$/,
        '$1E$2',
    );
}

/**
 * Convert any VBA value to a JS number.
 *
 * - Empty → 0
 * - Null  → Type mismatch (Error 13)
 * - Nothing → Object variable not set (Error 91)
 * - Boolean → -1 or 0
 * - numeric string → parsed value; non-numeric → Error 13
 * - ErrorValue → Error 13
 */
export function vbaToNumber(val: any): number {
    val = unwrapDefaultValue(val);
    if (val === vbaEmpty) return 0;                          // null === null
    if (val === vbaNull) throwVbaError(VbaErrorCode.TYPE_MISMATCH);
    if (val === vbaNothing) throwVbaError(VbaErrorCode.OBJECT_VARIABLE_NOT_SET);
    if (val instanceof VbaBoolean) return val.value;
    if (val instanceof VbaDate) return val.value;
    if (val instanceof VbaDecimal) return val.value;
    if (val instanceof VbaCurrency) return Number(val.internal) / 10000;
    if (typeof val === 'number') {
        if (!Number.isFinite(val)) throwVbaError(VbaErrorCode.OVERFLOW);
        return val;
    }
    if (typeof val === 'bigint') {
        const n = Number(val);
        if (!Number.isFinite(n)) throwVbaError(VbaErrorCode.OVERFLOW);
        return n;
    }
    if (typeof val === 'string') {
        const trimmed = normalizeVbaNumericString(val.trim());
        // Handle hex literals: &H<hex> or &h<hex>
        if (trimmed.toLowerCase().startsWith('&h')) {
            const hexPart = trimmed.slice(2);
            if (!/^[0-9a-fA-F]+$/.test(hexPart)) throwVbaError(VbaErrorCode.TYPE_MISMATCH);
            const n = parseInt(hexPart, 16);
            if (!Number.isFinite(n)) throwVbaError(VbaErrorCode.OVERFLOW);
            return n;
        }
        // Handle octal literals: &O<octal> or &o<octal>
        if (trimmed.toLowerCase().startsWith('&o')) {
            const octPart = trimmed.slice(2);
            if (!/^[0-7]+$/.test(octPart)) throwVbaError(VbaErrorCode.TYPE_MISMATCH);
            const n = parseInt(octPart, 8);
            if (!Number.isFinite(n)) throwVbaError(VbaErrorCode.OVERFLOW);
            return n;
        }
        // カンマ桁区切り（"1,234"）は区切りを除去して解釈（実 VBA 差分で裁定）
        const grouped = /^[+-]?\d{1,3}(,\d{3})+(\.\d+)?$/.test(trimmed)
            ? trimmed.replace(/,/g, '') : trimmed;
        const n = parseFloat(grouped);
        if (isNaN(n) || !/^[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$/.test(grouped)) {
            throwVbaError(VbaErrorCode.TYPE_MISMATCH);
        }
        if (!Number.isFinite(n)) throwVbaError(VbaErrorCode.OVERFLOW);
        return n;
    }
    if (val instanceof VbaErrorValue) throwVbaError(VbaErrorCode.TYPE_MISMATCH);
    if (val === undefined) return 0;                         // Empty 相当（IsEmpty と同じ扱い）
    if (Array.isArray(val)) throwVbaError(VbaErrorCode.TYPE_MISMATCH);
    const n = Number(val);
    if (!Number.isFinite(n)) throwVbaError(VbaErrorCode.OVERFLOW);
    return n;
}

// ---------------------------------------------------------------------------
// Banker's rounding (VBA Round() and implicit coercion to integer types)
// ---------------------------------------------------------------------------

/**
 * VBA "round half to even" (banker's rounding).
 */
export function vbaRound(val: number, decimals: number = 0): number {
    const factor = Math.pow(10, decimals);
    const scaled = val * factor;
    const i = Math.floor(scaled);
    const f = scaled - i;
    const epsilon = 1e-10;
    if (Math.abs(f - 0.5) < epsilon) {
        return (i % 2 === 0 ? i : i + 1) / factor;
    }
    return Math.round(scaled) / factor;
}

// ---------------------------------------------------------------------------
// §6.1.2.3.1.1 Boolean value coercion (CBool rules)
// ---------------------------------------------------------------------------

/**
 * Convert any VBA value to VbaBoolean, following CBool coercion rules.
 *
 * - VbaBoolean → passthrough
 * - Empty → False
 * - number → False if 0, True otherwise
 * - "True"/"False" (case-insensitive) → corresponding Boolean
 * - numeric string → False if 0, True otherwise
 * - empty string / non-numeric string → Type mismatch (Error 13)
 * - Null → caller must check (Error 94 per §5.6.9); this function throws Error 13
 */
export function vbaToBoolean(val: any): VbaBoolean {
    val = unwrapDefaultValue(val);
    // Null は Boolean 引数の型変換では Type mismatch ではなく、
    // VBA の「Invalid use of Null」(Error 94) になる。
    if (val === vbaNull) throwVbaError(VbaErrorCode.INVALID_USE_OF_NULL);
    if (val instanceof VbaBoolean) return val;
    if (val === vbaEmpty) return vbaFalse;
    if (val === vbaNothing) throwVbaError(VbaErrorCode.OBJECT_VARIABLE_NOT_SET);
    if (typeof val === 'number') return val !== 0 ? vbaTrue : vbaFalse;
    if (typeof val === 'boolean') return val ? vbaTrue : vbaFalse;
    if (val instanceof VbaDecimal) return val.value !== 0 ? vbaTrue : vbaFalse;
    if (val instanceof VbaCurrency) return val.internal !== 0n ? vbaTrue : vbaFalse;
    if (val instanceof VbaDate) return val.value !== 0 ? vbaTrue : vbaFalse;
    if (typeof val === 'string') {
        const trimmed = val.trim();
        const lc = trimmed.toLowerCase();
        if (lc === 'true') return vbaTrue;
        if (lc === 'false') return vbaFalse;
        if (trimmed === '') throwVbaError(VbaErrorCode.TYPE_MISMATCH);
        // &H/&O/指数/カンマ区切りも数値文字列として解釈（実 VBA 差分で裁定）
        const n = vbaToNumber(trimmed);
        return n !== 0 ? vbaTrue : vbaFalse;
    }
    throwVbaError(VbaErrorCode.TYPE_MISMATCH);
}

// ---------------------------------------------------------------------------
// §6.1.2.4 String value coercion (CStr rules)
// ---------------------------------------------------------------------------

/**
 * Convert any VBA value to a string, following CStr coercion rules.
 *
 * - Null  → Invalid use of Null (Error 94)
 * - Empty → ""
 * - Nothing → Object variable not set (Error 91)
 * - Others → String representation (delegates to toString())
 */
export function vbaToString(val: any): string {
    val = unwrapDefaultValue(val);
    if (val === vbaNull) throwVbaError(VbaErrorCode.INVALID_USE_OF_NULL);
    if (val === vbaEmpty) return '';
    if (val === vbaNothing) throwVbaError(VbaErrorCode.OBJECT_VARIABLE_NOT_SET);
    if (Array.isArray(val)) throwVbaError(VbaErrorCode.TYPE_MISMATCH);
    // Excel-like objects explicitly opt in to a default Value property.  In
    // VBA value contexts such as CStr(Range("A1")), the property's value is
    // coerced rather than the JavaScript object identity.
    // Date, Boolean, Currency, Decimal, and Error values have VBA string
    // representations. Other objects require a default property; the runner
    // cannot infer one, so reject them instead of leaking JS object text.
    if (typeof val === 'object' &&
        !(val instanceof VbaBoolean) && !(val instanceof VbaDate) &&
        !(val instanceof VbaDecimal) && !(val instanceof VbaCurrency) &&
        !(val instanceof VbaErrorValue)) {
        throwVbaError(VbaErrorCode.TYPE_MISMATCH);
    }
    if (typeof val === 'number' && Number.isFinite(val) && Math.abs(val) >= 1e15) {
        // CStr(Double) switches to VBA's general scientific form at this
        // magnitude. Keep 15 significant digits, matching Excel's display
        // for LongLong/String mixed arithmetic.
        return val.toExponential(14)
            .replace(/\.0+e/, 'e')
            .replace(/(\d)0+e/, '$1e')
            .replace('e', 'E');
    }
    return String(val);
}

// ---------------------------------------------------------------------------
// Display string — for Print # / Debug.Print (not the same as CStr)
// ---------------------------------------------------------------------------

/**
 * Convert any VBA value to a display string for Print / Debug.Print.
 * Unlike CStr, Null → "Null" and Nothing → "Nothing".
 */
export function vbaToDisplayString(val: any): string {
    if (val === vbaEmpty || val === undefined) return '';
    if (val === vbaNull) return 'Null';
    if (val === vbaNothing) return 'Nothing';
    if (val instanceof VbaBoolean) return val.toString();   // "True" / "False"
    if (val instanceof VbaDate) return val.toString();
    if (val instanceof VbaErrorValue) return val.toString();
    if (val instanceof VbaDecimal) return val.toString();
    if (val instanceof VbaCurrency) return val.toString();
    if (typeof val === 'bigint') return val.toString();
    return String(val);
}
