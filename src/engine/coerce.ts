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

// ---------------------------------------------------------------------------
// §6.1.2.2 Numeric value coercion
// ---------------------------------------------------------------------------

/**
 * Convert any VBA value to a JS number.
 *
 * - Empty → 0
 * - Null  → Type mismatch (Error 13)
 * - Boolean → -1 or 0
 * - numeric string → parsed value; non-numeric → Error 13
 * - ErrorValue → Error 13
 */
export function vbaToNumber(val: any): number {
    if (val === vbaEmpty) return 0;                          // null === null
    if (val === vbaNull) throwVbaError(VbaErrorCode.TYPE_MISMATCH);
    if (val instanceof VbaBoolean) return val.value;
    if (val instanceof VbaDate) return val.value;
    if (val instanceof VbaDecimal) return val.value;
    if (val instanceof VbaCurrency) return Number(val.internal) / 10000;
    if (typeof val === 'number') return val;
    if (typeof val === 'bigint') return Number(val);
    if (typeof val === 'string') {
        const trimmed = val.trim();
        // Handle hex literals: &H<hex> or &h<hex>
        if (trimmed.toLowerCase().startsWith('&h')) {
            const hexPart = trimmed.slice(2);
            if (!/^[0-9a-fA-F]+$/.test(hexPart)) throwVbaError(VbaErrorCode.TYPE_MISMATCH);
            return parseInt(hexPart, 16);
        }
        // Handle octal literals: &O<octal> or &o<octal>
        if (trimmed.toLowerCase().startsWith('&o')) {
            const octPart = trimmed.slice(2);
            if (!/^[0-7]+$/.test(octPart)) throwVbaError(VbaErrorCode.TYPE_MISMATCH);
            return parseInt(octPart, 8);
        }
        const n = parseFloat(trimmed);
        if (isNaN(n)) throwVbaError(VbaErrorCode.TYPE_MISMATCH);
        return n;
    }
    if (val instanceof VbaErrorValue) throwVbaError(VbaErrorCode.TYPE_MISMATCH);
    return Number(val);
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
    if (val instanceof VbaBoolean) return val;
    if (val === vbaEmpty) return vbaFalse;
    if (typeof val === 'number') return val !== 0 ? vbaTrue : vbaFalse;
    if (typeof val === 'boolean') return val ? vbaTrue : vbaFalse;
    if (typeof val === 'string') {
        const trimmed = val.trim();
        const lc = trimmed.toLowerCase();
        if (lc === 'true') return vbaTrue;
        if (lc === 'false') return vbaFalse;
        if (trimmed === '') throwVbaError(VbaErrorCode.TYPE_MISMATCH);
        const n = Number(trimmed);
        if (isNaN(n)) throwVbaError(VbaErrorCode.TYPE_MISMATCH);
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
 * - Others → String representation (delegates to toString())
 */
export function vbaToString(val: any): string {
    if (val === vbaNull) throwVbaError(VbaErrorCode.INVALID_USE_OF_NULL);
    if (val === vbaEmpty) return '';
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
