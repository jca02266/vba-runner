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

export class VbaDecimal {
    constructor(public value: number) {}
    toString() { return String(this.value); }
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
