/** Binary arithmetic with the 64-bit significand used by x87 extended values. */

interface ExtendedFloat {
    significand: bigint;
    exponent: number;
}

const EXTENDED_SIGNIFICAND_BITS = 64;
const DOUBLE_SIGNIFICAND_BITS = 53;

function bitLength(value: bigint): number {
    const absolute = value < 0n ? -value : value;
    return absolute === 0n ? 0 : absolute.toString(2).length;
}

function roundSignificand(
    significand: bigint,
    exponent: number,
    precision = EXTENDED_SIGNIFICAND_BITS,
): ExtendedFloat {
    if (significand === 0n) return { significand: 0n, exponent: 0 };

    const sign = significand < 0n ? -1n : 1n;
    let absolute = significand < 0n ? -significand : significand;
    const excessBits = bitLength(absolute) - precision;
    if (excessBits > 0) {
        let rounded = absolute >> BigInt(excessBits);
        const remainderMask = (1n << BigInt(excessBits)) - 1n;
        const remainder = absolute & remainderMask;
        const halfway = 1n << BigInt(excessBits - 1);
        if (remainder > halfway || (remainder === halfway && (rounded & 1n) !== 0n)) {
            rounded += 1n;
        }
        exponent += excessBits;
        if (bitLength(rounded) > precision) {
            rounded >>= 1n;
            exponent += 1;
        }
        absolute = rounded;
    }
    const missingBits = precision - bitLength(absolute);
    if (missingBits > 0) {
        absolute <<= BigInt(missingBits);
        exponent -= missingBits;
    }
    return { significand: sign * absolute, exponent };
}

function fromNumber(value: number): ExtendedFloat {
    if (!Number.isFinite(value)) throw new RangeError('extended precision requires a finite number');
    if (value === 0) return { significand: 0n, exponent: 0 };

    const buffer = new ArrayBuffer(8);
    const view = new DataView(buffer);
    view.setFloat64(0, value, false);
    const high = view.getUint32(0, false);
    const low = view.getUint32(4, false);
    const sign = (high >>> 31) === 0 ? 1n : -1n;
    const exponentBits = (high >>> 20) & 0x7ff;
    const fraction = (BigInt(high & 0xfffff) << 32n) | BigInt(low);

    if (exponentBits === 0) return roundSignificand(sign * fraction, -1074);
    return roundSignificand(
        sign * ((1n << 52n) | fraction),
        exponentBits - 1023 - 52,
    );
}

function add(left: ExtendedFloat, right: ExtendedFloat): ExtendedFloat {
    if (left.significand === 0n) return right;
    if (right.significand === 0n) return left;
    const exponentDistance = Math.abs(left.exponent - right.exponent);
    if (exponentDistance > EXTENDED_SIGNIFICAND_BITS + 2) {
        return left.exponent > right.exponent ? left : right;
    }
    const exponent = Math.min(left.exponent, right.exponent);
    return roundSignificand(
        (left.significand << BigInt(left.exponent - exponent))
            + (right.significand << BigInt(right.exponent - exponent)),
        exponent,
    );
}

function multiply(left: ExtendedFloat, right: ExtendedFloat): ExtendedFloat {
    return roundSignificand(
        left.significand * right.significand,
        left.exponent + right.exponent,
    );
}

function divide(left: ExtendedFloat, right: ExtendedFloat): ExtendedFloat {
    if (right.significand === 0n) throw new RangeError('extended precision division by zero');
    if (left.significand === 0n) return { significand: 0n, exponent: 0 };

    const negative = (left.significand < 0n) !== (right.significand < 0n);
    let numerator = left.significand < 0n ? -left.significand : left.significand;
    let denominator = right.significand < 0n ? -right.significand : right.significand;
    const guardBits = 3;
    const scale = EXTENDED_SIGNIFICAND_BITS + guardBits
        + bitLength(denominator) - bitLength(numerator);
    if (scale >= 0) numerator <<= BigInt(scale);
    else denominator <<= BigInt(-scale);

    let quotient = numerator / denominator;
    if (numerator % denominator !== 0n) quotient |= 1n;
    if (negative) quotient = -quotient;
    return roundSignificand(
        quotient,
        left.exponent - right.exponent - scale,
    );
}

function toNumber(value: ExtendedFloat): number {
    const rounded = roundSignificand(
        value.significand,
        value.exponent,
        DOUBLE_SIGNIFICAND_BITS,
    );
    return Number(rounded.significand) * Math.pow(2, rounded.exponent);
}

/**
 * Reproduce the signed NPV recurrence used by VBA financial functions. Each
 * arithmetic operation uses the runtime's extended significand, and the
 * completed accumulator is then stored back into a Double.
 */
export function vbaFinancialNpv(
    rate: number,
    values: number[],
    include: (value: number) => boolean,
): number {
    if (!Number.isFinite(rate) || values.some((value) => !Number.isFinite(value))) {
        return Number.NaN;
    }

    const extendedRate = fromNumber(rate);
    let divisor = fromNumber(1);
    let total = fromNumber(0);
    for (const value of values) {
        divisor = add(divisor, multiply(divisor, extendedRate));
        if (include(value)) total = add(total, divide(fromNumber(value), divisor));
    }
    return toNumber(total);
}
