/**
 * Common argument-contract predicates shared by user procedures, built-ins,
 * and dynamic members.  The caller remains responsible for mapping a
 * contract violation to the VBA error used by its call boundary.
 */
export interface VbaArgumentParameter {
    name?: string;
    optional?: boolean;
    isOptional?: boolean;
    defaultValue?: unknown;
    isParamArray?: boolean;
    isByVal?: boolean;
}

export type ArgumentContractViolation = 'missing' | 'excess' | null;

export function isOptionalArgument(parameter: VbaArgumentParameter): boolean {
    return parameter.optional === true || parameter.isOptional === true ||
        parameter.defaultValue != null;
}

export function isParamArrayArgument(parameter: VbaArgumentParameter): boolean {
    return parameter.isParamArray === true;
}

export function isRequiredArgument(parameter: VbaArgumentParameter): boolean {
    return !isOptionalArgument(parameter) && !isParamArrayArgument(parameter);
}

/** Property Let/Set's final parameter is the implicit assignment value. */
export function isPropertyValueParameter(
    procedure: { isProperty?: boolean; propertyType?: string; parameters: readonly VbaArgumentParameter[] },
    index: number,
): boolean {
    return procedure.isProperty === true &&
        (procedure.propertyType === 'let' || procedure.propertyType === 'set') &&
        index === procedure.parameters.length - 1;
}

/** Return whether a parameter is passed as a value at every call boundary. */
export function isEffectiveByValParameter(
    procedure: { isProperty?: boolean; propertyType?: string; parameters: readonly VbaArgumentParameter[] },
    index: number,
): boolean {
    const parameter = procedure.parameters[index] as (VbaArgumentParameter & { isByVal?: boolean }) | undefined;
    return parameter?.isByVal === true || isPropertyValueParameter(procedure, index);
}

export function requiredArgumentCount(parameters: readonly VbaArgumentParameter[]): number {
    return parameters.filter(isRequiredArgument).length;
}

export function hasParamArrayArgument(parameters: readonly VbaArgumentParameter[]): boolean {
    return parameters.some(isParamArrayArgument);
}

export function classifyArgumentCount(
    parameters: readonly VbaArgumentParameter[],
    providedCount: number,
): ArgumentContractViolation {
    if (providedCount < requiredArgumentCount(parameters)) return 'missing';
    if (!hasParamArrayArgument(parameters) && providedCount > parameters.length) return 'excess';
    return null;
}

/** Return whether a positional prefix leaves a required parameter omitted. */
export function hasRequiredArgumentAfterPrefix(
    parameters: readonly VbaArgumentParameter[],
    providedPrefixCount: number,
    endExclusive = parameters.length,
): boolean {
    for (let index = providedPrefixCount; index < endExclusive; index++) {
        if (isRequiredArgument(parameters[index])) {
            return true;
        }
    }
    return false;
}

/** Return whether a positional MissingArgument occupies a required slot. */
export function hasMissingRequiredArgument(
    parameters: readonly VbaArgumentParameter[],
    missingPositions: readonly boolean[],
): boolean {
    const paramArrayIndex = parameters.findIndex(isParamArrayArgument);
    const endExclusive = paramArrayIndex >= 0 ? paramArrayIndex : parameters.length;
    for (let index = 0; index < missingPositions.length && index < endExclusive; index++) {
        if (missingPositions[index] && isRequiredArgument(parameters[index])) return true;
    }
    return false;
}

export type NamedArgumentShapeViolation = 'missing' | 'invalid' | null;

/** Classify the declaration-shape part of a mixed positional/named call. */
export function classifyNamedArgumentShape(
    parameters: readonly VbaArgumentParameter[],
    positionalCount: number,
    namedNames: Iterable<string>,
    allowPropertyValueTail = false,
): NamedArgumentShapeViolation {
    const normalizedNamedNames = [...namedNames].map(name => name.toLowerCase());
    const propertyValueName = allowPropertyValueTail && parameters.length > 0
        ? parameters[parameters.length - 1].name?.toLowerCase()
        : undefined;
    const nonPropertyNamedNames = normalizedNamedNames.filter(name => name !== propertyValueName);
    const paramArrayIndex = parameters.findIndex(isParamArrayArgument);
    if (hasParamArrayArgument(parameters) && nonPropertyNamedNames.length > 0) return 'invalid';
    if (paramArrayIndex < 0 && positionalCount > parameters.length) return 'invalid';
    const names = new Set(parameters.map(parameter => parameter.name?.toLowerCase()));
    for (const name of normalizedNamedNames) {
        if (!names.has(name)) return 'invalid';
    }
    const positionalPrefixCount = paramArrayIndex >= 0
        ? Math.min(positionalCount, paramArrayIndex)
        : positionalCount;
    for (let index = 0; index < positionalPrefixCount; index++) {
        const name = parameters[index]?.name?.toLowerCase();
        if (name && normalizedNamedNames.includes(name)) return 'invalid';
    }
    for (let index = positionalPrefixCount; index < parameters.length; index++) {
        const parameter = parameters[index];
        const supplied = parameter.name !== undefined &&
            normalizedNamedNames.includes(parameter.name.toLowerCase());
        if (!supplied && isRequiredArgument(parameter)) return 'missing';
    }
    return null;
}

/** Validate the declaration-shape part of a mixed positional/named call. */
export function acceptsNamedArgumentShape(
    parameters: readonly VbaArgumentParameter[],
    positionalCount: number,
    namedNames: Iterable<string>,
    allowPropertyValueTail = false,
): boolean {
    return classifyNamedArgumentShape(parameters, positionalCount, namedNames, allowPropertyValueTail) === null;
}
