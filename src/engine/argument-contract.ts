/**
 * Common argument-contract predicates shared by user procedures, built-ins,
 * and dynamic members.  The caller remains responsible for mapping a
 * contract violation to the VBA error used by its call boundary.
 */
export interface VbaArgumentParameter {
    optional?: boolean;
    isOptional?: boolean;
    defaultValue?: unknown;
    isParamArray?: boolean;
}

export type ArgumentContractViolation = 'missing' | 'excess' | null;

export function isOptionalArgument(parameter: VbaArgumentParameter): boolean {
    return parameter.optional === true || parameter.isOptional === true ||
        parameter.defaultValue != null;
}

export function isParamArrayArgument(parameter: VbaArgumentParameter): boolean {
    return parameter.isParamArray === true;
}

export function requiredArgumentCount(parameters: readonly VbaArgumentParameter[]): number {
    return parameters.filter(parameter =>
        !isOptionalArgument(parameter) && !isParamArrayArgument(parameter)
    ).length;
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
        if (!isOptionalArgument(parameters[index]) && !isParamArrayArgument(parameters[index])) {
            return true;
        }
    }
    return false;
}
