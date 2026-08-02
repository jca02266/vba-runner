import type { ProcedureDeclaration } from './parser';

export type PropertyKind = NonNullable<ProcedureDeclaration['propertyType']>;

/** Resolve a VBA class property by case-insensitive name and exact accessor kind. */
export function findClassProperty(
    procedures: readonly ProcedureDeclaration[],
    name: string,
    kind: PropertyKind,
): ProcedureDeclaration | undefined {
    const normalizedName = name.toLowerCase();
    return procedures.find(proc =>
        proc.isProperty &&
        proc.propertyType === kind &&
        proc.name.name.toLowerCase() === normalizedName
    );
}
