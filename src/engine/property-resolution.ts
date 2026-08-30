import type { ClassDeclaration, ProcedureDeclaration, Statement } from './parser';

export type PropertyKind = NonNullable<ProcedureDeclaration['propertyType']>;

export type ClassProcedureKind = 'get' | 'let' | 'set' | 'method' | 'any';

/**
 * Find a procedure declaration in a parsed module, including class members.
 *
 * AST consumers used to each walk only the top-level declarations (or only
 * `ClassDeclaration.procedures`).  Keeping this traversal beside the class
 * member predicates makes the LSP and debug adapters apply the same
 * case-insensitive lookup rules as the evaluator.
 */
export function findProcedureDeclaration(
    statements: readonly Statement[],
    name: string,
    kind: ClassProcedureKind = 'any',
): ProcedureDeclaration | undefined {
    const normalizedName = name.toLowerCase();
    for (const statement of statements) {
        if (statement.type === 'ProcedureDeclaration') {
            const procedure = statement as ProcedureDeclaration;
            if (procedure.name.name.toLowerCase() !== normalizedName) continue;
            if (kind === 'method' && procedure.isProperty) continue;
            if (kind !== 'any' && kind !== 'method' &&
                (!procedure.isProperty || procedure.propertyType !== kind)) continue;
            return procedure;
        }
        if (statement.type === 'ClassDeclaration') {
            const member = findClassProcedure(
                (statement as ClassDeclaration).procedures,
                name,
                kind,
            );
            if (member) return member;
        }
    }
    return undefined;
}

/**
 * Resolve a class member by VBA's case-insensitive member name.
 *
 * Keeping this lookup in one place is important: the evaluator has several
 * entry points (expression calls, assignments, With statements, and event
 * wiring), but they must all apply the same accessor/name rules.
 */
export function findClassProcedure(
    procedures: readonly ProcedureDeclaration[],
    name: string,
    kind: ClassProcedureKind = 'any',
): ProcedureDeclaration | undefined {
    const normalizedName = name.toLowerCase();
    return procedures.find(proc => {
        if (proc.name.name.toLowerCase() !== normalizedName) return false;
        if (kind === 'method') return !proc.isProperty;
        if (kind === 'any') return true;
        return proc.isProperty && proc.propertyType === kind;
    });
}

/** Find the first class member matching a kind (used for default debug entry). */
export function findFirstClassProcedure(
    procedures: readonly ProcedureDeclaration[],
    kind: ClassProcedureKind = 'any',
): ProcedureDeclaration | undefined {
    return procedures.find(proc => {
        if (kind === 'method') return !proc.isProperty;
        if (kind === 'any') return true;
        return proc.isProperty && proc.propertyType === kind;
    });
}

/** Resolve either Property Let or Property Set, preserving declaration order. */
export function findClassSetter(
    procedures: readonly ProcedureDeclaration[],
    name: string,
): ProcedureDeclaration | undefined {
    const normalizedName = name.toLowerCase();
    return procedures.find(proc =>
        proc.name.name.toLowerCase() === normalizedName &&
        proc.isProperty && (proc.propertyType === 'let' || proc.propertyType === 'set')
    );
}

/** Resolve a VBA class property by case-insensitive name and exact accessor kind. */
export function findClassProperty(
    procedures: readonly ProcedureDeclaration[],
    name: string,
    kind: PropertyKind,
): ProcedureDeclaration | undefined {
    return findClassProcedure(procedures, name, kind);
}
