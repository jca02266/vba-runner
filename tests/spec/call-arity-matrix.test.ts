import { assertCompileErrorPreproc, assertCompileErrorExec, evalVBASingle, assert } from '../../test-libs/test-runner';

type ProcedureKind = 'Function' | 'Sub';

const definitions = (kind: ProcedureKind, params: string): string => kind === 'Function'
    ? String.raw`Function Target(${params}) As Long
    Target = 0
End Function`
    : String.raw`Sub Target(${params})
End Sub`;

const calls = (kind: ProcedureKind, argumentCount: number): string => {
    const args = Array.from({ length: argumentCount }, (_, index) => String(index + 1));
    if (kind === 'Function') return `result = Target(${args.join(', ')})`;
    return args.length === 0 ? 'Call Target' : `Call Target(${args.join(', ')})`;
};

const source = (kind: ProcedureKind, definitionParameters: string, argumentCount: number): string => String.raw`
${definitions(kind, definitionParameters)}
Sub Caller()
    Dim result As Long
    ${calls(kind, argumentCount)}
End Sub
`;

const implicitFunctionSource = (definitionParameters: string): string => String.raw`
${definitions('Function', definitionParameters)}
Sub Caller()
    Dim result As Long
    result = Target
End Sub
`;

const parameterCases = [
    { label: 'no-parameters', declaration: '', minCount: 0, maxCount: 0 },
    { label: 'one-required', declaration: 'value As Long', minCount: 1, maxCount: 1 },
    { label: 'two-required', declaration: 'left As Long, right As Long', minCount: 2, maxCount: 2 },
    { label: 'one-optional', declaration: 'Optional value As Long = 0', minCount: 0, maxCount: 1 },
    { label: 'two-optional', declaration: 'Optional left As Long = 0, Optional right As Long = 0', minCount: 0, maxCount: 2 },
];

for (const kind of ['Function', 'Sub'] as const) {
    for (const definition of parameterCases) {
        for (const argumentCount of [0, 1, 2]) {
            const label = `${kind}/${definition.label}/args-${argumentCount}`;
            const code = source(kind, definition.declaration, argumentCount);
            if (argumentCount < definition.minCount || argumentCount > definition.maxCount) {
                assertCompileErrorPreproc(code, 'Caller', kind === 'Function' ? 7 : 6,
                    /argument not optional|wrong number of arguments/i, label);
            } else {
                const runner = evalVBASingle(code);
                assert.doesNotThrow(() => runner.callProcedure('Caller', []), label);
            }
        }
    }
}

for (const definition of parameterCases) {
    const label = `Function/implicit-value/${definition.label}`;
    const code = implicitFunctionSource(definition.declaration);
    if (definition.minCount > 0) {
        assertCompileErrorPreproc(code, 'Caller', 7, /argument not optional/i, label);
    } else {
        const runner = evalVBASingle(code);
        assert.doesNotThrow(() => runner.callProcedure('Caller', []), label);
    }
}

type MemberKind = 'Function' | 'Property Get' | 'Sub';

const classSource = (memberKind: MemberKind, definitionParameters: string, argumentCount: number): string => {
    const args = Array.from({ length: argumentCount }, (_, index) => String(index + 1));
    const declaration = memberKind === 'Property Get'
        ? `Public Property Get Target(${definitionParameters}) As Long`
        : `Public ${memberKind} Target(${definitionParameters})${memberKind === 'Function' ? ' As Long' : ''}`;
    const body = memberKind === 'Function' || memberKind === 'Property Get' ? '    Target = 0\n' : '';
    const call = memberKind === 'Sub'
        ? (args.length === 0 ? 'Call instance.Target' : `Call instance.Target(${args.join(', ')})`)
        : `result = instance.Target(${args.join(', ')})`;
    return String.raw`Class ArityTarget
${declaration}
${body}End ${memberKind === 'Property Get' ? 'Property' : memberKind}
End Class
Sub Caller()
    Dim instance As New ArityTarget
    Dim result As Long
    ${call}
End Sub
`;
};

for (const memberKind of ['Function', 'Property Get', 'Sub'] as const) {
    for (const definition of parameterCases) {
        for (const argumentCount of [0, 1, 2]) {
            const label = `Class/${memberKind}/${definition.label}/args-${argumentCount}`;
            const code = classSource(memberKind, definition.declaration, argumentCount);
            if (memberKind === 'Property Get' && definition.minCount === 0 && argumentCount > 0) {
                // A zero-argument Property Get followed by arguments is VBA's
                // returned-value indexing path, not a member arity call.
                continue;
            }
            if (argumentCount < definition.minCount || argumentCount > definition.maxCount) {
                assertCompileErrorExec(code, 'Caller', memberKind === 'Sub' ? 8 : 9,
                    /argument not optional|wrong number of arguments/i, label);
            } else {
                const runner = evalVBASingle(code);
                assert.doesNotThrow(() => runner.callProcedure('Caller', []), label);
            }
        }
    }
}

console.log('✅ Function/Sub definition and call arity matrix passed');
