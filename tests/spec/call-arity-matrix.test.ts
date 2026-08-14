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
                if (kind === 'Function' && argumentCount > definition.maxCount) {
                    assertCompileErrorExec(code, 'Caller', undefined, /type mismatch/i, label);
                } else {
                    assertCompileErrorPreproc(code, 'Caller', kind === 'Function' ? 7 : 6,
                        /argument not optional|wrong number of arguments/i, label);
                }
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
                assertCompileErrorExec(code, 'Caller', undefined,
                    /argument not optional|wrong number of arguments/i, label);
            } else {
                const runner = evalVBASingle(code);
                assert.doesNotThrow(() => runner.callProcedure('Caller', []), label);
            }
        }
    }
}

// Statement-call routes without the Call keyword are parsed separately from
// CallExpression.  Keep the same zero/one/two-argument boundary matrix here.
for (const argumentCount of [0, 1, 2]) {
    const args = Array.from({ length: argumentCount }, (_, index) => String(index + 1));
    const code = String.raw`Sub Target(value As Long)
End Sub
Sub Caller()
    Target${args.length ? ` ${args.join(', ')}` : ''}
End Sub
`;
    if (argumentCount !== 1) {
        assertCompileErrorPreproc(code, 'Caller', 4,
            /argument not optional|wrong number of arguments/i, `Sub/no-call-keyword/args-${argumentCount}`);
    } else {
        const runner = evalVBASingle(code);
        assert.doesNotThrow(() => runner.callProcedure('Caller', []), 'Sub/no-call-keyword/args-1');
    }
}

// Module-qualified calls use getProcedureFromModule rather than the ordinary
// identifier lookup.
for (const argumentCount of [0, 1, 2]) {
    const args = Array.from({ length: argumentCount }, (_, index) => String(index + 1));
    const code = String.raw`Attribute VB_Name = "CallerModule"
Function Target(value As Long) As Long
    Target = value
End Function
Sub Caller()
    Dim result As Long
    result = CallerModule.Target(${args.join(', ')})
End Sub
`;
    if (argumentCount !== 1) {
        assertCompileErrorExec(code, 'Caller', 7,
            /argument not optional|wrong number of arguments/i, `Module-qualified/args-${argumentCount}`);
    } else {
        const runner = evalVBASingle(code);
        assert.doesNotThrow(() => runner.callProcedure('Caller', []), 'Module-qualified/args-1');
    }
}

// Optional holes and named arguments exercise MissingArgument and the
// name-to-parameter alignment path rather than only positional counts.
assertCompileErrorExec(String.raw`Function Target(left As Long, right As Long) As Long
    Target = left + right
End Function
Sub Caller()
    Dim result As Long
    result = Target(, 2)
End Sub
`, 'Caller', 6, /argument not optional/i, 'Optional-hole/required-gap');
{
    const runner = evalVBASingle(String.raw`Function Target(left As Long, Optional right As Long = 5) As Long
    Target = left + right
End Function
Sub Caller()
    Dim result As Long
    result = Target(right:=2, left:=1)
End Sub
`);
    assert.doesNotThrow(() => runner.callProcedure('Caller', []), 'Named-arguments/reordered');
}

// ParamArray is an unbounded tail, but its required prefix still participates
// in the same minimum-arity contract.
assertCompileErrorPreproc(String.raw`Function Target(required As Long, ParamArray values()) As Long
    Target = required
End Function
Sub Caller()
    Dim result As Long
    result = Target()
End Sub
`, 'Caller', 6, /argument not optional/i, 'ParamArray/required-prefix-missing');
{
    const runner = evalVBASingle(String.raw`Function Target(required As Long, ParamArray values()) As Long
    Target = required
End Function
Sub Caller()
    Dim result As Long
    result = Target(1, 2, 3, 4)
End Sub
`);
assert.doesNotThrow(() => runner.callProcedure('Caller', []), 'ParamArray/unbounded-tail');
}

// Property Let/Set assignments append the RHS as an implicit final
// parameter.  An indexed property must still reject an omitted required index.
for (const propertyKind of ['Let', 'Set'] as const) {
    const valueType = propertyKind === 'Set' ? 'Object' : 'Long';
    const assignment = propertyKind === 'Set' ? 'Set instance.Value = child' : 'instance.Value = 2';
    const source = String.raw`Class ArityProperty
Public Property Get Status() As Long
    Status = 0
End Property
Public Property ${propertyKind} Value(index As Long, value As ${valueType})
    Status = 1
End Property
End Class
Sub Caller()
    Dim instance As New ArityProperty
    ${propertyKind === 'Set' ? 'Dim child As New ArityProperty' : ''}
    ${assignment}
End Sub
`;
    assertCompileErrorExec(source, 'Caller', undefined, /argument not optional/i,
        `Property ${propertyKind}/required-index-omitted`);
}

// A bare Property Get or module-qualified Function is still an implicit call.
// Required parameters must be rejected during precheck, before object lookup
// or runtime member evaluation can produce a secondary error.
assertCompileErrorPreproc(String.raw`Class ArityGetter
Public Property Get Item(index As Long) As Long
    Item = index
End Property
End Class
Sub Caller()
    Dim instance As New ArityGetter
    Dim result As Long
    result = instance.Item
End Sub
`, 'Caller', undefined, /argument not optional/i, 'Property-Get/bare-required-index');

console.log('✅ Function/Sub definition and call arity matrix passed');
