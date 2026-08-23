import { assertCompileErrorPreproc, assertCompileErrorExec, evalVBAModules, evalVBASingle, assert } from '../../test-libs/test-runner';

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

// A declaration may carry a default value without the Optional keyword in
// exported VBA sources.  Every implicit-call route must treat that parameter
// as optional, just like the explicit Optional spelling.
{
    const moduleDefault = evalVBASingle(String.raw`Function Target(value As Long = 1) As Long
    Target = value
End Function
Sub Caller()
    Dim result As Long
    result = Target
End Sub
`);
    assert.doesNotThrow(() => moduleDefault.callProcedure('Caller', []),
        'Implicit-call/default-value/module');

    const classDefault = evalVBASingle(String.raw`Class DefaultValueTarget
Public Function Target(value As Long = 1) As Long
    Target = value
End Function
End Class
Sub Caller()
    Dim instance As New DefaultValueTarget
    Dim result As Long
    result = instance.Target
End Sub
`);
    assert.doesNotThrow(() => classDefault.callProcedure('Caller', []),
        'Implicit-call/default-value/class');

    const withDefault = evalVBASingle(String.raw`Class WithDefaultValueTarget
Public Function Target(value As Long = 1) As Long
    Target = value
End Function
End Class
Sub Caller()
    Dim instance As New WithDefaultValueTarget
    Dim result As Long
    With instance
        result = .Target
    End With
End Sub
`);
    assert.doesNotThrow(() => withDefault.callProcedure('Caller', []),
        'Implicit-call/default-value/with');
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
                if (memberKind === 'Function' && argumentCount > definition.maxCount) {
                    assertCompileErrorExec(code, 'Caller', undefined, /type mismatch/i, label);
                } else {
                    assertCompileErrorPreproc(code, 'Caller', undefined,
                        /argument not optional|wrong number of arguments/i, label);
                }
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
    if (argumentCount === 0) {
        assertCompileErrorPreproc(code, 'Caller', 7, /argument not optional/i,
            `Module-qualified/args-${argumentCount}`);
    } else if (argumentCount > 1) {
        assertCompileErrorExec(code, 'Caller', undefined, /type mismatch/i,
            `Module-qualified/args-${argumentCount}`);
    } else {
        const runner = evalVBASingle(code);
        assert.doesNotThrow(() => runner.callProcedure('Caller', []), 'Module-qualified/args-1');
    }
}

// A module-qualified Function used as a bare value must use the same
// required-argument contract as an unqualified implicit call.
assertCompileErrorPreproc(String.raw`Attribute VB_Name = "CallerModule"
Function Target(value As Long) As Long
    Target = value
End Function
Sub Caller()
    Dim result As Long
    result = CallerModule.Target
End Sub
`, 'Caller', 7, /argument not optional/i, 'Module-qualified/bare-required-argument');

// A module-qualified expression call must retain the parenthesized ByVal
// marker instead of routing through a value-only public call boundary.
{
    const runner = evalVBASingle(String.raw`Attribute VB_Name = "CallerModule"
Function Mutate(ByRef value As Long) As Long
    value = 99
    Mutate = value
End Function
Function Probe() As Long
    Dim value As Long
    value = 1
    CallerModule.Mutate (value)
    Probe = value
End Function
`);
    assert.strictEqual(runner.callProcedure('Probe', []), 1,
        'Module-qualified/parenthesized-ByVal');
}

// A zero-argument module Property Get that returns a declared array must use
// the shared array reader.  Direct JavaScript indexing would turn an invalid
// upper/lower subscript into Empty and would accept the wrong rank.
{
    const runner = evalVBAModules([{
        name: 'ArrayPropertyProvider',
        code: String.raw`Public Property Get Values() As Variant
    Dim data(0 To 1) As Long
    data(0) = 7
    data(1) = 8
    Values = data
End Property`,
    }]);

    assert.strictEqual(runner.evalExpression('ArrayPropertyProvider.Values(0)'), 7,
        'Module Property Get array lower bound');
    assert.strictEqual(runner.evalExpression('ArrayPropertyProvider.Values(1)'), 8,
        'Module Property Get array upper bound');
    assert.throwsMatch(() => runner.evalExpression('ArrayPropertyProvider.Values(-1)'),
        /error '9'/i, 'Module Property Get array negative subscript');
    assert.throwsMatch(() => runner.evalExpression('ArrayPropertyProvider.Values(2)'),
        /error '9'/i, 'Module Property Get array upper subscript');
    assert.throwsMatch(() => runner.evalExpression('ArrayPropertyProvider.Values(0, 0)'),
        /error '9'/i, 'Module Property Get array rank mismatch');
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

// Supplying a later named Optional argument must not bypass a required prefix.
// The same contract applies to module and class procedure dispatch.
assertCompileErrorExec(String.raw`Function Target(required As Long, Optional tail As Long = 2) As Long
    Target = required + tail
End Function
Sub Caller()
    Dim result As Long
    result = Target(tail:=4)
End Sub
`, 'Caller', undefined, /argument not optional/i, 'Named-arguments/required-prefix-module');
assertCompileErrorExec(String.raw`Class NamedArityTarget
Public Function Target(required As Long, Optional tail As Long = 2) As Long
    Target = required + tail
End Function
End Class
Sub Caller()
    Dim instance As New NamedArityTarget
    Dim result As Long
    result = instance.Target(tail:=4)
End Sub
`, 'Caller', undefined, /argument not optional/i, 'Named-arguments/required-prefix-class');

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

// A parameter may shadow a procedure name.  The identifier on the RHS is
// then a variable, not an implicit zero-argument function call.
{
    const runner = evalVBASingle(String.raw`Function D(value As Long) As Long
    D = value
End Function
Function Wrapper(d As Long) As Long
    Wrapper = d
End Function
Sub Caller()
    Dim result As Long
    result = Wrapper(1)
End Sub
`);
    assert.doesNotThrow(() => runner.callProcedure('Caller', []),
        'Implicit function precheck respects parameter shadowing');
}

console.log('✅ Function/Sub definition and call arity matrix passed');
