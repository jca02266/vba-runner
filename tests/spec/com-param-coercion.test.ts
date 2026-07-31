import { evalVBASingle, assert } from '../../test-libs/test-runner';

const result = evalVBASingle(`
    Function ProbeHostFlag() As String
        Dim host As Object
        Set host = CreateObject("Test.FlagHost")
        ProbeHostFlag = host.SetFlag("True") & "," & host.SetFlag("0")
    End Function
`, {
    setup: (ev) => {
        ev.registerComObject(() => ({
            __progId__: 'Test.FlagHost',
            __vbaParamSpecs__: {
                setflag: [{ name: 'Value', coerce: 'boolean' }],
            },
            setflag: (value: any) => value.value,
        } as any));
    },
});

assert.strictEqual(result.callProcedure('ProbeHostFlag', []), '-1,0',
    'generic COM metadata uses shared VBA Boolean coercion');
console.log('[PASS] generic COM parameter coercion');
