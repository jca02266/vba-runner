import { evalVBASingle, assert } from '../../test-libs/test-runner';

const ev = evalVBASingle(`
    Function AddressText() As String
        AddressText = AddressOf NS.Module.Proc
    End Function
`);

assert.strictEqual(ev.callProcedure('AddressText', []), 'NS.Module.Proc');

console.log('✅ AddressOf accepts multi-qualified procedure names');
