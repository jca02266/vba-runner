import { assert, evalVBASingle } from '../../test-libs/test-runner';
import { MemoryFileSystem } from '../../src/engine/filesystem';

const fs = new MemoryFileSystem();

const errors = evalVBASingle(`
    Function ProbeWidthErrors() As String
        On Error Resume Next
        Width #1, 80
        ProbeWidthErrors = CStr(Err.Number)
    End Function
    Function ProbeInvalidWidth() As String
        Open "width-invalid.dat" For Output As #1
        On Error Resume Next
        Width #1, 256
        ProbeInvalidWidth = CStr(Err.Number)
        Close #1
    End Function
`, { fs });
assert.strictEqual(errors.callProcedure('ProbeWidthErrors', []), '52', 'Width requires an open file');
assert.strictEqual(errors.callProcedure('ProbeInvalidWidth', []), '5', 'Width rejects values outside 0..255');

const wrapping = evalVBASingle(`
    Sub WriteWrapped()
        Open "width-wrap.dat" For Output As #1
        Width #1, 5
        Print #1, "123456789"
        Close #1
    End Sub
`, { fs });
wrapping.callProcedure('WriteWrapped', []);
assert.strictEqual(fs.readFileSync('/sandbox/width-wrap.dat', 'utf8'), '12345\r\n6789\r\n',
    'Width inserts line breaks at the configured character count');

const columns = evalVBASingle(`
    Sub WriteColumns()
        Open "width-columns.dat" For Output As #1
        Width #1, 20
        Print #1, "12";
        Print #1, Tab(4); "X"
        Close #1
    End Sub
`, { fs });
columns.callProcedure('WriteColumns', []);
assert.strictEqual(fs.readFileSync('/sandbox/width-columns.dat', 'utf8'), '12 X\r\n',
    'Tab uses the column carried across semicolon-terminated Print statements');

console.log('[PASS] Width statement validation and wrapping');
