import { evalVBAModules } from '../../test-libs/test-runner';

const modules = [
  {
    name: 'Box',
    code: `Class Box
Public Seen As Long
Public Property Let Values(ByRef incoming() As Long)
  Seen = incoming(0)
End Property
End Class`,
  },
  {
    name: 'Module1',
    code: `Function Probe() As Long
Dim b As New Box, values(0 To 0) As Long
values(0) = 7
b.Values = values
Probe = b.Seen
End Function`,
  },
];

const runner = evalVBAModules(modules);
const result = runner.callProcedure('Probe', []);
if (result !== 7) throw new Error(`expected Property Let array value 7, got ${result}`);
console.log('[PASS] Property Let accepts typed array argument');
