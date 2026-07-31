import { evalVBASingle, assert } from '../../test-libs/test-runner';

console.log("Running CDec tests...");

const tests = [
    { code: "Debug.Print CDec(123.45)", expected: "123.45" },
    { code: "Debug.Print TypeName(CDec(1))", expected: "Decimal" },
    { code: "Debug.Print VarType(CDec(1))", expected: "14" },
];

for (const t of tests) {
    let output = '';
    evalVBASingle(t.code, { onPrint: (o) => { output = o; } });
    assert.strictEqual(output.trim(), t.expected, t.code);
    console.log(`[PASS] ${t.code}`);
}

const boundary = evalVBASingle(`
    Function ProbeCDecForms() As String
        ProbeCDecForms = CStr(CDec("+1")) & "|" & CStr(CDec("&H10")) & "|" & _
            CStr(CDec("&O10")) & "|" & CStr(CDec("1,234")) & "|" & CStr(CDec("1.2E3"))
    End Function
`);
assert.strictEqual(boundary.callProcedure('ProbeCDecForms', []), '1|16|8|1234|1200',
    'CDec accepts the Excel-verified signed, radix, grouped, and exponent forms');

console.log('✅ CDec: 全テスト通過');
