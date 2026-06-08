import { evalVBASingle, assert } from '../../test-libs/test-runner';

function runFunc(code: string, name: string, args: any[] = []): any {
    const ev = evalVBASingle(code, { onPrint: () => {} });
    return ev.callProcedure(name, args);
}

// 1. LSet Statement
{
    const code = `
        Function TestLSet()
            Dim s As String
            s = "ABCDE"
            LSet s = "123"
            TestLSet = s
        End Function
    `;
    assert.strictEqual(runFunc(code, 'TestLSet'), "123  ", 'LSet shorter');
}

{
    const code = `
        Function TestLSetLonger()
            Dim s As String
            s = "ABC"
            LSet s = "12345"
            TestLSetLonger = s
        End Function
    `;
    assert.strictEqual(runFunc(code, 'TestLSetLonger'), "123", 'LSet longer');
}

// 2. RSet Statement
{
    const code = `
        Function TestRSet()
            Dim s As String
            s = "ABCDE"
            RSet s = "123"
            TestRSet = s
        End Function
    `;
    assert.strictEqual(runFunc(code, 'TestRSet'), "  123", 'RSet shorter');
}

console.log('\n✅ LSet/RSet Statement: 全テスト通過');
