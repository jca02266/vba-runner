import { evalVBASingle, assert } from '../../test-libs/test-runner';

function runFunc(code: string, name: string, args: any[] = []): any {
    const ev = evalVBASingle(code, { onPrint: () => {} });
    return ev.callProcedure(name, args);
}

// 1. Exit Property Statement
{
    const code = `
        Property Get TestProp()
            TestProp = 1
            Exit Property
            TestProp = 2
        End Property
    `;
    // Note: callProcedure might need to handle property names if they are stored differently,
    // but in my implementation it just finds the procedure by name.
    assert.strictEqual(runFunc(code, 'TestProp'), 1, 'Exit Property should terminate execution immediately');
}

console.log('\n✅ Exit Property Statement: 全テスト通過');
