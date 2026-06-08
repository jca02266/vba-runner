import { evalVBASingle, assert, vbaNull } from '../../test-libs/test-runner';

function evalVBA(code: string): any {
    return evalVBASingle(code);
}

function runFunc(code: string, name: string, args: any[] = []): any {
    return evalVBA(code).callProcedure(name, args);
}

// --- 1. LTrim ---
{
    const code = `
    Function TestLTrim()
        TestLTrim = LTrim("   Hello   ")
    End Function
    `;
    assert.strictEqual(runFunc(code, 'TestLTrim'), "Hello   ", 'LTrim removes leading spaces');
}
console.log('[PASS] LTrim');

// --- 2. RTrim ---
{
    const code = `
    Function TestRTrim()
        TestRTrim = RTrim("   Hello   ")
    End Function
    `;
    assert.strictEqual(runFunc(code, 'TestRTrim'), "   Hello", 'RTrim removes trailing spaces');
}
console.log('[PASS] RTrim');

// --- 3. Trim ---
{
    const code = `
    Function TestTrim()
        TestTrim = Trim("   Hello   ")
    End Function
    `;
    assert.strictEqual(runFunc(code, 'TestTrim'), "Hello", 'Trim removes both leading and trailing spaces');
}
console.log('[PASS] Trim');

// --- 4. Empty and Null ---
{
    const code = `
    Function TestEmpty()
        TestEmpty = LTrim("")
    End Function
    Function TestNull()
        TestNull = RTrim(Null)
    End Function
    `;
    assert.strictEqual(runFunc(code, 'TestEmpty'), "", 'LTrim of empty string');
    assert.strictEqual(runFunc(code, 'TestNull'), vbaNull, 'RTrim of Null is Null');
}
console.log('[PASS] Empty and Null');

console.log('\n✅ LTrim / RTrim / Trim: 全テスト通過');
