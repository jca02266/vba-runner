/**
 * Identifier type-declaration characters (MS-VBAL §3.3.5.2)
 * %=Integer  &=Long  ^=LongLong  !=Single  #=Double  @=Currency  $=String(identifier-only)
 *
 * '!' and '^' are also operators (bang member-access / exponentiation).
 * They are type suffixes only when NOT immediately followed by alphanumeric or '_'.
 */
import { evalVBASingle, assert } from '../../test-libs/test-runner';

function runFunc(code: string, name: string, args: any[] = []): any {
    const ev = evalVBASingle(code, { onPrint: () => {} });
    return ev.callProcedure(name, args);
}

console.log('--- Starting Identifier Type Suffix Tests ---');

// --- 1. Basic type suffixes ---
{
    const code = `
        Function TestPercent()
            Dim x%
            x = 32767
            TestPercent = x
        End Function
        Function TestAmpersand()
            Dim x&
            x = 100000
            TestAmpersand = x
        End Function
        Function TestHash()
            Dim x#
            x = 3.14
            TestHash = x
        End Function
        Function TestDollar()
            Dim x$
            x = "hello"
            TestDollar = x
        End Function
    `;
    const ev = evalVBASingle(code, { onPrint: () => {} });
    assert.strictEqual(ev.callProcedure('TestPercent', []), 32767, '% = Integer');
    assert.strictEqual(ev.callProcedure('TestAmpersand', []), 100000, '& = Long');
    assert.strictEqual(ev.callProcedure('TestHash', []), 3.14, '# = Double');
    assert.strictEqual(ev.callProcedure('TestDollar', []), 'hello', '$ = String');
    console.log('[PASS] Basic type suffixes (%, &, #, $)');
}

// --- 2. ! as Single type suffix (not followed by alphanumeric) ---
{
    const r = runFunc(`
        Function TestSingle()
            Dim x!
            x = 1.5
            TestSingle = x
        End Function
    `, 'TestSingle');
    assert.strictEqual(r, 1.5, 'x! declares Single');
    console.log('[PASS] ! as Single type suffix');
}

// --- 3. ^ as LongLong type suffix (not followed by alphanumeric) ---
{
    const r = runFunc(`
        Function TestLongLong()
            Dim x^
            x = 42
            TestLongLong = x
        End Function
    `, 'TestLongLong');
    assert.strictEqual(r, 42n, 'x^ declares LongLong (BigInt value)');
    console.log('[PASS] ^ as LongLong type suffix');
}

// --- 4. ! as bang member-access operator (followed by identifier) ---
{
    const r = runFunc(`
        Function TestBangAccess()
            Dim dict As Object
            Set dict = CreateObject("Scripting.Dictionary")
            dict.Add "Key", 99
            TestBangAccess = dict!Key
        End Function
    `, 'TestBangAccess');
    assert.strictEqual(r, 99, 'dict!Key uses bang access, not type suffix');
    console.log('[PASS] ! as bang member-access (dict!Key)');
}

// --- 5. ^ as exponentiation operator (followed by digit) ---
{
    const r = runFunc(`
        Function TestExponent()
            Dim x
            x = 3
            TestExponent = x^2
        End Function
    `, 'TestExponent');
    assert.strictEqual(r, 9, 'x^2 uses ^ as exponentiation, not type suffix');
    console.log('[PASS] ^ as exponentiation (x^2)');
}

// --- 6. Type suffix in function name (return type) ---
{
    const r = runFunc(`
        Function GetVal!()
            GetVal = 2.5
        End Function
    `, 'GetVal');
    assert.strictEqual(r, 2.5, 'Function GetVal! returns Single');
    console.log('[PASS] ! as function return type suffix');
}

console.log('\n✅ Identifier Type Suffix: 全テスト通過');
