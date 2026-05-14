import { Lexer } from '../../src/compiler/lexer';
import { Parser } from '../../src/compiler/parser';
import { Evaluator } from '../../src/compiler/evaluator';
import { assert } from '../../test-libs/test-runner';

function evalVBA(code: string): Evaluator {
    const tokens = new Lexer(code).tokenize();
    const ast = new Parser(tokens).parse();
    const ev = new Evaluator(console.log);
    ev.evaluate(ast);
    return ev;
}

function runFunc(code: string, name: string, args: any[] = []): any {
    return evalVBA(code).callProcedure(name, args);
}

// Test 1: Integer overflow - assignment
{
    const code = `
        Function TestIntegerOverflowPos()
            Dim x As Integer
            x = 32768  ' Should throw Error 6
            TestIntegerOverflowPos = x
        End Function
    `;

    try {
        runFunc(code, 'TestIntegerOverflowPos');
        console.log('[FAIL] Integer overflow positive not detected');
    } catch (e: any) {
        if (e && e.type === 'VbaError' && e.number === 6) {
            console.log('[PASS] Integer overflow positive detected');
        } else {
            console.log('[FAIL] Wrong error type:', e.message);
        }
    }
}

// Test 2: Integer overflow - negative
{
    const code = `
        Function TestIntegerOverflowNeg()
            Dim x As Integer
            x = -32769  ' Should throw Error 6
            TestIntegerOverflowNeg = x
        End Function
    `;

    try {
        runFunc(code, 'TestIntegerOverflowNeg');
        console.log('[FAIL] Integer overflow negative not detected');
    } catch (e: any) {
        if (e && e.type === 'VbaError' && e.number === 6) {
            console.log('[PASS] Integer overflow negative detected');
        } else {
            console.log('[FAIL] Wrong error type:', e.message);
        }
    }
}

// Test 3: Integer within range
{
    const code = `
        Function TestIntegerValidRange()
            Dim x As Integer
            x = 32767
            x = -32768
            TestIntegerValidRange = x
        End Function
    `;

    try {
        const result = runFunc(code, 'TestIntegerValidRange');
        assert.strictEqual(result, -32768, 'Integer valid range works');
        console.log('[PASS] Integer valid range');
    } catch (e: any) {
        console.log('[FAIL] Integer valid range threw error:', e.message);
    }
}

// Test 4: Long overflow - positive
{
    const code = `
        Function TestLongOverflowPos()
            Dim x As Long
            x = 2147483648  ' Should throw Error 6
            TestLongOverflowPos = x
        End Function
    `;

    try {
        runFunc(code, 'TestLongOverflowPos');
        console.log('[FAIL] Long overflow positive not detected');
    } catch (e: any) {
        if (e && e.type === 'VbaError' && e.number === 6) {
            console.log('[PASS] Long overflow positive detected');
        } else {
            console.log('[FAIL] Wrong error type:', e.message);
        }
    }
}

// Test 5: Long overflow - negative
{
    const code = `
        Function TestLongOverflowNeg()
            Dim x As Long
            x = -2147483649  ' Should throw Error 6
            TestLongOverflowNeg = x
        End Function
    `;

    try {
        runFunc(code, 'TestLongOverflowNeg');
        console.log('[FAIL] Long overflow negative not detected');
    } catch (e: any) {
        if (e && e.type === 'VbaError' && e.number === 6) {
            console.log('[PASS] Long overflow negative detected');
        } else {
            console.log('[FAIL] Wrong error type:', e.message);
        }
    }
}

// Test 6: Long within range
{
    const code = `
        Function TestLongValidRange()
            Dim x As Long
            x = 2147483647
            x = -2147483648
            TestLongValidRange = x
        End Function
    `;

    try {
        const result = runFunc(code, 'TestLongValidRange');
        assert.strictEqual(result, -2147483648, 'Long valid range works');
        console.log('[PASS] Long valid range');
    } catch (e: any) {
        console.log('[FAIL] Long valid range threw error:', e.message);
    }
}

// Test 7: CInt overflow
{
    const code = `
        Function TestCIntOverflow()
            Dim x
            x = CInt(32768)  ' Should throw Error 6
            TestCIntOverflow = x
        End Function
    `;

    try {
        runFunc(code, 'TestCIntOverflow');
        console.log('[FAIL] CInt overflow not detected');
    } catch (e: any) {
        if (e && e.type === 'VbaError' && e.number === 6) {
            console.log('[PASS] CInt overflow detected');
        } else {
            console.log('[FAIL] Wrong error:', e.message);
        }
    }
}

// Test 8: CLng overflow
{
    const code = `
        Function TestCLngOverflow()
            Dim x
            x = CLng(2147483648)  ' Should throw Error 6
            TestCLngOverflow = x
        End Function
    `;

    try {
        runFunc(code, 'TestCLngOverflow');
        console.log('[FAIL] CLng overflow not detected');
    } catch (e: any) {
        if (e && e.type === 'VbaError' && e.number === 6) {
            console.log('[PASS] CLng overflow detected');
        } else {
            console.log('[FAIL] Wrong error:', e.message);
        }
    }
}

// Test 9: Byte overflow
{
    const code = `
        Function TestByteOverflow()
            Dim x As Byte
            x = 256  ' Should throw Error 6
            TestByteOverflow = x
        End Function
    `;

    try {
        runFunc(code, 'TestByteOverflow');
        console.log('[FAIL] Byte overflow not detected');
    } catch (e: any) {
        if (e && e.type === 'VbaError' && e.number === 6) {
            console.log('[PASS] Byte overflow detected');
        } else {
            console.log('[FAIL] Wrong error:', e.message);
        }
    }
}

// Test 10: Real-world pattern - calculations resulting in overflow
{
    const code = `
        Function TestCalculationOverflow()
            Dim x As Integer
            Dim y As Integer
            y = 30000
            x = y + 5000  ' 30000 + 5000 = 35000, exceeds Integer range
            TestCalculationOverflow = x
        End Function
    `;

    try {
        runFunc(code, 'TestCalculationOverflow');
        console.log('[FAIL] Calculation overflow not detected');
    } catch (e: any) {
        if (e && e.type === 'VbaError' && e.number === 6) {
            console.log('[PASS] Calculation overflow detected');
        } else {
            console.log('[FAIL] Wrong error:', e.message);
        }
    }
}

console.log('\n✅ Overflow Detection: テスト完了');
