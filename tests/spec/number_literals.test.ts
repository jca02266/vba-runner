import { evalVBASingle, assert } from '../../test-libs/test-runner';

function evalVBA(code: string): any {
    return evalVBASingle(code);
}

const code = `
    Function TestHex()
        TestHex = &HFF
    End Function

    Function TestHexLongSuffix()
        TestHexLongSuffix = &H1&
    End Function

    Function TestOctal()
        TestOctal = &O10
    End Function

    Function TestOctalLongSuffix()
        TestOctalLongSuffix = &O17&
    End Function

    Function TestOctalShort()
        TestOctalShort = &10
    End Function

    Function TestSuffixes()
        Dim results(5)
        results(0) = 123%   ' Integer
        results(1) = 123&   ' Long
        results(2) = 123.4! ' Single
        results(3) = 123.4# ' Double
        results(4) = 123.4@ ' Currency
        results(5) = 123^   ' LongLong
        TestSuffixes = results
    End Function

    Function TestStringEscape()
        TestStringEscape = "a""b"
    End Function

    Function TestScientific()
        TestScientific = 1.23E+2
    End Function
`;

const ev = evalVBA(code);

console.log('[Test Suite] リテラルの検証');

assert.strictEqual(ev.callProcedure('TestHex', []), 255, 'Hex &HFF -> 255');
assert.strictEqual(ev.callProcedure('TestHexLongSuffix', []), 1, 'Hex &H1& (Long suffix) -> 1');
assert.strictEqual(ev.callProcedure('TestOctal', []), 8, 'Octal &O10 -> 8');
assert.strictEqual(ev.callProcedure('TestOctalLongSuffix', []), 15, 'Octal &O17& (Long suffix) -> 15');
assert.strictEqual(ev.callProcedure('TestOctalShort', []), 8, 'Octal &10 -> 8');

// Hex$() / Oct$() 関数呼び出し（リテラルとは別に関数として動作すること）
{
    const e2 = evalVBA(`
        Function TestHexFn() As String: TestHexFn = Hex$(255): End Function
        Function TestHexFn0() As String: TestHexFn0 = Hex$(0): End Function
        Function TestOctFn() As String: TestOctFn = Oct$(8): End Function
    `);
    assert.strictEqual(e2.callProcedure('TestHexFn', []), 'FF', 'Hex$(255) = "FF"');
    assert.strictEqual(e2.callProcedure('TestHexFn0', []), '0', 'Hex$(0) = "0"');
    assert.strictEqual(e2.callProcedure('TestOctFn', []), '10', 'Oct$(8) = "10"');
}
assert.strictEqual(ev.callProcedure('TestStringEscape', []), 'a"b', 'String escape "a""b" -> a"b');
assert.strictEqual(ev.callProcedure('TestScientific', []), 123, 'Scientific 1.23E+2 -> 123');

const sfx = ev.callProcedure('TestSuffixes', []);
assert.strictEqual(sfx[0], 123, 'Suffix % -> 123');
assert.strictEqual(sfx[1], 123, 'Suffix & -> 123');
assert.strictEqual(sfx[2], 123.4, 'Suffix ! -> 123.4');
assert.strictEqual(sfx[3], 123.4, 'Suffix # -> 123.4');
assert.strictEqual(sfx[4], 123.4, 'Suffix @ -> 123.4');
assert.strictEqual(sfx[5], 123, 'Suffix ^ -> 123');

console.log('[PASS] 数値リテラルの検証');

console.log('\n✅ Number Literals: 全テスト通過');
