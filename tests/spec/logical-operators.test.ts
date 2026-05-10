import { VBATest, assert } from '../ts/test-runner';

async function main() {
    console.log("--- Starting Logical/Bitwise Operators Tests ---");
    const vbaTest = new VBATest('sample/src/vba/Sample.vba');

    console.log("[Test] And");
    assert.strictEqual(vbaTest.eval("-1 And -1"), -1); // True And True
    assert.strictEqual(vbaTest.eval("-1 And 0"), 0);   // True And False
    assert.strictEqual(vbaTest.eval("5 And 3"), 1);    // 101 And 011 = 001

    console.log("[Test] Or");
    assert.strictEqual(vbaTest.eval("-1 Or 0"), -1);  // True Or False
    assert.strictEqual(vbaTest.eval("0 Or 0"), 0);    // False Or False
    assert.strictEqual(vbaTest.eval("5 Or 3"), 7);    // 101 Or 011 = 111

    console.log("[Test] Xor");
    assert.strictEqual(vbaTest.eval("-1 Xor -1"), 0); // True Xor True = False
    assert.strictEqual(vbaTest.eval("-1 Xor 0"), -1);  // True Xor False = True
    assert.strictEqual(vbaTest.eval("5 Xor 3"), 6);    // 101 Xor 011 = 110

    console.log("[Test] Eqv");
    assert.strictEqual(vbaTest.eval("-1 Eqv -1"), -1); // True Eqv True = True
    assert.strictEqual(vbaTest.eval("-1 Eqv 0"), 0);   // True Eqv False = False
    assert.strictEqual(vbaTest.eval("5 Eqv 3"), -7);   // Not(101 Xor 011) = Not(110) = ...11111001

    console.log("[Test] Imp");
    assert.strictEqual(vbaTest.eval("-1 Imp -1"), -1); // True Imp True = True
    assert.strictEqual(vbaTest.eval("-1 Imp 0"), 0);   // True Imp False = False
    assert.strictEqual(vbaTest.eval("0 Imp 0"), -1);   // False Imp False = True
    assert.strictEqual(vbaTest.eval("5 Imp 3"), -5);   // (Not 5) Or 3 = -6 Or 3 = -5

    console.log("[Test] Not");
    assert.strictEqual(vbaTest.eval("Not -1"), 0);    // Not True = False
    assert.strictEqual(vbaTest.eval("Not 0"), -1);    // Not False = True
    assert.strictEqual(vbaTest.eval("Not 5"), -6);    // Not 5 = -6

    console.log("[Test] Precedence");
    // And > Or > Xor > Eqv > Imp
    assert.strictEqual(vbaTest.eval("-1 Or -1 And 0"), -1); // -1 Or 0 = -1
    assert.strictEqual(vbaTest.eval("0 Xor -1 Or -1"), -1); // 0 Xor -1 = -1
    assert.strictEqual(vbaTest.eval("-1 Eqv -1 Xor -1"), 0); // -1 Eqv 0 = 0
    assert.strictEqual(vbaTest.eval("0 Imp -1 Eqv 0"), -1); // 0 Imp 0 = -1

    console.log("[Test] Like");
    assert.strictEqual(vbaTest.eval(`"abc" Like "a*"`), -1);      // True
    assert.strictEqual(vbaTest.eval(`"abc" Like "?b?"`), -1);     // True
    assert.strictEqual(vbaTest.eval(`"a1b" Like "a#b"`), -1);     // True
    assert.strictEqual(vbaTest.eval(`"a2b" Like "a[123]b"`), -1); // True
    assert.strictEqual(vbaTest.eval(`"a4b" Like "a[123]b"`), 0);  // False
    assert.strictEqual(vbaTest.eval(`"a4b" Like "a[!123]b"`), -1); // True
    assert.strictEqual(vbaTest.eval(`"ABC" Like "abc"`), -1);      // Case-insensitive

    console.log("--- All Logical Operators Tests passed! ---");
}

main().catch(e => {
    console.error(e);
    process.exit(1);
});
