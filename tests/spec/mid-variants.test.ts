import { evalVBASingle, assert } from '../../test-libs/test-runner';

function evalVBA(code: string): any {
    return evalVBASingle(code);
}

function runFunc(code: string, name: string, args: any[] = []): any {
    const ev = evalVBA(code);
    return ev.callProcedure(name, args);
}

console.log("--- Starting Mid/Left/Right Variants Tests ---");

// 1. Mid$ (Function)
assert.strictEqual(runFunc('Function Test() : Test = Mid$("ABC", 2, 1) : End Function', 'Test'), "B", 'Mid$ Function');

// 2. MidB (Function) - In this implementation, characters are treated as 2 bytes (UTF-16)
assert.strictEqual(runFunc('Function Test() : Test = MidB("ABC", 3, 2) : End Function', 'Test'), "B", 'MidB Function');

// 3. MidB$ (Function)
assert.strictEqual(runFunc('Function Test() : Test = MidB$("ABC", 1, 2) : End Function', 'Test'), "A", 'MidB$ Function');

// 4. LeftB
assert.strictEqual(runFunc('Function Test() : Test = LeftB("ABC", 2) : End Function', 'Test'), "A", 'LeftB Function');

// 5. RightB
assert.strictEqual(runFunc('Function Test() : Test = RightB("ABC", 2) : End Function', 'Test'), "C", 'RightB Function');

// 6. Mid Statement (Assignment) variants
{
    const code = `
        Function TestMidS()
            Dim s As String
            s = "ABCDE"
            Mid$(s, 2, 2) = "XY"
            TestMidS = s
        End Function
    `;
    assert.strictEqual(runFunc(code, 'TestMidS'), "AXYDE", 'Mid$ Statement');
}

{
    const code = `
        Function TestMidB()
            Dim s As String
            s = "ABCDE"
            MidB(s, 3, 2) = "Z"
            TestMidB = s
        End Function
    `;
    // MidB(s, 3, 2) starts at byte 3 (char 2 'B'), replaces 2 bytes (1 char).
    assert.strictEqual(runFunc(code, 'TestMidB'), "AZCDE", 'MidB Statement');
}

console.log("✅ Mid/Left/Right Variants: All tests passed!");
