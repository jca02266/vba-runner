import { evalVBASingle, assert, vbaNull } from '../../test-libs/test-runner';

function evalVBA(code: string): any {
    return evalVBASingle(code);
}

function runFunc(code: string, name: string, args: any[] = []): any {
    return evalVBA(code).callProcedure(name, args);
}

// 基本動作
{
    const code = `
        Function Test1()
            Test1 = InStrRev("alphabet", "a")
        End Function
        Function Test2()
            Test2 = InStrRev("alphabet", "a", 5)
        End Function
        Function Test3()
            Test3 = InStrRev("alphabet", "z")
        End Function
        Function Test4()
            Test4 = InStrRev("alphabet", "A")
        End Function
        Function Test5()
            Test5 = InStrRev("alphabet", "A", -1, 1) ' vbTextCompare
        End Function
    `;
    assert.strictEqual(runFunc(code, 'Test1'), 5, 'InStrRev: last "a" in "alphabet" should be at 5');
    // Wait, "alphabet" has "a" at 1. 
    // "banana" -> InStrRev("banana", "a") -> 6
    
    const code2 = `
        Function Test6()
            Test6 = InStrRev("banana", "a")
        End Function
        Function Test7()
            Test7 = InStrRev("banana", "a", 5)
        End Function
    `;
    assert.strictEqual(runFunc(code2, 'Test6'), 6, 'InStrRev: "banana", "a" -> 6');
    assert.strictEqual(runFunc(code2, 'Test7'), 4, 'InStrRev: "banana", "a", 5 -> 4');
    
    assert.strictEqual(runFunc(code, 'Test3'), 0, 'InStrRev: not found -> 0');
    assert.strictEqual(runFunc(code, 'Test4'), 0, 'InStrRev: case sensitive by default');
    assert.strictEqual(runFunc(code, 'Test5'), 5, 'InStrRev: case insensitive with vbTextCompare');
    console.log('[PASS] 基本動作');
}

// エッジケース
{
    const code = `
        Function TestNull1()
            TestNull1 = InStrRev(Null, "a")
        End Function
        Function TestNull2()
            TestNull2 = InStrRev("a", Null)
        End Function
        Function TestEmpty1()
            TestEmpty1 = InStrRev("", "a")
        End Function
        Function TestEmpty2()
            TestEmpty2 = InStrRev("abc", "")
        End Function
        Function TestStartOver()
            TestStartOver = InStrRev("abc", "a", 10)
        End Function
    `;
    assert.strictEqual(runFunc(code, 'TestNull1'), vbaNull, 'InStrRev: StringCheck is Null -> Null');
    assert.strictEqual(runFunc(code, 'TestNull2'), vbaNull, 'InStrRev: StringMatch is Null -> Null');
    assert.strictEqual(runFunc(code, 'TestEmpty1'), 0, 'InStrRev: StringCheck is empty -> 0');
    // In VBA: InStrRev("abc", "") with default start (-1) returns 3 (length of StringCheck)
    assert.strictEqual(runFunc(code, 'TestEmpty2'), 3, 'InStrRev: StringMatch is empty -> Len(StringCheck)');
    assert.strictEqual(runFunc(code, 'TestStartOver'), 0, 'InStrRev: Start > Len(StringCheck) -> 0');
    console.log('[PASS] エッジケース');
}

console.log('\n✅ InStrRev: 全テスト通過');
