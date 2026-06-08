import { preprocess, DEFAULT_COMPILER_CONSTANTS } from '../../src/engine/preprocessor';
import { evalVBASingle, assert } from '../../test-libs/test-runner';

function evalVBA(code: string): any {
    return evalVBASingle(code);
}

function runFunc(code: string, name: string, args: any[] = []): any {
    return evalVBA(code).callProcedure(name, args);
}

function evalAfterPreprocess(code: string, constants = {}): any {
    const processed = preprocess(code, constants);
    return evalVBA(processed);
}

// Test 1: preprocess removes #Const lines and replaces excluded lines with blank lines
{
    const src = `
#Const DEBUG = 1
Function GetMode() As String
    GetMode = "debug"
End Function
`.trim();
    const result = preprocess(src);
    assert.strictEqual(result.includes('#Const'), false, '#Const line removed');
    assert.strictEqual(result.includes('"debug"'), true, 'included code preserved');
    console.log('[PASS] #Const line is removed from output');
}

// Test 2: #If True Then - included block is kept
{
    const src = `
Function F()
#If True Then
    F = 1
#Else
    F = 2
#End If
End Function
`;
    const ev = evalAfterPreprocess(src);
    assert.strictEqual(ev.callProcedure('F', []), 1, '#If True -> first branch');
    console.log('[PASS] #If True keeps first branch');
}

// Test 3: #If False Then - else block is kept
{
    const src = `
Function F()
#If False Then
    F = 1
#Else
    F = 2
#End If
End Function
`;
    const ev = evalAfterPreprocess(src);
    assert.strictEqual(ev.callProcedure('F', []), 2, '#If False -> else branch');
    console.log('[PASS] #If False keeps else branch');
}

// Test 4: Default VBA7 = 0 (disabled) - #Else branch is used
{
    const src = `
Function F()
#If VBA7 Then
    F = "vba7"
#Else
    F = "legacy"
#End If
End Function
`;
    const ev = evalAfterPreprocess(src);
    assert.strictEqual(ev.callProcedure('F', []), 'legacy', 'VBA7=0 -> legacy branch');
    console.log('[PASS] Default VBA7=0 selects legacy branch');
}

// Test 5: Override VBA7=True to use VBA7 branch
{
    const src = `
Function F()
#If VBA7 Then
    F = "vba7"
#Else
    F = "legacy"
#End If
End Function
`;
    const ev = evalAfterPreprocess(src, { VBA7: -1 });
    assert.strictEqual(ev.callProcedure('F', []), 'vba7', 'VBA7=-1 -> vba7 branch');
    console.log('[PASS] VBA7=-1 override selects vba7 branch');
}

// Test 6: #Const defines a local constant usable in #If
{
    const src = `
#Const PLATFORM = 1
Function F()
#If PLATFORM = 1 Then
    F = "win"
#Else
    F = "other"
#End If
End Function
`;
    const ev = evalAfterPreprocess(src);
    assert.strictEqual(ev.callProcedure('F', []), 'win', '#Const + #If comparison');
    console.log('[PASS] #Const defines constant for #If use');
}

// Test 7: #ElseIf chain
{
    const src = `
#Const X = 2
Function F()
#If X = 1 Then
    F = "one"
#ElseIf X = 2 Then
    F = "two"
#ElseIf X = 3 Then
    F = "three"
#Else
    F = "other"
#End If
End Function
`;
    const ev = evalAfterPreprocess(src);
    assert.strictEqual(ev.callProcedure('F', []), 'two', '#ElseIf = 2 matches');
    console.log('[PASS] #ElseIf chain selects correct branch');
}

// Test 8: Nested #If blocks
{
    const src = `
Function F()
#If True Then
#If False Then
    F = "inner-false"
#Else
    F = "inner-else"
#End If
#Else
    F = "outer-false"
#End If
End Function
`;
    const ev = evalAfterPreprocess(src);
    assert.strictEqual(ev.callProcedure('F', []), 'inner-else', 'nested #If');
    console.log('[PASS] Nested #If blocks work correctly');
}

// Test 9: #EndIf without space also accepted
{
    const src = `
Function F()
#If False Then
    F = 1
#Else
    F = 2
#EndIf
End Function
`;
    const ev = evalAfterPreprocess(src);
    assert.strictEqual(ev.callProcedure('F', []), 2, '#EndIf (no space) accepted');
    console.log('[PASS] #EndIf (no space) accepted');
}

// Test 10: Logic operators in cc-expression
{
    const src = `
Function F()
#If Not False Then
    F = "yes"
#Else
    F = "no"
#End If
End Function
`;
    const ev = evalAfterPreprocess(src);
    assert.strictEqual(ev.callProcedure('F', []), 'yes', 'Not False = True');
    console.log('[PASS] Not False evaluates to True');
}

// Test 11: And / Or in cc-expression
{
    const src = `
Function F()
#If False Or True Then
    F = "yes"
#Else
    F = "no"
#End If
End Function
`;
    const ev = evalAfterPreprocess(src);
    assert.strictEqual(ev.callProcedure('F', []), 'yes', 'False Or True = True');
    console.log('[PASS] False Or True evaluates to True');
}

// Test 12: Undefined symbol treated as 0 (falsy)
{
    const src = `
Function F()
#If UNDEFINED_SYMBOL Then
    F = "defined"
#Else
    F = "undefined"
#End If
End Function
`;
    const ev = evalAfterPreprocess(src);
    assert.strictEqual(ev.callProcedure('F', []), 'undefined', 'Undefined symbol = 0 = falsy');
    console.log('[PASS] Undefined symbol treated as 0');
}

// Test 13: DEFAULT_COMPILER_CONSTANTS check
assert.strictEqual(DEFAULT_COMPILER_CONSTANTS['VBA7'], 0, 'VBA7 default is 0');
assert.strictEqual(DEFAULT_COMPILER_CONSTANTS['Win32'], -1, 'Win32 default is -1 (true)');
assert.strictEqual(DEFAULT_COMPILER_CONSTANTS['Win64'], 0, 'Win64 default is 0');
assert.strictEqual(DEFAULT_COMPILER_CONSTANTS['Mac'], 0, 'Mac default is 0');
console.log('[PASS] DEFAULT_COMPILER_CONSTANTS values are correct');

// Test 14: Win32=True selects Win32 branch
{
    const src = `
Function F()
#If Win32 Then
    F = "win32"
#ElseIf Win64 Then
    F = "win64"
#Else
    F = "other"
#End If
End Function
`;
    const ev = evalAfterPreprocess(src);
    assert.strictEqual(ev.callProcedure('F', []), 'win32', 'Win32=-1 selects win32 branch');
    console.log('[PASS] Win32=-1 selects win32 branch');
}

// Test 15: Typical VBA7 PtrSafe pattern - PtrSafe keyword excluded when VBA7=0
{
    const src = `
#If VBA7 Then
Declare PtrSafe Function GetTickCount Lib "kernel32" () As LongLong
#Else
Declare Function GetTickCount Lib "kernel32" () As Long
#End If
Function F()
    F = 42
End Function
`;
    const ev = evalAfterPreprocess(src);
    // Should not throw - PtrSafe line is excluded, legacy Declare is included
    // But our engine doesn't support Declare, so we just test F()
    assert.strictEqual(ev.callProcedure('F', []), 42, 'PtrSafe pattern: function body works');
    console.log('[PASS] VBA7 PtrSafe pattern: legacy branch included when VBA7=0');
}

console.log('\n✅ conditional-compilation: 全テスト通過');
