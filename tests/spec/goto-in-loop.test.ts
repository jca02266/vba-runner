/**
 * For / ForEach / DoWhile / While ループ本体内のラベルへの GoTo テスト
 * MS-VBAL §5.2.1.2: ラベルのスコープは手続き全体（ループ内外を問わない）
 */
import { evalVBASingle, assert } from '../../test-libs/test-runner';

function run(code: string, name: string, args: any[] = []): any {
    return evalVBASingle(code).callProcedure(name, args);
}

// --- 1. For ループ内ラベルへの GoTo（スキップ） ---
const forSkipCode = `
Function CountSkipped(n As Long) As Long
    Dim i As Long
    Dim count As Long
    count = 0
    For i = 1 To n
        If i Mod 2 = 0 Then GoTo Skip
        count = count + 1
Skip:
    Next i
    CountSkipped = count
End Function
`;
assert.strictEqual(run(forSkipCode, 'CountSkipped', [6]), 3, 'For GoTo: 偶数スキップで奇数のみカウント');
assert.strictEqual(run(forSkipCode, 'CountSkipped', [0]), 0, 'For GoTo: n=0 で 0');

// --- 2. GoTo が実際に実行されるケース（重複スキップ）---
const forDupSkipCode = `
Function DedupeCount() As Long
    Dim i As Long
    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")
    Dim count As Long
    count = 0
    For i = 1 To 5
        If seen.exists(i Mod 3) Then GoTo SkipRow
        seen.Add i Mod 3, True
        count = count + 1
SkipRow:
    Next i
    DedupeCount = count
End Function
`;
assert.strictEqual(run(forDupSkipCode, 'DedupeCount', []), 3, 'For GoTo: Dictionary 重複スキップ');

// --- 3. For Each ループ内ラベルへの GoTo ---
const forEachCode = `
Function SumPositive() As Long
    Dim arr(4) As Long
    arr(0) = 1 : arr(1) = -2 : arr(2) = 3 : arr(3) = -4 : arr(4) = 5
    Dim v As Variant
    Dim total As Long
    total = 0
    For Each v In arr
        If v < 0 Then GoTo Skip
        total = total + v
Skip:
    Next v
    SumPositive = total
End Function
`;
assert.strictEqual(run(forEachCode, 'SumPositive', []), 9, 'ForEach GoTo: 負の値をスキップして合計');

// --- 4. Do While ループ内ラベルへの GoTo ---
const doWhileCode = `
Function CountOdd(n As Long) As Long
    Dim i As Long
    Dim count As Long
    i = 0
    count = 0
    Do While i < n
        i = i + 1
        If i Mod 2 = 0 Then GoTo Next2
        count = count + 1
Next2:
    Loop
    CountOdd = count
End Function
`;
assert.strictEqual(run(doWhileCode, 'CountOdd', [6]), 3, 'DoWhile GoTo: 偶数スキップ');

// --- 5. While ループ内ラベルへの GoTo ---
const whileCode = `
Function SumOdd(n As Long) As Long
    Dim i As Long
    Dim s As Long
    i = 0
    s = 0
    While i < n
        i = i + 1
        If i Mod 2 = 0 Then GoTo Skip
        s = s + i
Skip:
    Wend
    SumOdd = s
End Function
`;
assert.strictEqual(run(whileCode, 'SumOdd', [6]), 9, 'While GoTo: 奇数のみ合計');

// --- 6. GoTo のジャンプ先がループ先頭にある場合（スキップ対象なし）---
const noSkipCode = `
Function AllCount(n As Long) As Long
    Dim i As Long
    Dim count As Long
    count = 0
    For i = 1 To n
        If i > 100 Then GoTo Skip
        count = count + 1
Skip:
    Next i
    AllCount = count
End Function
`;
assert.strictEqual(run(noSkipCode, 'AllCount', [5]), 5, 'For GoTo: 条件未達 → 全カウント');

console.log('goto-in-loop: all tests passed');
