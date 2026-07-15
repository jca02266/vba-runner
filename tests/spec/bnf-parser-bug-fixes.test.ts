/**
 * BNF監査で発見したパーサーバグの修正テスト (§5.4 各種)
 */
import { evalVBASingle, assert } from '../../test-libs/test-runner';
import { MemoryFileSystem } from '../../src/engine/filesystem';

// ─── §5.4.4.3 On Error GoTo -1 ───────────────────────────────────────────────
{
    const code = `
        Function TestGoToMinusOne() As Integer
            TestGoToMinusOne = 1
            On Error GoTo -1
            TestGoToMinusOne = 2
        End Function
    `;
    const ev = evalVBASingle(code);
    const result = ev.callProcedure('TestGoToMinusOne', []);
    // Just verify parsing and execution succeed; On Error GoTo -1 resets error state
    assert.strictEqual(result, 2, 'On Error GoTo -1 は構文エラーなく実行できる');
    console.log('[PASS] On Error GoTo -1');
}

// ─── §5.4.3.4 Erase 複数要素 ─────────────────────────────────────────────────
{
    const code = `
        Public arr1() As Integer
        Public arr2() As Integer
        Sub TestEraseMultiple()
            ReDim arr1(1)
            arr1(0) = 10: arr1(1) = 20
            ReDim arr2(1)
            arr2(0) = 30: arr2(1) = 40
            Erase arr1, arr2
        End Sub
    `;
    const ev = evalVBASingle(code);
    ev.callProcedure('TestEraseMultiple', []);
    const a1 = ev.env.get('arr1');
    const a2 = ev.env.get('arr2');
    assert.strictEqual(a1.length, 0, 'Erase arr1, arr2 — arr1 が解放される');
    assert.strictEqual(a2.length, 0, 'Erase arr1, arr2 — arr2 が解放される');
    console.log('[PASS] Erase 複数要素');
}

// ─── §5.4.3.4 Erase 3要素 ────────────────────────────────────────────────────
{
    const code = `
        Public a() As Integer
        Public b() As Integer
        Public c() As Integer
        Sub TestEraseThree()
            ReDim a(0): ReDim b(0): ReDim c(0)
            a(0) = 1: b(0) = 2: c(0) = 3
            Erase a, b, c
        End Sub
    `;
    const ev = evalVBASingle(code);
    ev.callProcedure('TestEraseThree', []);
    assert.strictEqual(ev.env.get('a').length, 0, 'Erase 3要素 — a が解放される');
    assert.strictEqual(ev.env.get('b').length, 0, 'Erase 3要素 — b が解放される');
    assert.strictEqual(ev.env.get('c').length, 0, 'Erase 3要素 — c が解放される');
    console.log('[PASS] Erase 3要素');
}

// ─── §5.4.3.3 ReDim 複数変数 ─────────────────────────────────────────────────
{
    const code = `
        Public arr1() As Integer
        Public arr2() As String
        Sub TestReDimMultiple()
            ReDim arr1(2), arr2(4)
            arr1(0) = 1: arr1(2) = 3
            arr2(0) = "a": arr2(4) = "e"
        End Sub
    `;
    const ev = evalVBASingle(code);
    ev.callProcedure('TestReDimMultiple', []);
    const a1 = ev.env.get('arr1');
    const a2 = ev.env.get('arr2');
    assert.strictEqual(a1.length, 3, 'ReDim arr1(2) → length 3');
    assert.strictEqual(a2.length, 5, 'ReDim arr2(4) → length 5');
    assert.strictEqual(a1[0], 1, 'arr1(0) = 1');
    assert.strictEqual(a1[2], 3, 'arr1(2) = 3');
    assert.strictEqual(a2[0], 'a', 'arr2(0) = "a"');
    assert.strictEqual(a2[4], 'e', 'arr2(4) = "e"');
    console.log('[PASS] ReDim 複数変数');
}

// ─── §5.4.3.3 ReDim Preserve 複数変数 ────────────────────────────────────────
{
    const code = `
        Public arr1() As Integer
        Public arr2() As Integer
        Sub TestReDimPreserveMultiple()
            ReDim arr1(1): arr1(0) = 10: arr1(1) = 20
            ReDim arr2(1): arr2(0) = 30: arr2(1) = 40
            ReDim Preserve arr1(2), arr2(2)
            arr1(2) = 30
            arr2(2) = 50
        End Sub
    `;
    const ev = evalVBASingle(code);
    ev.callProcedure('TestReDimPreserveMultiple', []);
    const a1 = ev.env.get('arr1');
    const a2 = ev.env.get('arr2');
    assert.strictEqual(a1[0], 10, 'Preserve: arr1(0) 保持');
    assert.strictEqual(a1[2], 30, 'Preserve: arr1(2) 追加');
    assert.strictEqual(a2[0], 30, 'Preserve: arr2(0) 保持');
    assert.strictEqual(a2[2], 50, 'Preserve: arr2(2) 追加');
    console.log('[PASS] ReDim Preserve 複数変数');
}

// ─── §5.4.3.2 Const 複数定数 ─────────────────────────────────────────────────
{
    const code = `
        Const A = 1, B = 2
        Function GetSum() As Long
            GetSum = A + B
        End Function
    `;
    const ev = evalVBASingle(code);
    const result = ev.callProcedure('GetSum', []);
    assert.strictEqual(result, 3, 'Const A = 1, B = 2 → A + B = 3');
    console.log('[PASS] Const 複数定数 (基本)');
}

// ─── §5.4.3.2 Const 複数定数 — 型付き ────────────────────────────────────────
{
    const code = `
        Const MAX As Long = 100, MIN As Long = 0, STEP_VAL As Long = 5
        Function GetRange() As Long
            GetRange = MAX - MIN
        End Function
        Function GetStep() As Long
            GetStep = STEP_VAL
        End Function
    `;
    const ev = evalVBASingle(code);
    assert.strictEqual(ev.callProcedure('GetRange', []), 100, 'Const 型付き複数定数 — MAX - MIN = 100');
    assert.strictEqual(ev.callProcedure('GetStep', []), 5, 'Const 型付き複数定数 — STEP_VAL = 5');
    console.log('[PASS] Const 複数定数 (型付き)');
}

// ─── §5.4.3.2 Const 複数定数 — Public/Private スコープ ──────────────────────
{
    const code = `
        Public Const X = 10, Y = 20
        Function GetX() As Long: GetX = X: End Function
        Function GetY() As Long: GetY = Y: End Function
    `;
    const ev = evalVBASingle(code);
    assert.strictEqual(ev.callProcedure('GetX', []), 10, 'Public Const X = 10');
    assert.strictEqual(ev.callProcedure('GetY', []), 20, 'Public Const Y = 20');
    console.log('[PASS] Const 複数定数 (Public スコープ)');
}

// ─── §5.4.5.1 Open 「For mode」省略 ──────────────────────────────────────────
{
    const vfs = new MemoryFileSystem();
    const code = `
        Sub TestOpenNoMode()
            Open "test.txt" As #1
            Close #1
        End Sub
    `;
    const ev = evalVBASingle(code, { fs: vfs });
    ev.callProcedure('TestOpenNoMode', []);
    console.log('[PASS] Open 「For mode」省略 (構文チェック)');
}

// ─── §5.4.5.1.1 Print # なし ─────────────────────────────────────────────────
{
    const vfs = new MemoryFileSystem();
    const lines: string[] = [];
    const code = `
        Sub TestPrintNoHash()
            Open "out.txt" For Output As #1
            Print 1, "hello"
            Print 1, "world"
            Close #1
            Open "out.txt" For Input As #2
            Dim s As String
            Line Input #2, s
            Debug.Print s
            Line Input #2, s
            Debug.Print s
            Close #2
        End Sub
    `;
    const ev = evalVBASingle(code, { fs: vfs, onPrint: (o) => lines.push(o.trim()) });
    ev.callProcedure('TestPrintNoHash', []);
    assert.strictEqual(lines[0], 'hello', 'Print 1, "hello" → hello');
    assert.strictEqual(lines[1], 'world', 'Print 1, "world" → world');
    console.log('[PASS] Print # なし');
}

// ─── §5.4.5.1.1 Input # なし ─────────────────────────────────────────────────
{
    const vfs2 = new MemoryFileSystem();
    const lines2: string[] = [];
    const code2 = `
        Sub TestInputNoHash()
            Open "data.txt" For Output As #1
            Write #1, "value1"
            Close #1
            Open "data.txt" For Input As #2
            Dim s As String
            Input 2, s
            Debug.Print s
            Close #2
        End Sub
    `;
    const ev2 = evalVBASingle(code2, { fs: vfs2, onPrint: (o) => lines2.push(o.trim()) });
    ev2.callProcedure('TestInputNoHash', []);
    assert.strictEqual(lines2[0], 'value1', 'Input 2, s → value1');
    console.log('[PASS] Input # なし');
}

// ─── §5.4.5.3 Seek # なし ────────────────────────────────────────────────────
{
    const vfs3 = new MemoryFileSystem();
    const code3 = `
        Sub TestSeekNoHash()
            Open "seek.txt" For Random As #1 Len = 10
            Seek 1, 1
            Close #1
        End Sub
    `;
    const ev3 = evalVBASingle(code3, { fs: vfs3 });
    ev3.callProcedure('TestSeekNoHash', []);
    console.log('[PASS] Seek # なし');
}

// ─── §5.4.5.8 Print # あり (既存動作の保持確認) ──────────────────────────────
{
    const vfs4 = new MemoryFileSystem();
    const lines4: string[] = [];
    const code4 = `
        Sub TestPrintWithHash()
            Open "hash.txt" For Output As #1
            Print #1, "withHash"
            Close #1
            Open "hash.txt" For Input As #2
            Dim s As String
            Line Input #2, s
            Debug.Print s
            Close #2
        End Sub
    `;
    const ev4 = evalVBASingle(code4, { fs: vfs4, onPrint: (o) => lines4.push(o.trim()) });
    ev4.callProcedure('TestPrintWithHash', []);
    assert.strictEqual(lines4[0], 'withHash', 'Print #1 (既存 # あり) 動作保持');
    console.log('[PASS] Print # あり (既存動作の保持)');
}

// ─── §5.4.3.3 ReDim obj.Arr (メンバーアクセス対象) ──────────────────────────
{
    const code = `
        Class Container
            Public Items() As Integer
        End Class
        Sub TestReDimMember()
            Dim c As New Container
            ReDim c.Items(3)
            c.Items(0) = 10
            c.Items(3) = 40
        End Sub
    `;
    const ev = evalVBASingle(code);
    ev.callProcedure('TestReDimMember', []);
    // Verify via With block access
    console.log('[PASS] ReDim obj.Arr (メンバーアクセス対象)');
}

// ─── §5.4.3.3 ReDim .Items (With式対象) ─────────────────────────────────────
{
    const code = `
        Class Container
            Public Items() As Integer
        End Class
        Function TestReDimWith() As Long
            Dim c As New Container
            With c
                ReDim .Items(2)
            End With
            TestReDimWith = UBound(c.Items)
        End Function
    `;
    const ev = evalVBASingle(code);
    const result = ev.callProcedure('TestReDimWith', []);
    assert.strictEqual(result, 2, 'ReDim .Items(2) → UBound = 2');
    console.log('[PASS] ReDim .Items (With式対象)');
}

// ─── Global キーワード (§5.2.3.1) ────────────────────────────────────────────
{
    // Global は Public の別名
    const code = `
        Global x As Integer
        Global Const MAX As Long = 100
        Global Sub TestGlobal()
            x = MAX
        End Sub
    `;
    const ev = evalVBASingle(code);
    ev.callProcedure('TestGlobal', []);
    assert.strictEqual(ev.env.get('x'), 100, 'Global x = Global Const MAX = 100');
    console.log('[PASS] Global キーワード (変数・Const・Sub)');
}

// ─── TypeOf ... Is Library.ClassName (ドット修飾型名) ─────────────────────────
{
    const code = `
        Function TestTypeOf() As Boolean
            Dim d As Object
            Set d = CreateObject("Scripting.Dictionary")
            TestTypeOf = TypeOf d Is Scripting.Dictionary
        End Function
    `;
    const ev = evalVBASingle(code);
    const result = ev.callProcedure('TestTypeOf', []);
    assert.strictEqual(result, -1, 'TypeOf d Is Scripting.Dictionary → True (-1)');
    console.log('[PASS] TypeOf ... Is Library.ClassName (ドット修飾型名)');
}

console.log('\n✅ BNF パーサーバグ修正: 全テスト通過');
