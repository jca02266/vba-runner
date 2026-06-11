import { evalVBASingle, assert } from '../../test-libs/test-runner';

function runFunc(code: string, name: string, args: any[] = []): any {
    const ev = evalVBASingle(code, { onPrint: () => {} });
    return ev.callProcedure(name, args);
}

// 1. 基本的なGoTo (前方へのジャンプ)
{
    const code = `
        Function TestForward()
            GoTo MyLabel
            TestForward = 1
            Exit Function
        MyLabel:
            TestForward = 2
        End Function
    `;
    assert.strictEqual(runFunc(code, 'TestForward'), 2, 'GoTo should jump forward');
}

// 2. 後方へのジャンプ (ループ的な動作)
{
    const code = `
        Function TestBackward()
            Dim i
            i = 0
        StartLoop:
            i = i + 1
            If i < 5 Then GoTo StartLoop
            TestBackward = i
        End Function
    `;
    assert.strictEqual(runFunc(code, 'TestBackward'), 5, 'GoTo should jump backward');
}

// 3. ブロック内からの脱出
{
    const code = `
        Function TestJumpOut()
            Dim i, total
            total = 0
            For i = 1 To 10
                total = total + i
                If i = 5 Then GoTo ExitPoint
            Next i
        ExitPoint:
            TestJumpOut = total
        End Function
    `;
    assert.strictEqual(runFunc(code, 'TestJumpOut'), 15, 'GoTo should jump out of For block');
}

// 4. 数値ラベル
{
    const code = `
        Function TestNumericLabel()
            GoTo 100
            TestNumericLabel = 1
            Exit Function
        100:
            TestNumericLabel = 2
        End Function
    `;
    assert.strictEqual(runFunc(code, 'TestNumericLabel'), 2, 'GoTo should support numeric labels');
}

// 5. コンテキストキーワードをラベル名に使う（Error:, Property: など）
{
    const code = `
        Function TestContextualKeywordLabel()
            On Error GoTo Error
            Err.Raise 1
            TestContextualKeywordLabel = 1
            Exit Function
        Error:
            TestContextualKeywordLabel = 2
        End Function
    `;
    assert.strictEqual(runFunc(code, 'TestContextualKeywordLabel'), 2, 'contextual keyword Error can be used as label name');
}

console.log('\n✅ GoTo Statement: 全テスト通過');
