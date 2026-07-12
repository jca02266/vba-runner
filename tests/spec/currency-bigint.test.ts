import { evalVBASingle, evalVBAModules, assert } from '../../test-libs/test-runner';

function runFunc(code: string, name: string, args: any[] = []): any {
    return evalVBASingle(code).callProcedure(name, args);
}

function cur(v: any): string { return String(v); }

// ── Phase 1: VbaCurrency BigInt 精度テスト ──────────────────────────────────

// 主要精度テスト: CCur(0.1) + CCur(0.2) = CCur(0.3) (float では 0.30000000000000004)
{
    const code = `
    Function Test()
        Dim a As Currency, b As Currency
        a = CCur(0.1)
        b = CCur(0.2)
        Test = a + b
    End Function
    `;
    const result = runFunc(code, 'Test');
    assert.strictEqual(cur(result), '0.3', 'CCur(0.1)+CCur(0.2) = 0.3 (no float error)');
    console.log('[PASS] 0.1 + 0.2 = 0.3 (BigInt 精度)');
}

// CStr 経由でも float 誤差なし
{
    const code = `
    Sub Test()
        Dim a As Currency, b As Currency
        a = CCur(0.1)
        b = CCur(0.2)
        Debug.Print CStr(a + b)
    End Sub
    `;
    let output = '';
    evalVBASingle(code, { onPrint: s => output += s + '\n' }).callProcedure('Test', []);
    assert.strictEqual(output.trim(), '0.3', 'CStr(CCur(0.1)+CCur(0.2)) = "0.3"');
    console.log('[PASS] CStr(CCur(0.1)+CCur(0.2)) = "0.3" (Debug.Print)');
}

// 大きな金額: MAX_SAFE_INTEGER(≈9.0×10^15) を超える領域
{
    const code = `
    Function Test()
        Dim a As Currency, b As Currency
        a = CCur("900000000000000.1234")
        b = CCur("0.0001")
        Test = a + b
    End Function
    `;
    const result = runFunc(code, 'Test');
    assert.strictEqual(cur(result), '900000000000000.1235', 'Large Currency addition (beyond MAX_SAFE_INTEGER)');
    console.log('[PASS] 大きな金額の加算 (MAX_SAFE_INTEGER 超え)');
}

// TypeName が Currency を返す (AST 推定でなく実値判定)
{
    const code = `
    Function Test()
        Dim c As Currency
        c = 1.5
        Test = TypeName(c)
    End Function
    `;
    const result = runFunc(code, 'Test');
    assert.strictEqual(result, 'Currency', 'TypeName(Currency var) = "Currency"');
    console.log('[PASS] TypeName(Currency) = "Currency" (実値判定)');
}

// VarType が 6 を返す (実値判定)
{
    const code = `
    Function Test()
        Dim c As Currency
        c = 1.5
        Test = VarType(c)
    End Function
    `;
    const result = runFunc(code, 'Test');
    assert.strictEqual(result, 6, 'VarType(Currency) = 6');
    console.log('[PASS] VarType(Currency) = 6 (実値判定)');
}

// CCur 演算後の TypeName も Currency
{
    const code = `
    Function Test()
        Test = TypeName(CCur(1.5) + CCur(2.5))
    End Function
    `;
    const result = runFunc(code, 'Test');
    assert.strictEqual(result, 'Currency', 'TypeName(CCur+CCur) = "Currency"');
    console.log('[PASS] TypeName(CCur(1.5)+CCur(2.5)) = "Currency"');
}

// Currency / Currency は Double
{
    const code = `
    Function Test()
        Test = TypeName(CCur(1) / CCur(2))
    End Function
    `;
    const result = runFunc(code, 'Test');
    assert.strictEqual(result, 'Double', 'Currency/Currency → Double');
    console.log('[PASS] Currency/Currency → Double');
}

// 乗算のバンカーズ丸め: 下位桁が5ちょうどのとき偶数側へ
{
    // CCur(0.0003) * CCur(0.5) = 0.00015 → 4桁で 0.0002 (偶数丸め)
    const code1 = `
    Function Test()
        Test = CCur(0.0003) * CCur(0.5)
    End Function
    `;
    const r1 = runFunc(code1, 'Test');
    assert.strictEqual(cur(r1), '0.0002', 'Banker round: 0.0003*0.5=0.00015 → 0.0002 (even)');
    console.log('[PASS] 乗算バンカーズ丸め: 0.0003*0.5=0.00015 → 0.0002');

    // CCur(0.0009) * CCur(0.5) = 0.00045 → 4桁で 0.0004 (偶数丸め)
    const code2 = `
    Function Test()
        Test = CCur(0.0009) * CCur(0.5)
    End Function
    `;
    const r2 = runFunc(code2, 'Test');
    assert.strictEqual(cur(r2), '0.0004', 'Banker round: 0.0009*0.5=0.00045 → 0.0004 (even)');
    console.log('[PASS] 乗算バンカーズ丸め: 0.0009*0.5=0.00045 → 0.0004');
}

// 比較演算子: Currency 同士の厳密比較
{
    const code = `
    Function Test()
        Dim a As Currency, b As Currency
        a = CCur(0.1) + CCur(0.2)
        b = CCur(0.3)
        If a = b Then
            Test = 1
        Else
            Test = 0
        End If
    End Function
    `;
    const result = runFunc(code, 'Test');
    assert.strictEqual(result, 1, 'CCur(0.1)+CCur(0.2) = CCur(0.3) (strict BigInt comparison)');
    console.log('[PASS] CCur(0.1)+CCur(0.2) = CCur(0.3) (厳密比較)');
}

// 単項マイナス
{
    const code = `
    Function Test()
        Dim c As Currency
        c = 1.5
        Test = -c
    End Function
    `;
    const result = runFunc(code, 'Test');
    assert.strictEqual(cur(result), '-1.5', 'Unary minus on Currency');
    console.log('[PASS] 単項マイナス: -CCur(1.5) = "-1.5"');
}

// Dim c As Currency の初期値は 0
{
    const code = `
    Function Test()
        Dim c As Currency
        Test = c
    End Function
    `;
    const result = runFunc(code, 'Test');
    assert.strictEqual(cur(result), '0', 'Dim c As Currency initializes to 0');
    console.log('[PASS] Dim c As Currency の初期値 = "0"');
}

// CCur(string) のパース: float を介さず文字列から直接構築
{
    const code = `
    Sub Test()
        Debug.Print CCur("0.1")
        Debug.Print CCur("922337203685477.5807")
    End Sub
    `;
    let output = '';
    evalVBASingle(code, { onPrint: s => output += s + '\n' }).callProcedure('Test', []);
    const lines = output.trim().split('\n');
    assert.strictEqual(lines[0], '0.1', 'CCur("0.1") from string');
    assert.strictEqual(lines[1], '922337203685477.5807', 'CCur max value from string');
    console.log('[PASS] CCur(string) 文字列直接パース');
}

// オーバーフロー: 加算で範囲超過 → Error 6
{
    const code = `
    Sub Test()
        On Error Resume Next
        Dim a As Currency
        a = CCur("922337203685477.5807")
        a = a + CCur(0.0001)
        If Err.Number <> 0 Then Debug.Print "Error:" & Err.Number
    End Sub
    `;
    let output = '';
    evalVBASingle(code, { onPrint: s => output += s + '\n' }).callProcedure('Test', []);
    assert.strictEqual(output.trim(), 'Error:6', 'Currency + overflow → Error 6');
    console.log('[PASS] オーバーフロー: Currency 加算超過 → Error 6');
}

// 末尾ゼロの除去: CStr(CCur(1.5)) = "1.5" (not "1.5000")
{
    const code = `
    Sub Test()
        Debug.Print CStr(CCur(1.5))
        Debug.Print CStr(CCur(1.0))
        Debug.Print CStr(CCur(1.1230))
    End Sub
    `;
    let output = '';
    evalVBASingle(code, { onPrint: s => output += s + '\n' }).callProcedure('Test', []);
    const lines = output.trim().split('\n');
    assert.strictEqual(lines[0], '1.5', 'CStr(CCur(1.5)) = "1.5" (no trailing zeros)');
    assert.strictEqual(lines[1], '1', 'CStr(CCur(1.0)) = "1"');
    assert.strictEqual(lines[2], '1.123', 'CStr(CCur(1.1230)) = "1.123"');
    console.log('[PASS] 末尾ゼロ除去: CStr(CCur) は必要最小桁のみ');
}

// @ リテラルサフィックス → VbaCurrency
{
    const code = `
    Function Test()
        Test = TypeName(1.5@)
    End Function
    `;
    const result = runFunc(code, 'Test');
    assert.strictEqual(result, 'Currency', 'TypeName(1.5@) = "Currency"');
    console.log('[PASS] @ リテラルサフィックス → Currency');
}

// Currency + Integer → Currency (型保持)
{
    const code = `
    Function Test()
        Dim c As Currency
        c = 1.5
        Test = TypeName(c + 2)
    End Function
    `;
    const result = runFunc(code, 'Test');
    assert.strictEqual(result, 'Currency', 'Currency + Integer → Currency');
    console.log('[PASS] Currency + Integer → Currency (型保持)');
}

// Currency + Double (non-integer) → Double
{
    const code = `
    Function Test()
        Dim c As Currency
        c = 1.5
        Test = TypeName(c + 0.5)
    End Function
    `;
    const result = runFunc(code, 'Test');
    assert.strictEqual(result, 'Double', 'Currency + Double → Double');
    console.log('[PASS] Currency + Double (非整数) → Double');
}

// CDbl(Currency) → Double
{
    const code = `
    Function Test()
        Dim c As Currency
        c = 1.5
        Test = CDbl(c)
    End Function
    `;
    const result = runFunc(code, 'Test');
    assert.strictEqual(result, 1.5, 'CDbl(Currency) = 1.5 (Double)');
    console.log('[PASS] CDbl(Currency) → Double');
}

// CLng(Currency) → Long (rounded)
{
    const code = `
    Function Test()
        Test = CLng(CCur(1.6))
    End Function
    `;
    const result = runFunc(code, 'Test');
    assert.strictEqual(result, 2, 'CLng(CCur(1.6)) = 2');
    console.log('[PASS] CLng(Currency) → Long');
}

console.log('\n✅ Currency BigInt: 全テスト通過');
