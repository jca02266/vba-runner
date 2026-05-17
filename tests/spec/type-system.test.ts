import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';
import { Evaluator } from '../../src/engine/evaluator';
import { assert } from '../../test-libs/test-runner';
import { NodeFileSystem } from '../../src/engine/node_filesystem';

function evalVBA(code: string): Evaluator {
    const tokens = new Lexer(code).tokenize();
    const ast = new Parser(tokens).parse();
    const ev = new Evaluator(console.log, { fs: new NodeFileSystem() });
    ev.evaluate(ast);
    return ev;
}

// === TypeName with typed variables ===
{
    const ev = evalVBA(`
        Dim xi As Integer
        Dim xl As Long
        Dim xs As Single
        Dim xd As Double
        Dim xb As Byte
        Dim xc As Currency
        Dim xstr As String
        Dim xbool As Boolean
    `);
    assert.strictEqual(ev.evalExpression('TypeName(xi)'), 'Integer');
    assert.strictEqual(ev.evalExpression('TypeName(xl)'), 'Long');
    assert.strictEqual(ev.evalExpression('TypeName(xs)'), 'Single');
    assert.strictEqual(ev.evalExpression('TypeName(xd)'), 'Double');
    assert.strictEqual(ev.evalExpression('TypeName(xb)'), 'Byte');
    assert.strictEqual(ev.evalExpression('TypeName(xc)'), 'Currency');
    assert.strictEqual(ev.evalExpression('TypeName(xstr)'), 'String');
    assert.strictEqual(ev.evalExpression('TypeName(xbool)'), 'Boolean');
    console.log('✅ TypeName with typed variables');
}

// === VarType with typed variables ===
{
    const ev = evalVBA(`
        Dim xi As Integer
        Dim xl As Long
        Dim xs As Single
        Dim xd As Double
        Dim xb As Byte
    `);
    assert.strictEqual(ev.evalExpression('VarType(xi)'), 2);
    assert.strictEqual(ev.evalExpression('VarType(xl)'), 3);
    assert.strictEqual(ev.evalExpression('VarType(xs)'), 4);
    assert.strictEqual(ev.evalExpression('VarType(xd)'), 5);
    assert.strictEqual(ev.evalExpression('VarType(xb)'), 17);
    console.log('✅ VarType with typed variables');
}

// === TypeName with literals (type inference) ===
{
    const ev = evalVBA('');
    assert.strictEqual(ev.evalExpression('TypeName(10)'), 'Integer');
    assert.strictEqual(ev.evalExpression('TypeName(40000)'), 'Long');
    assert.strictEqual(ev.evalExpression('TypeName(10.5)'), 'Double');
    assert.strictEqual(ev.evalExpression('TypeName("abc")'), 'String');
    assert.strictEqual(ev.evalExpression('TypeName(True)'), 'Boolean');
    console.log('✅ TypeName with literals (type inference)');
}

// === VarType with literals ===
{
    const ev = evalVBA('');
    assert.strictEqual(ev.evalExpression('VarType(10)'), 2);    // vbInteger
    assert.strictEqual(ev.evalExpression('VarType(40000)'), 3);  // vbLong
    assert.strictEqual(ev.evalExpression('VarType(10.5)'), 5);   // vbDouble
    console.log('✅ VarType with literals (type inference)');
}

// === Overflow checks ===
{
    const ev = evalVBA('');

    // Integer overflow
    let caught = false;
    try {
        ev.evalExpression('Dim xo As Integer\nxo = 40000');
    } catch (e: any) {
        if (e.number === 6) caught = true;
    }
    assert.strictEqual(caught, true);
    console.log('✅ Integer overflow detected');

    // Byte overflow
    caught = false;
    try {
        ev.evalExpression('Dim bo As Byte\nbo = 256');
    } catch (e: any) {
        if (e.number === 6) caught = true;
    }
    assert.strictEqual(caught, true);
    console.log('✅ Byte overflow detected');

    // Byte negative overflow
    caught = false;
    try {
        ev.evalExpression('Dim bn As Byte\nbn = -1');
    } catch (e: any) {
        if (e.number === 6) caught = true;
    }
    assert.strictEqual(caught, true);
    console.log('✅ Byte negative overflow detected');

    // Long overflow
    caught = false;
    try {
        ev.evalExpression('Dim lo As Long\nlo = 3000000000');
    } catch (e: any) {
        if (e.number === 6) caught = true;
    }
    assert.strictEqual(caught, true);
    console.log('✅ Long overflow detected');
}

// === Valid assignments within range ===
{
    const ev = evalVBA(`
        Dim xi As Integer
        xi = 100
    `);
    assert.strictEqual(ev.evalExpression('xi'), 100);
    console.log('✅ Valid Integer assignment');
}

// === Banker's rounding on Integer assignment ===
{
    const ev = evalVBA(`
        Dim xi As Integer
        xi = 10.5
    `);
    // Banker's rounding: 10.5 rounds to 10 (even)
    // Actually: 10.5 -> nearest even -> 10
    const val = ev.evalExpression('xi');
    assert.strictEqual(val, 10);
    console.log('✅ Banker\'s rounding on Integer assignment (10.5 -> 10)');
}

{
    const ev = evalVBA(`
        Dim xi As Integer
        xi = 11.5
    `);
    // 11.5 -> nearest even -> 12
    const val = ev.evalExpression('xi');
    assert.strictEqual(val, 12);
    console.log('✅ Banker\'s rounding on Integer assignment (11.5 -> 12)');
}

// === Parameter type enforcement ===
{
    const ev = evalVBA(`
        Function AddOne(x As Integer) As Integer
            AddOne = x + 1
        End Function
    `);
    assert.strictEqual(ev.callProcedure('AddOne', [10]), 11);
    console.log('✅ Parameter with type works correctly');
}

// === Parameter overflow ===
{
    const ev = evalVBA(`
        Sub TakeInt(x As Integer)
            x = x + 1
        End Sub
    `);
    let caught = false;
    try {
        ev.callProcedure('TakeInt', [40000]);
    } catch (e: any) {
        if (e.number === 6) caught = true;
    }
    assert.strictEqual(caught, true);
    console.log('✅ Parameter overflow detected');
}

// === TypeName inside a function for parameter ===
{
    const code = `
        Function GetParamType(x As Integer)
            GetParamType = TypeName(x)
        End Function
    `;
    const ev = evalVBA(code);
    assert.strictEqual(ev.callProcedure('GetParamType', [10]), 'Integer');
    console.log('✅ TypeName for parameter inside function');
}

console.log('--- All Type System tests passed! ---');
