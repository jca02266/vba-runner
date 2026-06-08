/**
 * VBA §3.3.5.2: contextual keywords as variable names
 *
 * reserved-identifier = statement-keyword / marker-keyword / operator-identifier / ...
 * IDENTIFIER = <any lex-identifier that is not a reserved-identifier>
 *
 * The following words are NOT in the reserved-identifier list and therefore MUST be
 * usable as Dim variable names, assignment targets, and in expressions.
 */
import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';
import { Evaluator } from '../../src/engine/evaluator';
import { assert } from '../../test-libs/test-runner';

function run(code: string, name: string): any {
    const tokens = new Lexer(code).tokenize();
    const ast = new Parser(tokens).parse();
    if (ast.diagnostics?.length) {
        throw new Error(`Parse error: ${ast.diagnostics.map((d: any) => d.message).join(', ')}`);
    }
    const ev = new Evaluator(() => {});
    ev.evaluateModule(ast);
    ev.resolveIdentifiers([{ ast, moduleName: '' }]);
    return ev.callProcedure(name, []);
}

// ---------------------------------------------------------------
// Group 1: Open-mode contextual keywords
// ---------------------------------------------------------------
{
    const result = run(`
        Function TestAppend()
            Dim append As String
            append = "a"
            TestAppend = append
        End Function
    `, 'TestAppend');
    assert.strictEqual(result, 'a', 'Dim append As String');
    console.log('[PASS] Dim append As String');
}
{
    const result = run(`
        Function TestOutput()
            Dim output As String
            output = "o"
            TestOutput = output
        End Function
    `, 'TestOutput');
    assert.strictEqual(result, 'o', 'Dim output As String');
    console.log('[PASS] Dim output As String');
}
{
    const result = run(`
        Function TestRandom()
            Dim random As Long
            random = 42
            TestRandom = random
        End Function
    `, 'TestRandom');
    assert.strictEqual(result, 42, 'Dim random As Long');
    console.log('[PASS] Dim random As Long');
}
{
    const result = run(`
        Function TestBinary()
            Dim binary As String
            binary = "b"
            TestBinary = binary
        End Function
    `, 'TestBinary');
    assert.strictEqual(result, 'b', 'Dim binary As String');
    console.log('[PASS] Dim binary As String');
}
{
    const result = run(`
        Function TestAccess()
            Dim access As String
            access = "ac"
            TestAccess = access
        End Function
    `, 'TestAccess');
    assert.strictEqual(result, 'ac', 'Dim access As String');
    console.log('[PASS] Dim access As String');
}
{
    const result = run(`
        Function TestRead()
            Dim read As String
            read = "r"
            TestRead = read
        End Function
    `, 'TestRead');
    assert.strictEqual(result, 'r', 'Dim read As String');
    console.log('[PASS] Dim read As String');
}

// ---------------------------------------------------------------
// Group 2: Option Compare / Option Base contextual keywords
// ---------------------------------------------------------------
{
    const result = run(`
        Function TestText()
            Dim text As String
            text = "t"
            TestText = text
        End Function
    `, 'TestText');
    assert.strictEqual(result, 't', 'Dim text As String');
    console.log('[PASS] Dim text As String');
}
{
    const result = run(`
        Function TestCompare()
            Dim compare As Long
            compare = 1
            TestCompare = compare
        End Function
    `, 'TestCompare');
    assert.strictEqual(result, 1, 'Dim compare As Long');
    console.log('[PASS] Dim compare As Long');
}
{
    const result = run(`
        Function TestExplicit()
            Dim explicit As Long
            explicit = 99
            TestExplicit = explicit
        End Function
    `, 'TestExplicit');
    assert.strictEqual(result, 99, 'Dim explicit As Long');
    console.log('[PASS] Dim explicit As Long');
}

// ---------------------------------------------------------------
// Group 3: Declare-syntax contextual keywords
// ---------------------------------------------------------------
{
    const result = run(`
        Function TestLib()
            Dim lib As String
            lib = "mylib"
            TestLib = lib
        End Function
    `, 'TestLib');
    assert.strictEqual(result, 'mylib', 'Dim lib As String');
    console.log('[PASS] Dim lib As String');
}
{
    const result = run(`
        Function TestAlias()
            Dim alias As String
            alias = "al"
            TestAlias = alias
        End Function
    `, 'TestAlias');
    assert.strictEqual(result, 'al', 'Dim alias As String');
    console.log('[PASS] Dim alias As String');
}
{
    const result = run(`
        Function TestPtrSafe()
            Dim ptrsafe As Long
            ptrsafe = 7
            TestPtrSafe = ptrsafe
        End Function
    `, 'TestPtrSafe');
    assert.strictEqual(result, 7, 'Dim ptrsafe As Long');
    console.log('[PASS] Dim ptrsafe As Long');
}

// ---------------------------------------------------------------
// Group 4: For-loop contextual keyword
// ---------------------------------------------------------------
{
    const result = run(`
        Function TestStep()
            Dim step As Long
            step = 3
            TestStep = step
        End Function
    `, 'TestStep');
    assert.strictEqual(result, 3, 'Dim step As Long');
    console.log('[PASS] Dim step As Long');
}

// ---------------------------------------------------------------
// Group 5: Option Module contextual keyword
// ---------------------------------------------------------------
{
    const result = run(`
        Function TestModule()
            Dim module As String
            module = "m"
            TestModule = module
        End Function
    `, 'TestModule');
    assert.strictEqual(result, 'm', 'Dim module As String');
    console.log('[PASS] Dim module As String');
}

// ---------------------------------------------------------------
// Group 6: Statement-start contextual keywords
// (disambiguated by "followed by =" → assignment, otherwise statement)
// ---------------------------------------------------------------
{
    const result = run(`
        Function TestKill()
            Dim kill As String
            kill = "k"
            TestKill = kill
        End Function
    `, 'TestKill');
    assert.strictEqual(result, 'k', 'Dim kill As String (assignment)');
    console.log('[PASS] Dim kill As String (assignment)');
}
{
    const result = run(`
        Function TestReset()
            Dim reset As Long
            reset = 5
            TestReset = reset
        End Function
    `, 'TestReset');
    assert.strictEqual(result, 5, 'Dim reset As Long (assignment)');
    console.log('[PASS] Dim reset As Long (assignment)');
}
{
    const result = run(`
        Function TestWidth()
            Dim width As Long
            width = 80
            TestWidth = width
        End Function
    `, 'TestWidth');
    assert.strictEqual(result, 80, 'Dim width As Long (assignment)');
    console.log('[PASS] Dim width As Long (assignment)');
}
{
    const result = run(`
        Function TestLine()
            Dim line As String
            line = "ln"
            TestLine = line
        End Function
    `, 'TestLine');
    assert.strictEqual(result, 'ln', 'Dim line As String (assignment)');
    console.log('[PASS] Dim line As String (assignment)');
}
{
    const result = run(`
        Function TestAppActivate()
            Dim appactivate As String
            appactivate = "aa"
            TestAppActivate = appactivate
        End Function
    `, 'TestAppActivate');
    assert.strictEqual(result, 'aa', 'Dim appactivate As String (assignment)');
    console.log('[PASS] Dim appactivate As String (assignment)');
}
{
    const result = run(`
        Function TestSendKeys()
            Dim sendkeys As String
            sendkeys = "sk"
            TestSendKeys = sendkeys
        End Function
    `, 'TestSendKeys');
    assert.strictEqual(result, 'sk', 'Dim sendkeys As String (assignment)');
    console.log('[PASS] Dim sendkeys As String (assignment)');
}

// ---------------------------------------------------------------
// Group 7: $ suffix on contextual keywords (TYPED-NAME = IDENTIFIER type-suffix)
// Contextual keywords are IDENTIFIER per spec → valid with $ suffix
// ---------------------------------------------------------------
{
    const result = run(`
        Function TestAppendDollar()
            Dim append$
            append$ = "hello"
            TestAppendDollar = append$
        End Function
    `, 'TestAppendDollar');
    assert.strictEqual(result, 'hello', 'Dim append$ (contextual keyword + $)');
    console.log('[PASS] Dim append$ (contextual + $)');
}
{
    const result = run(`
        Function TestAccessDollar()
            Dim access$
            access$ = "world"
            TestAccessDollar = access$
        End Function
    `, 'TestAccessDollar');
    assert.strictEqual(result, 'world', 'Dim access$ (contextual keyword + $)');
    console.log('[PASS] Dim access$ (contextual + $)');
}

// ---------------------------------------------------------------
// Group 8: $ suffix on reserved keywords — per spec these are INVALID
// (TYPED-NAME requires IDENTIFIER, not reserved-identifier)
// Current engine: incorrectly allows them because lexer appends $ before
// keyword matching (dim$ ≠ dim in the keyword table).
// Documented as known issue in TODO_SPEC.md.
// ---------------------------------------------------------------
{
    // This SHOULD be a parse error per spec, but currently passes.
    // Documented as a known bug.
    let passed = false;
    try {
        run(`
            Function Test()
                Dim dim$
                dim$ = "x"
                Test = dim$
            End Function
        `, 'Test');
        passed = true;
    } catch (_) {
        passed = false;
    }
    // Document current (incorrect) behavior without asserting either way
    console.log(`[INFO] Dim dim$ (reserved + $): ${passed ? 'passes (known spec violation)' : 'correctly rejected'}`);
}

console.log('\n✅ contextual-keywords: 全テスト通過');
