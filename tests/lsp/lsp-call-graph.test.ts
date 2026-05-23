import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';
import { CallGraphProvider } from '../../src/lsp/call-graph-provider';
import { assert } from '../../test-libs/test-runner';

const provider = new CallGraphProvider();

function parseVBA(src: string) {
    const tokens = new Lexer(src).tokenize();
    const ast = new Parser(tokens).parse();
    return ast.body;
}

function getCallGraph(src1: string, src2?: string) {
    const fileMap = new Map();
    fileMap.set('file1.bas', { statements: parseVBA(src1), uri: 'file:///file1.bas' });
    if (src2) {
        fileMap.set('file2.bas', { statements: parseVBA(src2), uri: 'file:///file2.bas' });
    }
    return provider.buildCallGraph(fileMap);
}

// Test 1: Single file, two procs, one calls the other
{
    const src = `
        Sub Main()
            Call CalcTotal(1, 2)
        End Sub

        Function CalcTotal(a, b)
            CalcTotal = a + b
        End Function
    `;
    const graph = getCallGraph(src);
    assert.strictEqual(graph.nodes.size, 2);
    assert.strictEqual(graph.edges.length, 1);
    assert.strictEqual(graph.edges[0].from, 'main');
    assert.strictEqual(graph.edges[0].to, 'calctotal');
    const callTotalNode = graph.nodes.get('calctotal');
    assert.strictEqual(callTotalNode?.referenceCount, 1);
    console.log('[PASS] Test 1: Single file, two procs, one calls the other');
}

// Test 2: Entry point detection (Public + refs=0)
{
    const src = `
        Public Sub AutoScheduleTasks()
            MsgBox "Hello"
        End Sub
    `;
    const graph = getCallGraph(src);
    const proc = graph.nodes.get('autoscheduletasks');
    assert.ok(proc);
    assert.strictEqual(proc.scope, 'public');
    assert.strictEqual(proc.referenceCount, 0);
    console.log('[PASS] Test 2: Entry point detection (Public + refs=0)');
}

// Test 3: Dead code detection (Private + refs=0)
{
    const src = `
        Private Sub UnusedHelper()
            Debug.Print "unused"
        End Sub
    `;
    const graph = getCallGraph(src);
    const proc = graph.nodes.get('unusedhelper');
    assert.ok(proc);
    assert.strictEqual(proc.scope, 'private');
    assert.strictEqual(proc.referenceCount, 0);
    console.log('[PASS] Test 3: Dead code detection (Private + refs=0)');
}

// Test 4: Excel-dependent detection
{
    const src = `
        Sub ReadData()
            Dim val
            val = Sheets(1).Range("A1").Value
        End Sub
    `;
    const graph = getCallGraph(src);
    const proc = graph.nodes.get('readdata');
    assert.ok(proc?.isExcelDependent);
    console.log('[PASS] Test 4: Excel-dependent detection');
}

// Test 5: Non-Excel proc
{
    const src = `
        Sub Calculate()
            Dim x, y
            x = 10
            y = 20
        End Sub
    `;
    const graph = getCallGraph(src);
    const proc = graph.nodes.get('calculate');
    assert.strictEqual(proc?.isExcelDependent, false);
    console.log('[PASS] Test 5: Non-Excel proc');
}

// Test 6: Recursive call (self-reference)
{
    const src = `
        Function Factorial(n)
            If n <= 1 Then
                Factorial = 1
            Else
                Factorial = n * Factorial(n - 1)
            End If
        End Function
    `;
    const graph = getCallGraph(src);
    const proc = graph.nodes.get('factorial');
    assert.strictEqual(proc?.referenceCount, 1);  // Self-edge counts
    console.log('[PASS] Test 6: Recursive call (self-reference)');
}

// Test 7: Multi-file, cross-file call
{
    const src1 = `
        Sub Main()
            Call Helper()
        End Sub
    `;
    const src2 = `
        Sub Helper()
            Debug.Print "Help"
        End Sub
    `;
    const graph = getCallGraph(src1, src2);
    assert.strictEqual(graph.nodes.size, 2);
    assert.strictEqual(graph.edges.length, 1);
    assert.strictEqual(graph.edges[0].from, 'main');
    assert.strictEqual(graph.edges[0].to, 'helper');
    const helperNode = graph.nodes.get('helper');
    assert.strictEqual(helperNode?.referenceCount, 1);
    assert.ok(helperNode?.uri.includes('file2'));
    console.log('[PASS] Test 7: Multi-file, cross-file call');
}

// Test 8: Built-in function (MsgBox) not in graph
{
    const src = `
        Sub Warn()
            MsgBox "Warning"
        End Sub
    `;
    const graph = getCallGraph(src);
    assert.strictEqual(graph.nodes.size, 1);
    assert.strictEqual(graph.nodes.has('msgbox'), false);
    assert.strictEqual(graph.edges.length, 0);
    console.log('[PASS] Test 8: Built-in function (MsgBox) not in graph');
}

// Test 9: UDT constructor excluded from call graph
{
    const src = `
        Type Config
            Name As String
        End Type

        Sub Setup()
            Call Config
        End Sub
    `;
    const graph = getCallGraph(src);
    // Config is a type, not a procedure, so no edge should exist
    assert.strictEqual(graph.nodes.has('config'), false);
    assert.strictEqual(graph.edges.length, 0);
    console.log('[PASS] Test 9: UDT constructor excluded from call graph');
}

// Test 10: Call Statement syntax (Call Foo)
{
    const src = `
        Sub Main()
            Call Worker
        End Sub

        Sub Worker()
            Debug.Print "Work"
        End Sub
    `;
    const graph = getCallGraph(src);
    assert.strictEqual(graph.edges.length, 1);
    assert.strictEqual(graph.edges[0].from, 'main');
    assert.strictEqual(graph.edges[0].to, 'worker');
    console.log('[PASS] Test 10: Call Statement syntax (Call Foo)');
}

// Test 11: Empty module
{
    const src = `
        ' Just a comment
    `;
    const graph = getCallGraph(src);
    assert.strictEqual(graph.nodes.size, 0);
    assert.strictEqual(graph.edges.length, 0);
    console.log('[PASS] Test 11: Empty module');
}

// Test 12: Three procedures, no calls
{
    const src = `
        Sub First()
        End Sub

        Sub Second()
        End Sub

        Sub Third()
        End Sub
    `;
    const graph = getCallGraph(src);
    assert.strictEqual(graph.nodes.size, 3);
    assert.strictEqual(graph.edges.length, 0);
    console.log('[PASS] Test 12: Three procedures, no calls');
}

console.log('\n✅ Call Graph Provider: 全テスト通過');
