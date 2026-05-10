import { assert } from '../ts/test-runner';
import { Evaluator } from '../../src/compiler/evaluator';
import { Parser } from '../../src/compiler/parser';
import { Lexer } from '../../src/compiler/lexer';

function evalVBA(code: string) {
    const tokens = new Lexer(code).tokenize();
    const parser = new Parser(tokens);
    const program = parser.parse();
    const evaluator = new Evaluator((s) => console.log(s));
    evaluator.evaluate(program);
    return evaluator;
}

console.log('[Test Suite] CreateObject (Dictionary / FSO) の検証');

const code = `
    Function TestDictionary()
        Dim dict
        Set dict = CreateObject("Scripting.Dictionary")
        dict.Add "A", 100
        dict.Add "B", 200
        
        Dim res
        res = 0
        If dict.Exists("A") Then res = res + dict.Item("A")
        If dict.Exists("B") Then res = res + dict.Item("B")
        
        dict.Remove "A"
        If Not dict.Exists("A") Then res = res + 50
        
        TestDictionary = res + dict.Count
    End Function

    Function TestFSO()
        Dim fso, f
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set f = fso.CreateTextFile("test.txt")
        f.WriteLine "Hello"
        f.Close
        TestFSO = 1
    End Function
`;

const ev = evalVBA(code);

assert.strictEqual(ev.callProcedure('TestDictionary', []), 351, 'Dictionary operations (Add, Exists, Item, Remove, Count) should work (100+200+50+1 = 351)');
assert.strictEqual(ev.callProcedure('TestFSO', []), 1, 'FSO stub should work without error');

console.log('✅ CreateObject: 全テスト通過');
