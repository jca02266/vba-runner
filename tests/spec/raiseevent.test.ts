/**
 * Event Declaration (§5.2.4.3) & RaiseEvent (§5.4.2.20) のテスト
 */
import { Lexer } from '../../src/engine/lexer';
import { Parser } from '../../src/engine/parser';
import { Evaluator } from '../../src/engine/evaluator';
import { assert } from '../../test-libs/test-runner';

function evalVBA(code: string): any {
    const tokens = new Lexer(code).tokenize();
    const ast = new Parser(tokens).parse();
    const ev = new Evaluator(console.log);
    ev.evaluate(ast);
    return ev;
}

// --- 1. Event 宣言と RaiseEvent の基本 ---
const eventCode = `
    Class MyEmitter
        Public Event StatusChanged(msg As String, code As Integer)

        Public Sub DoSomething()
            RaiseEvent StatusChanged("Processing", 1)
            RaiseEvent StatusChanged("Done", 0)
        End Sub
    End Class

    ' Client Module
    Dim WithEvents obj As MyEmitter
    Dim lastMsg
    Dim lastCode

    Sub obj_StatusChanged(msg, code)
        lastMsg = msg
        lastCode = code
    End Sub

    Sub Main()
        Set obj = New MyEmitter
        obj.DoSomething
    End Sub
`;

const ev = evalVBA(eventCode);
ev.callProcedure('Main', []);

assert.strictEqual(ev.env.get('lastmsg'), "Done", '最終的なイベントメッセージが "Done"');
assert.strictEqual(ev.env.get('lastcode'), 0, '最終的なイベントコードが 0');

console.log('[PASS] Event & RaiseEvent (WithEvents 経由)');

// --- 2. 引数なしイベント ---
const simpleEventCode = `
    Class SimpleEmitter
        Event Fired()
        Sub Fire()
            RaiseEvent Fired
        End Sub
    End Class

    Dim WithEvents s As SimpleEmitter
    Dim fireCount
    
    Sub s_Fired()
        fireCount = fireCount + 1
    End Sub
    
    Sub Test()
        fireCount = 0
        Set s = New SimpleEmitter
        s.Fire
        s.Fire
    End Sub
`;

const ev2 = evalVBA(simpleEventCode);
ev2.callProcedure('Test', []);
assert.strictEqual(ev2.env.get('firecount'), 2, 'イベントが2回発火');
console.log('[PASS] 引数なしイベント');

console.log('\n✅ Event & RaiseEvent: 全テスト通過');
