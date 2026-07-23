import { evalVBASingle, assert } from '../../test-libs/test-runner';

const ev = evalVBASingle(`
    Function Greeting$()
        Greeting$ = "hello"
    End Function

    Function CallGreeting$()
        CallGreeting$ = Greeting()
    End Function
`);

assert.strictEqual(ev.callProcedure('CallGreeting', []), 'hello');

console.log('✅ Typed procedure declarations use their unsuffixed call name');
