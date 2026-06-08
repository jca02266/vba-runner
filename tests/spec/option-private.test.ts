import { evalVBASingle, assert } from '../../test-libs/test-runner';

function evalVBA(code: string): any {
    return evalVBASingle(code);
}

console.log("--- Starting Option Private Module Tests ---");

const code = `
    Option Private Module
    
    Public Function Hello()
        Hello = "World"
    End Function
`;

const ev = evalVBA(code);
assert.strictEqual(ev.callProcedure('Hello', []), "World", 'Option Private Module allowed and code executes');

console.log("✅ Option Private Module: All tests passed!");
