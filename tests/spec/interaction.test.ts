import { evalVBASingle, assert } from '../../test-libs/test-runner';

let output = "";
const onPrint = (msg: string) => {
    output += msg + "\n";
};

function evalVBA(code: string): any {
    const ev = evalVBASingle(code, { onPrint: onPrint });
    return ev;
}

console.log("--- Starting Interaction Module Stub Tests ---");

// 1. MsgBox
{
    const code = `
        Function TestMsg()
            TestMsg = MsgBox("Hello", 64, "MyTitle")
        End Function
    `;
    output = "";
    const ev = evalVBA(code);
    const res = ev.callProcedure('TestMsg', []);
    assert.strictEqual(res, 1, 'MsgBox returns 1 (vbOK) as default stub');
    assert.ok(output.includes("[MSGBOX] MyTitle: Hello"), 'MsgBox output captured');
}

// 2. InputBox
{
    const code = `
        Function TestInput()
            TestInput = InputBox("Enter name", "User Prompt", "Guest")
        End Function
    `;
    output = "";
    const ev = evalVBA(code);
    const res = ev.callProcedure('TestInput', []);
    assert.strictEqual(res, "Guest", 'InputBox returns default value as stub');
    assert.ok(output.includes("[INPUTBOX] Enter name"), 'InputBox output captured');
}

console.log("✅ Interaction Module: All tests passed!");
