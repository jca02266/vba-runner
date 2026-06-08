import { evalVBASingle, assert } from '../../test-libs/test-runner';

console.log('[Test Suite] MsgBox / InputBox / Attributes の検証');

const code = `
    Attribute VB_Name = "Module1"
    Attribute VB_Description = "Test module"

    Function TestMsgBox()
        TestMsgBox = MsgBox("Hello", vbYesNo, "Title")
    End Function

    Function TestInputBox()
        TestInputBox = InputBox("Prompt", "Title", "DefaultVal")
    End Function

    Function TestConstants()
        TestConstants = (vbOK = 1 And vbYes = 6)
    End Function
`;

const ev = evalVBASingle(code);

assert.strictEqual(ev.callProcedure('TestMsgBox', []), 1, 'MsgBox should return vbOK(1) by default');
assert.strictEqual(ev.callProcedure('TestInputBox', []), 'DefaultVal', 'InputBox should return default value');
assert.strictEqual(ev.callProcedure('TestConstants', []).value, -1, 'VBA Constants (vbOK, vbYes) should be defined');

console.log('✅ MsgBox / InputBox / Attributes: 全テスト通過');
