import { evalVBASingle, assert } from '../../test-libs/test-runner';

console.log("Running Registry tests...");

const vbaCode = `
    SaveSetting "MyApp", "Section1", "Key1", "Value1"
    Debug.Print GetSetting("MyApp", "Section1", "Key1")
    SaveSetting "MyApp", "Section1", "Key1", "NewValue"
    Debug.Print GetSetting("MyApp", "Section1", "Key1")
    Debug.Print GetSetting("MyApp", "Section1", "Key2", "Default")
    DeleteSetting "MyApp", "Section1", "Key1"
    Debug.Print GetSetting("MyApp", "Section1", "Key1", "Deleted")
`;

const lines: string[] = [];
evalVBASingle(vbaCode, { onPrint: (o) => lines.push(o.trim()) });

const expected = ["Value1", "NewValue", "Default", "Deleted"];
expected.forEach((exp, i) => {
    assert.strictEqual(lines[i], exp, `Registry line ${i + 1}`);
});

console.log('✅ Registry (Settings): 全テスト通過');

const matrixLines: string[] = [];
evalVBASingle(`
    SaveSetting "MatrixApp", "Startup", "Top", 75
    SaveSetting "MatrixApp", "Startup", "Left", 50
    Dim settings As Variant
    settings = GetAllSettings("MatrixApp", "Startup")
    Debug.Print UBound(settings, 1) & ":" & UBound(settings, 2)
    Debug.Print settings(0, 0) & "=" & settings(0, 1)
    Debug.Print settings(1, 0) & "=" & settings(1, 1)
`, { onPrint: (o) => matrixLines.push(o.trim()) });
assert.deepStrictEqual(matrixLines, ['1:1', 'Top=75', 'Left=50'], 'GetAllSettings returns a 2D array');

const emptySectionLines: string[] = [];
evalVBASingle(`
    SaveSetting "EmptyMatrixApp", "EmptySection", "Only", "value"
    DeleteSetting "EmptyMatrixApp", "EmptySection", "Only"
    Dim emptySettings As Variant
    emptySettings = GetAllSettings("EmptyMatrixApp", "EmptySection")
    Debug.Print TypeName(emptySettings)
    Debug.Print LBound(emptySettings, 1) & ":" & UBound(emptySettings, 1)
    Debug.Print LBound(emptySettings, 2) & ":" & UBound(emptySettings, 2)
`, { onPrint: (o) => emptySectionLines.push(o.trim()) });
assert.deepStrictEqual(
    emptySectionLines,
    ['Variant()', '0:-1', '0:1'],
    'GetAllSettings preserves empty two-dimensional section bounds',
);

assert.throwsMatch(
    () => evalVBASingle('SaveSetting Null, "S", "K", "V"'),
    /Invalid use of Null/,
    'SaveSetting required String arguments reject Null',
);
assert.throwsMatch(
    () => evalVBASingle('AppActivate Null'),
    /Invalid use of Null/,
    'AppActivate required String arguments reject Null',
);
assert.throwsMatch(
    () => evalVBASingle('SendKeys Null'),
    /Invalid use of Null/,
    'SendKeys required String arguments reject Null',
);
assert.throwsMatch(
    () => evalVBASingle('AppActivate "x", Wait:=Null'),
    /Invalid use of Null/,
    'AppActivate statement named Wait rejects Null',
);
assert.throwsMatch(
    () => evalVBASingle('SendKeys "x", Wait:=Null'),
    /Invalid use of Null/,
    'SendKeys statement named Wait rejects Null',
);
assert.throwsMatch(
    () => evalVBASingle('AppActivate "x", Null'),
    /Invalid use of Null/,
    'AppActivate statement positional Wait rejects Null',
);
assert.throwsMatch(
    () => evalVBASingle('SendKeys "x", Null'),
    /Invalid use of Null/,
    'SendKeys statement positional Wait rejects Null',
);

const interactionWaitEvaluation = evalVBASingle(String.raw`
    Dim waitEvaluations As Long
    Function NextWait() As Boolean
        waitEvaluations = waitEvaluations + 1
        NextWait = True
    End Function
    Function ProbeInteractionWaitEvaluation() As String
        AppActivate "x", Wait:=NextWait()
        SendKeys "x", Wait:=NextWait()
        ProbeInteractionWaitEvaluation = CStr(waitEvaluations)
    End Function
`);
assert.strictEqual(
    interactionWaitEvaluation.callProcedure('ProbeInteractionWaitEvaluation', []),
    '2',
    'AppActivate and SendKeys evaluate named Wait expressions once',
);

console.log('✅ Registry/Interaction optional and array boundaries: 全テスト通過');
