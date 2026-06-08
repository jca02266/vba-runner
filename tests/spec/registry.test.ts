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
