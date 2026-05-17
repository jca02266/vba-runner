import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { VBARunner, assert } from '../../test-libs/test-runner';

function runnerFromSource(source: string): VBARunner {
    const tmpFile = path.join(os.tmpdir(), `vba_test_${Date.now()}.bas`);
    fs.writeFileSync(tmpFile, source, 'utf-8');
    try {
        return new VBARunner(tmpFile);
    } finally {
        fs.unlinkSync(tmpFile);
    }
}

console.log('--- getTypeDefinitions() Tests ---');

// 1. Basic: single Type with multiple fields
{
    const runner = runnerFromSource(`
        Type Point
            X As Long
            Y As Long
        End Type
    `);
    const types = runner.getTypeDefinitions();
    assert.strictEqual(Object.keys(types).length, 1, '1 type declared');
    assert.strictEqual(types['Point'] !== undefined, true, 'Point type exists');
    assert.strictEqual(types['Point']['X'], 'number', 'X → number');
    assert.strictEqual(types['Point']['Y'], 'number', 'Y → number');
    console.log('[PASS] Basic single Type');
}

// 2. Multiple Type declarations
{
    const runner = runnerFromSource(`
        Type Rect
            LeftEdge  As Long
            TopEdge   As Long
            RectW     As Long
            RectH     As Long
        End Type
        Type Color
            R As Byte
            G As Byte
            B As Byte
        End Type
    `);
    const types = runner.getTypeDefinitions();
    assert.strictEqual(Object.keys(types).length, 2, '2 types declared');
    assert.strictEqual(types['Rect']['RectW'], 'number', 'Rect.RectW (Long) → number');
    assert.strictEqual(types['Color']['R'], 'number', 'Color.R (Byte) → number');
    console.log('[PASS] Multiple Type declarations');
}

// 3. Type mapping: all numeric VBA types → number
{
    const runner = runnerFromSource(`
        Type NumericTypes
            A As Integer
            B As Long
            C As Single
            D As Double
            E As Currency
            F As Byte
            G As Date
        End Type
    `);
    const fields = runner.getTypeDefinitions()['NumericTypes'];
    for (const [name, tsType] of Object.entries(fields)) {
        assert.strictEqual(tsType, 'number', `${name} → number`);
    }
    console.log('[PASS] All numeric VBA types map to number');
}

// 4. String and Boolean mapping
{
    const runner = runnerFromSource(`
        Type StrBool
            Name    As String
            Active  As Boolean
        End Type
    `);
    const fields = runner.getTypeDefinitions()['StrBool'];
    assert.strictEqual(fields['Name'], 'string', 'String → string');
    assert.strictEqual(fields['Active'], 'boolean', 'Boolean → boolean');
    console.log('[PASS] String → string, Boolean → boolean');
}

// 5. Variant and Object mapping
{
    const runner = runnerFromSource(`
        Type Mixed
            Val  As Variant
            Obj  As Object
            Cust As SomeCustomType
        End Type
    `);
    const fields = runner.getTypeDefinitions()['Mixed'];
    assert.strictEqual(fields['Val'], 'any', 'Variant → any');
    assert.strictEqual(fields['Obj'], 'object', 'Object → object');
    assert.strictEqual(fields['Cust'], 'any', 'unknown type → any');
    console.log('[PASS] Variant → any, Object → object, unknown → any');
}

// 6. No Type declarations → empty result
{
    const runner = runnerFromSource(`
        Function Add(a As Long, b As Long) As Long
            Add = a + b
        End Function
    `);
    const types = runner.getTypeDefinitions();
    assert.strictEqual(Object.keys(types).length, 0, 'no Type → empty object');
    console.log('[PASS] No Type declarations → empty result');
}

// 7. Case-insensitive VBA type names
{
    const runner = runnerFromSource(`
        Type LowerCase
            A As string
            B As boolean
            C As long
        End Type
    `);
    const fields = runner.getTypeDefinitions()['LowerCase'];
    assert.strictEqual(fields['A'], 'string', 'lowercase string → string');
    assert.strictEqual(fields['B'], 'boolean', 'lowercase boolean → boolean');
    assert.strictEqual(fields['C'], 'number', 'lowercase long → number');
    console.log('[PASS] Case-insensitive VBA type names');
}

// 8. Type co-located with functions (realistic usage)
{
    const runner = runnerFromSource(`
        Type InventoryParams
            CurrentStock  As Long
            SoldUnits     As Long
            RestockAmount As Long
            MinStock      As Long
            MaxStock      As Long
        End Type

        Function CalculateNewInventory(p As InventoryParams) As Long
            Dim newStock As Long
            newStock = p.CurrentStock - p.SoldUnits + p.RestockAmount
            If newStock < p.MinStock Then newStock = p.MinStock
            If newStock > p.MaxStock Then newStock = p.MaxStock
            CalculateNewInventory = newStock
        End Function
    `);
    const types = runner.getTypeDefinitions();
    assert.strictEqual(Object.keys(types).length, 1, '1 type in mixed source');
    const fields = types['InventoryParams'];
    assert.strictEqual(fields['CurrentStock'], 'number', 'CurrentStock → number');
    assert.strictEqual(fields['SoldUnits'], 'number', 'SoldUnits → number');
    assert.strictEqual(fields['MinStock'], 'number', 'MinStock → number');
    assert.strictEqual(fields['MaxStock'], 'number', 'MaxStock → number');
    console.log('[PASS] Type co-located with functions');
}

console.log('\n✅ getTypeDefinitions: 全テスト通過');
