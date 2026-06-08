import { evalVBASingle, assert } from '../../test-libs/test-runner';

console.log('[Test Suite] Option Base and Collection の検証');

// Test Option Base 1
{
    const code = `
        Option Base 1
        Sub TestArray()
            Dim arr(5)
            arr(1) = 10
            arr(5) = 50
            Debug.Print arr(1)
            Debug.Print arr(5)
            Debug.Print LBound(arr)
            Debug.Print UBound(arr)
        End Sub
    `;
    const output: string[] = [];
    const ev = evalVBASingle(code, { onPrint: (s) => output.push(s) });
    ev.callProcedure('TestArray', []);

    console.log('Actual output:', output);
    assert.strictEqual(output[0], '10', 'arr(1) should be 10');
    assert.strictEqual(output[1], '50', 'arr(5) should be 50');
    assert.strictEqual(output[2], '1', 'LBound should be 1');
    assert.strictEqual(output[3], '5', 'UBound should be 5');
    console.log('✅ Option Base 1: 通過');
}

// Test Collection
{
    const code = `
        Sub TestCollection()
            Dim col As New Collection
            col.Add "Item1", "Key1"
            col.Add "Item2"
            Debug.Print col.Count
            Debug.Print col.Item(1)
            Debug.Print col.Item("Key1")
            col.Remove 1
            Debug.Print col.Count
        End Sub
    `;
    const output: string[] = [];
    const ev = evalVBASingle(code, { onPrint: (s) => output.push(s) });
    ev.callProcedure('TestCollection', []);

    assert.strictEqual(output[0], '2', 'Initial count should be 2');
    assert.strictEqual(output[1], 'Item1', 'Item(1) should be Item1');
    assert.strictEqual(output[2], 'Item1', 'Item("Key1") should be Item1');
    assert.strictEqual(output[3], '1', 'Count after remove should be 1');
    console.log('✅ Collection: 通過');
}

console.log('✅ すべてのテストを通過しました');
