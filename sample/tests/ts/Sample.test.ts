import { VBARunner, assert } from '../../../test-libs/test-runner';

async function main() {
    console.log("--- Starting Sample VBA Tests ---\n");

    const vbaTest = new VBARunner('sample/src/vba/Sample.vba');

    // Test Add
    console.log("[Test] Add");
    assert.strictEqual(vbaTest.run('Add', [10, 20]), 30, "Add(10, 20) -> 30");
    assert.strictEqual(vbaTest.run('Add', [-5, 5]), 0, "Add(-5, 5) -> 0");

    // Test Multiply
    console.log("\n[Test] Multiply");
    assert.strictEqual(vbaTest.run('Multiply', [3, 4]), 12, "Multiply(3, 4) -> 12");
    assert.strictEqual(vbaTest.run('Multiply', [10, 0]), 0, "Multiply(10, 0) -> 0");

    // Test Eval
    console.log("\n[Test] Eval");
    assert.strictEqual(vbaTest.eval("Add(2, 3)"), 5, "eval Add(2, 3)");
    assert.strictEqual(vbaTest.eval("Multiply(2, 3)"), 6, "eval Multiply(2, 3)");
    assert.strictEqual(vbaTest.eval("1 + 2 * 3"), 7, "eval 1 + 2 * 3");

    console.log("\n--- All sample tests passed! ---");
}

main().catch(err => {
    console.error(err);
    process.exit(1);
});
