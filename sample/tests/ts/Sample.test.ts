import { VBARunner, assert } from '../../../test-libs/test-runner';

async function main() {
    console.log("--- Starting Sample VBA Tests ---\n");

    const vbaRunner = new VBARunner('sample/src/vba/Sample.vba');

    // Test Add
    console.log("[Test] Add");
    assert.strictEqual(vbaRunner.run('Add', [10, 20]), 30, "Add(10, 20) -> 30");
    assert.strictEqual(vbaRunner.run('Add', [-5, 5]), 0, "Add(-5, 5) -> 0");

    // Test Multiply
    console.log("\n[Test] Multiply");
    assert.strictEqual(vbaRunner.run('Multiply', [3, 4]), 12, "Multiply(3, 4) -> 12");
    assert.strictEqual(vbaRunner.run('Multiply', [10, 0]), 0, "Multiply(10, 0) -> 0");

    // Test Eval
    console.log("\n[Test] Eval");
    assert.strictEqual(vbaRunner.eval("Add(2, 3)"), 5, "eval Add(2, 3)");
    assert.strictEqual(vbaRunner.eval("Multiply(2, 3)"), 6, "eval Multiply(2, 3)");
    assert.strictEqual(vbaRunner.eval("1 + 2 * 3"), 7, "eval 1 + 2 * 3");

    console.log("\n--- All sample tests passed! ---");
}

main().catch(err => {
    console.error(err);
    process.exit(1);
});
