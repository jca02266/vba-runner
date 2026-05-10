import { VBATest, assert } from '../ts/test-runner';

async function main() {
    console.log("--- Starting Dictionary Bang (!) Access Tests ---");
    const vbaTest = new VBATest('sample/src/vba/Sample.vba');

    // Create a mock Dictionary
    const mockDict = {
        __isVbaDict__: true,
        __map__: new Map<string, any>()
    };
    mockDict.__map__.set("Price", 100);
    mockDict.__map__.set("Name", "Apple");

    vbaTest.set("myDict", mockDict);

    console.log("[Test] Reading from Dictionary with !");
    assert.strictEqual(vbaTest.eval("myDict!Price"), 100);
    assert.strictEqual(vbaTest.eval("myDict!Name"), "Apple");

    console.log("[Test] Nested access with ! and .");
    // If dictionary contains an object
    mockDict.__map__.set("SubObj", { value: 42 });
    assert.strictEqual(vbaTest.eval("myDict!SubObj.value"), 42);

    console.log("--- All Dictionary Bang (!) Access Tests passed! ---");
}

main().catch(e => {
    console.error(e);
    process.exit(1);
});
