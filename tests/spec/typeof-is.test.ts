import { VBATest, assert } from '../ts/test-runner';

async function main() {
    console.log("--- Starting TypeOf...Is Tests ---");
    const vbaTest = new VBATest('sample/src/vba/Sample.vba');

    // Mock objects
    const mockDict = { __isVbaDict__: true };
    const mockColl = { __isVbaCollection__: true };
    const mockUserObj = { __vbaTypeName__: "MyClass" };

    vbaTest.set("myDict", mockDict);
    vbaTest.set("myColl", mockColl);
    vbaTest.set("myObj", mockUserObj);

    console.log("[Test] TypeOf ... Is Object");
    assert.strictEqual(vbaTest.eval("TypeOf myDict Is Object"), -1);
    assert.strictEqual(vbaTest.eval("TypeOf myObj Is Object"), -1);
    assert.strictEqual(vbaTest.eval("TypeOf 123 Is Object"), 0);

    console.log("[Test] TypeOf ... Is Dictionary");
    assert.strictEqual(vbaTest.eval("TypeOf myDict Is Dictionary"), -1);
    assert.strictEqual(vbaTest.eval("TypeOf myColl Is Dictionary"), 0);

    console.log("[Test] TypeOf ... Is Collection");
    assert.strictEqual(vbaTest.eval("TypeOf myColl Is Collection"), -1);

    console.log("[Test] TypeOf ... Is UserType");
    assert.strictEqual(vbaTest.eval("TypeOf myObj Is MyClass"), -1);
    assert.strictEqual(vbaTest.eval("TypeOf myObj Is OtherClass"), 0);

    console.log("--- All TypeOf...Is Tests passed! ---");
}

main().catch(e => {
    console.error(e);
    process.exit(1);
});
