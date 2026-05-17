import { VBARunner, assert } from '../../../test-libs/test-runner';

async function main() {
    console.log("--- Starting LibSheet.vba Unit Tests ---");
    const vbaTest = new VBARunner('sample/src/vba/LibSheet.vba');

    // Create mock Worksheet factory
    function createMockWorksheet(name: string): any {
        const ws: any = {
            name: name,
            cellsDeleted: false,
        };
        ws.cells = {
            delete: function () {
                ws.cellsDeleted = true;
            }
        };
        return ws;
    }

    // Create a Worksheets collection mock that supports both:
    //   wb.Worksheets(name) -> index access (CallExpression)
    //   wb.Worksheets.Add   -> method call on the collection (MemberExpression)
    // The evaluator resolves MemberExpression by looking for properties,
    // and resolves CallExpression by looking for functions/arrays.
    // So we use a function object with .add as a property.
    function createMockWorksheets(sheets: Map<string, any>, addFn: () => any): any {
        const fn = function (name: string) {
            if (sheets.has(name)) return sheets.get(name);
            throw new Error("Subscript out of range");
        };
        fn.add = addFn;
        return fn;
    }

    // Test 1: Sheet already exists in Workbook → clear cells and return it
    console.log("\n[Test Suite] NewWorksheet - existing sheet");
    {
        const existingSheet = createMockWorksheet("Sheet1");
        const sheetsMap = new Map<string, any>();
        sheetsMap.set("Sheet1", existingSheet);

        const mockWb: any = {
            worksheets: createMockWorksheets(sheetsMap, () => createMockWorksheet(""))
        };

        const result = vbaTest.run('NewWorksheet', ["Sheet1", mockWb]);

        // The function should return the existing sheet
        assert.strictEqual(result === existingSheet, true, "Should return existing sheet");
        // Cells should have been cleared
        assert.strictEqual(existingSheet.cellsDeleted, true, "Cells should be deleted on existing sheet");
    }

    // Test 2: Sheet does NOT exist → create new sheet via error handler
    console.log("\n[Test Suite] NewWorksheet - new sheet");
    {
        const newSheet = createMockWorksheet("");
        const sheetsMap = new Map<string, any>(); // Empty - no sheets exist

        const mockWb: any = {
            worksheets: createMockWorksheets(sheetsMap, () => newSheet)
        };

        const result = vbaTest.run('NewWorksheet', ["NewSheet", mockWb]);

        // The error handler should have created a new sheet and named it
        assert.strictEqual(result === newSheet, true, "Should return newly created sheet");
        assert.strictEqual(newSheet.name, "NewSheet", "New sheet should be named correctly");
    }

    // Test 3: Sheet already exists → clear cells so it can be used as a fresh sheet
    console.log("\n[Test Suite] NewWorksheet - existing sheet cleared for reuse");
    {
        const existingSheet = createMockWorksheet("DataSheet");
        const sheetsMap = new Map<string, any>();
        sheetsMap.set("DataSheet", existingSheet);

        const mockWb: any = {
            worksheets: createMockWorksheets(sheetsMap, () => createMockWorksheet(""))
        };

        const result = vbaTest.run('NewWorksheet', ["DataSheet", mockWb]);

        // Should return the same existing sheet object
        assert.strictEqual(result === existingSheet, true, "Should return existing sheet");
        // Cells.Delete should have been called to clear all content
        assert.strictEqual(existingSheet.cellsDeleted, true, "Cells should be cleared so the sheet can be reused as new");
    }

    console.log("\n--- All tests passed! ---");
}

main().catch(console.error);
