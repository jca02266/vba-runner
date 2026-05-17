import { VBARunner, assert } from '../../../test-libs/test-runner';

async function main() {
    console.log("--- Starting LibBook.vba Unit Tests ---");
    const vbaTest = new VBARunner('sample/src/vba/LibBook.vba');

    // Mock Application object with SheetsInNewWorkbook property
    function createMockApplication(initialSheetCount: number): any {
        return { sheetsinnewworkbook: initialSheetCount };
    }

    // Mock Workbook object
    function createMockWorkbook(): any {
        return { name: 'NewBook.xlsx', sheets: [] };
    }

    // Mock Workbooks collection with Add method
    function createMockWorkbooks(addResult: any): any {
        return {
            add: function () { return addResult; }
        };
    }

    // Test 1: Normal case – creates workbook and restores SheetsInNewWorkbook
    console.log("\n[Test Suite] NewWorkbook - normal case");
    {
        const mockApp = createMockApplication(3); // default: 3 sheets
        const mockWb = createMockWorkbook();
        const mockWorkbooks = createMockWorkbooks(mockWb);

        vbaTest.set('application', mockApp);
        vbaTest.set('workbooks', mockWorkbooks);

        const result = vbaTest.run('NewWorkbook', [1]);

        assert.strictEqual(result === mockWb, true, "Should return new workbook");
        // SheetsInNewWorkbook should be restored to original value (3)
        assert.strictEqual(mockApp.sheetsinnewworkbook, 3, "SheetsInNewWorkbook should be restored to 3");
    }

    // Test 2: Custom SheetCount parameter
    console.log("\n[Test Suite] NewWorkbook - custom sheet count");
    {
        const mockApp = createMockApplication(5); // default: 5 sheets
        const mockWb = createMockWorkbook();
        let capturedSheetCount = 0;
        const mockWorkbooks = {
            add: function () {
                // Capture the SheetsInNewWorkbook value at time of Add
                capturedSheetCount = mockApp.sheetsinnewworkbook;
                return mockWb;
            }
        };

        vbaTest.set('application', mockApp);
        vbaTest.set('workbooks', mockWorkbooks);

        const result = vbaTest.run('NewWorkbook', [3]);

        assert.strictEqual(result === mockWb, true, "Should return new workbook");
        // During Add, SheetsInNewWorkbook should have been set to 3
        assert.strictEqual(capturedSheetCount, 3, "SheetsInNewWorkbook should be 3 during Add");
        // After function returns, SheetsInNewWorkbook should be restored to 5
        assert.strictEqual(mockApp.sheetsinnewworkbook, 5, "SheetsInNewWorkbook should be restored to 5");
    }

    // Test 3: Error case – Workbooks.Add throws, SheetsInNewWorkbook is still restored
    console.log("\n[Test Suite] NewWorkbook - error restores SheetsInNewWorkbook");
    {
        const mockApp = createMockApplication(4);
        const mockWorkbooks = {
            add: function () {
                throw new Error("Cannot create workbook");
            }
        };

        vbaTest.set('application', mockApp);
        vbaTest.set('workbooks', mockWorkbooks);

        let errorThrown = false;
        try {
            vbaTest.run('NewWorkbook', [2]);
        } catch {
            errorThrown = true;
        }

        // Error should have been re-raised via Err.Raise
        assert.strictEqual(errorThrown, true, "Error should be re-raised");
        // SheetsInNewWorkbook should still be restored to original value
        assert.strictEqual(mockApp.sheetsinnewworkbook, 4, "SheetsInNewWorkbook should be restored to 4 even on error");
    }

    console.log("\n--- All tests passed! ---");
}

main().catch(console.error);
