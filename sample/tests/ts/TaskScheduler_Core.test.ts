import { VBARunner, assert } from '../../../test-libs/test-runner';

async function main() {
    console.log("--- Starting VBA Unit Tests ---");
    const vbaTest = new VBARunner('sample/src/vba/TaskScheduler_Core.vba');

    // Initialize Config types (shared across tests that need them)
    const taskCfg = vbaTest.run('InitTaskConfig', []);
    const calCfg = vbaTest.run('InitCalendarConfig', []);

    // Test #1: CalcBaseStartIdx
    console.log("\n[Test Suite] CalcBaseStartIdx");

    // Level 1: Always starts at 1, regardless of parent finish
    assert.strictEqual(vbaTest.run('CalcBaseStartIdx', [1, 10, 1.0]), 1, "Level 1 ignores parent finish");

    // Level 2+ with parent alloc < 0.5: Starts on SAME day
    assert.strictEqual(vbaTest.run('CalcBaseStartIdx', [2, 10, 0.49]), 10, "Alloc < 0.5 starts same day");
    assert.strictEqual(vbaTest.run('CalcBaseStartIdx', [3, 5, 0.0]), 5, "Alloc 0 starts same day");

    // Level 2+ with parent alloc >= 0.5: Starts NEXT day
    assert.strictEqual(vbaTest.run('CalcBaseStartIdx', [2, 10, 0.5]), 11, "Alloc >= 0.5 starts next day");
    assert.strictEqual(vbaTest.run('CalcBaseStartIdx', [4, 20, 1.0]), 21, "Alloc 1.0 starts next day");

    // Test #2: GetMaxDailyLoad
    console.log("\n[Test Suite] GetMaxDailyLoad");

    // Setup simulated VBA Dictionary using JS Map proxy
    const mockDict = new Map<string, number>();
    mockDict.set("Alice", 0.5);
    mockDict.set("Bob", 0.8);

    const capacityLimits: any = function (key: string) {
        return mockDict.get(key);
    };
    capacityLimits.Exists = function (key: string) {
        return mockDict.has(key);
    };

    // Default when assignee is empty/not explicitly handled properly by dictionary vs missing
    assert.strictEqual(vbaTest.run('GetMaxDailyLoad', ["Charlie", capacityLimits]), 1.0, "Missing user defaults to 1.0");
    assert.strictEqual(vbaTest.run('GetMaxDailyLoad', ["Alice", capacityLimits]), 0.5, "Known user returns configured capacity");
    assert.strictEqual(vbaTest.run('GetMaxDailyLoad', ["Bob", capacityLimits]), 0.8, "Known user returns configured capacity");

    // Test #3: CalcDailyAllocation
    console.log("\n[Test Suite] CalcDailyAllocation");

    // Standard Task logic
    assert.strictEqual(vbaTest.run('CalcDailyAllocation', [1.0, 0.5, false]), 0.5, "Standard: exact fit");
    assert.strictEqual(vbaTest.run('CalcDailyAllocation', [0.25, 0.5, false]), 0.25, "Standard: limited by capacity");
    assert.strictEqual(vbaTest.run('CalcDailyAllocation', [1.0, 0.1, false]), 0, "Standard: rounds to 0 if < 0.125 needed");

    // Micro-Task logic
    assert.strictEqual(vbaTest.run('CalcDailyAllocation', [1.0, 0.1, true]), 0.1, "Micro: capacity 1.0 > 0.1, gets 0.1");
    assert.strictEqual(vbaTest.run('CalcDailyAllocation', [0.05, 0.1, true]), 0, "Micro: capacity < 0.1, skips");

    // Test #4: UpdateLevelFinish (ByRef array mutation)
    console.log("\n[Test Suite] UpdateLevelFinish");

    const levelFinish = [0, 10, 5];
    const levelAlloc = [0.0, 0.5, 1.0];

    vbaTest.run('UpdateLevelFinish', [1, 15, 0.8, levelFinish, levelAlloc]);
    assert.strictEqual(levelFinish[1], 15, "Level 1 finish index updated due to > 10");
    assert.strictEqual(levelAlloc[1], 0.8, "Level 1 alloc updated due to index change");

    // Same day index (15), but smaller alloc (0.1) -> No change
    vbaTest.run('UpdateLevelFinish', [1, 15, 0.1, levelFinish, levelAlloc]);
    assert.strictEqual(levelAlloc[1], 0.8, "Level 1 alloc unchanged, 0.1 not > 0.8");

    // Same day index (15), but larger alloc (1.0) -> Alloc updates
    vbaTest.run('UpdateLevelFinish', [1, 15, 1.0, levelFinish, levelAlloc]);
    assert.strictEqual(levelAlloc[1], 1.0, "Level 1 alloc updated, 1.0 > 0.8");

    console.log("\n[Test Suite] BuildCapacityDict");

    const mockConfigData = [
        null, // row 0
        [null, "Alice", 0.5],   // row 1
        [null, "Bob", null],    // row 2
        [null, "   ", 2.0],     // row 3
        [null, "Dan", 1.25]     // row 4
    ];

    const resultDict = vbaTest.run('BuildCapacityDict', [mockConfigData]);

    // Check if the Map/Dictionary populated correctly
    assert.strictEqual(resultDict.__map__.get("Alice"), 0.5, "Alice config matches");
    assert.strictEqual(resultDict.__map__.get("Bob"), 1.0, "Bob config defaults to 1.0");
    assert.strictEqual(resultDict.__map__.get("Dan"), 1.25, "Dan config matches");
    assert.isFalse(resultDict.exists("   ") || resultDict.exists(""), "Empty name skipped");
    assert.strictEqual(resultDict.__map__.size, 3, "Only Alice, Bob, and Dan were added");

    console.log("\n[Test Suite] ScanLockedRows");

    const mockMetaData = [
        null,
        // Row 1: Unlocked, Alice
        Array(20).fill("").map((_, i) => i === 13 ? "" : (i === 17 ? "Alice" : "")),
        // Row 2: Locked, Bob
        Array(20).fill("").map((_, i) => i === 13 ? "L" : (i === 17 ? "Bob" : "")),
        // Row 3: Locked, Alice
        Array(20).fill("").map((_, i) => i === 13 ? "L" : (i === 17 ? "Alice" : ""))
    ];

    const mockGridData = [
        null,
        [null, 1.0, 0.0, 0.0], // Row 1 (Unlocked) - should be ignored by ScanLockedRows
        [null, 0.0, 0.5, 0.5], // Row 2 (Bob, Locked)
        [null, 0.5, 0.0, 0.25] // Row 3 (Alice, Locked)
    ];

    const assigneeUsage = {
        __isVbaDict__: true,
        __map__: new Map<string, any>(),
        add: function (k: string, v: any) { this.__map__.set(k, v); },
        exists: function (k: string) { return this.__map__.has(k); }
    };

    vbaTest.run('ScanLockedRows', [taskCfg, 3, 3, mockMetaData, mockGridData, assigneeUsage]);

    // Bob has 0.5 on Day 2, 0.5 on Day 3
    const bobUsage = assigneeUsage.__map__.get("Bob");
    assert.strictEqual(bobUsage[2], 0.5, "Bob day 2 usage pre-allocated");
    assert.strictEqual(bobUsage[3], 0.5, "Bob day 3 usage pre-allocated");

    // Alice is Row 1 (Unlocked) padding to 0, Row 3 (Locked) adding 0.5 and 0.25
    const aliceUsage = assigneeUsage.__map__.get("Alice");
    assert.strictEqual(aliceUsage[1], 0.5, "Alice day 1 usage pre-allocated");
    assert.strictEqual(aliceUsage[2], 0.0, "Alice day 2 remains 0");
    assert.strictEqual(aliceUsage[3], 0.25, "Alice day 3 usage pre-allocated");

    console.log("\n[Test Suite] ClearTaskGridRow");

    const gridToClear: any[] = [
        null,
        [null, 0.5, 0.5, 0.5], // Row 1
        [null, 1.0, 1.0, 1.0], // Row 2
        [null, 0.25, 0.0, 0.5] // Row 3
    ];

    vbaTest.run('ClearTaskGridRow', [2, 3, gridToClear]);

    assert.strictEqual(gridToClear[2][1], null, "Row 2 Day 1 cleared");
    assert.strictEqual(gridToClear[2][2], null, "Row 2 Day 2 cleared");
    assert.strictEqual(gridToClear[2][3], null, "Row 2 Day 3 cleared");
    assert.strictEqual(gridToClear[1][1], 0.5, "Row 1 unaffected");

    console.log("\n[Test Suite] FindLockedTaskFinish");
    const mockLockedGrid = [
        null, // 0-based offset
        [null, 0.5, 1.0, 0.0, 0.0], // Row 1 (finishes on day 2 with 1.0)
        [null, 0.0, 0.0, 0.0, 0.0]  // Row 2 (empty)
    ];

    const resRow1 = vbaTest.run('TestFindLockedTaskFinish', [1, 4, mockLockedGrid]);
    assert.strictEqual(resRow1, "2|1", "Row 1 Finish Idx 2, Alloc 1.0");

    const resRow2 = vbaTest.run('TestFindLockedTaskFinish', [2, 4, mockLockedGrid]);
    assert.strictEqual(resRow2, "0|0", "Row 2 empty Finish Idx 0, Alloc 0");

    console.log("\n[Test Suite] ScheduleUnlockedTask");

    const mockMetaSchedule = [
        null,
        // Row 1: Alice, duration = 1.0, lag = 1 (col 9)
        Array(20).fill("").map((_, i) => i === 9 ? 1 : (i === 15 ? 1.0 : (i === 17 ? "Alice" : ""))),
        // Row 2: Bob, duration = 0.5, lag = ""
        Array(20).fill("").map((_, i) => i === 9 ? "" : (i === 15 ? 0.5 : (i === 17 ? "Bob" : ""))),
        // Row 3: Charlie, duration = 0.1 (microtask true), lag = ""
        Array(20).fill("").map((_, i) => i === 9 ? "" : (i === 15 ? 0.1 : (i === 17 ? "Charlie" : "")))
    ];

    const mockHolidayData = [
        null,
        [null, "", "休", "", "", "休"] // Day 2 and Day 5 are holidays
    ];

    const mockGridSchedule: any[] = [
        null,
        [null, 0, 0, 0, 0, 0], // Row 1
        [null, 0, 0, 0, 0, 0], // Row 2
        [null, 0, 0, 0, 0, 0]  // Row 3
    ];

    const mockPersonUsage = {
        __isVbaDict__: true,
        __map__: new Map<string, any>(),
        add: function (k: string, v: any) { this.__map__.set(k, v); },
        exists: function (k: string) { return this.__map__.has(k); }
    };

    const resRow1Schedule = vbaTest.run('TestScheduleUnlockedTask', [taskCfg, calCfg, 1, 5, 1, mockMetaSchedule, mockHolidayData, capacityLimits, mockGridSchedule, mockPersonUsage]);
    assert.strictEqual(resRow1Schedule, "4|0.5", "Row 1 Finish Idx 4, Alloc 0.5");
    assert.strictEqual(mockGridSchedule[1][2], null, "Alice Day 2 is holiday (cleared to null)");
    assert.strictEqual(mockGridSchedule[1][3], 0.5, "Alice Day 3 gets 0.5 (capacity limit)");
    assert.strictEqual(mockGridSchedule[1][4], 0.5, "Alice Day 4 gets 0.5");

    const resRow2Schedule = vbaTest.run('TestScheduleUnlockedTask', [taskCfg, calCfg, 2, 5, 3, mockMetaSchedule, mockHolidayData, capacityLimits, mockGridSchedule, mockPersonUsage]);
    assert.strictEqual(resRow2Schedule, "3|0.5", "Row 2 Finish Idx 3, Alloc 0.5");
    assert.strictEqual(mockGridSchedule[2][3], 0.5, "Bob Day 3 gets 0.5");

    // Test #9: CalcNumRows / CalcNumDays
    console.log("\n[Test Suite] CalcNumRows / CalcNumDays");
    // taskCfg.ROW_START = 19, calCfg.COL_CALENDAR_START = 24
    assert.strictEqual(vbaTest.run('CalcNumRows', [19, taskCfg]), 1, "lastRow=19 -> 1 row");
    assert.strictEqual(vbaTest.run('CalcNumRows', [25, taskCfg]), 7, "lastRow=25 -> 7 rows");
    assert.strictEqual(vbaTest.run('CalcNumRows', [18, taskCfg]), 0, "lastRow=18 -> 0 rows (no data)");

    assert.strictEqual(vbaTest.run('CalcNumDays', [24, calCfg]), 1, "lastCol=24 -> 1 day");
    assert.strictEqual(vbaTest.run('CalcNumDays', [30, calCfg]), 7, "lastCol=30 -> 7 days");
    assert.strictEqual(vbaTest.run('CalcNumDays', [23, calCfg]), 0, "lastCol=23 -> 0 days (no data)");

    // Test #10: IsRowLocked / GetAssigneeName / GetTaskLevel
    console.log("\n[Test Suite] IsRowLocked / GetAssigneeName / GetTaskLevel");
    // taskCfg.COL_LOCK = 13, STR_LOCK_MARK = "L"
    // taskCfg.COL_ASSIGNEE = 17
    // taskCfg.COL_LEVEL = 8
    const mockMetaRow: any = [[null]]; // dummy row 0
    const row1: any[] = new Array(18).fill("");
    row1[13] = "L";      // COL_LOCK
    row1[17] = " Alice "; // COL_ASSIGNEE (with spaces)
    row1[8] = 2;          // COL_LEVEL
    mockMetaRow.push(row1);

    const row2: any[] = new Array(18).fill("");
    row2[13] = "";        // not locked
    row2[17] = "Bob";
    row2[8] = "";         // empty level
    mockMetaRow.push(row2);

    assert.strictEqual(vbaTest.run('IsRowLocked', [mockMetaRow, 1, taskCfg]).valueOf(), -1, "Row 1 is locked");
    assert.strictEqual(vbaTest.run('IsRowLocked', [mockMetaRow, 2, taskCfg]).valueOf(), 0, "Row 2 is not locked");
    assert.strictEqual(vbaTest.run('GetAssigneeName', [mockMetaRow, 1, taskCfg]), "Alice", "Row 1 assignee trimmed");
    assert.strictEqual(vbaTest.run('GetAssigneeName', [mockMetaRow, 2, taskCfg]), "Bob", "Row 2 assignee");
    assert.strictEqual(vbaTest.run('GetTaskLevel', [mockMetaRow, 1, taskCfg]), 2, "Row 1 level = 2");
    assert.strictEqual(vbaTest.run('GetTaskLevel', [mockMetaRow, 2, taskCfg]), 0, "Row 2 empty level = 0");

    console.log("\n--- All tests passed! ---");
}

main().catch(console.error);
