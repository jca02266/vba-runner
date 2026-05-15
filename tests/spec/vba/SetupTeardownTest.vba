' Test Suite with SetUp and TearDown
' This example demonstrates SetUp/TearDown support in the test runner

Option Explicit

' Global state for testing
Dim testCounter As Integer
Dim testState As String

' Setup routine - called before each test
Sub SetUp()
    testCounter = 0
    testState = "initialized"
End Sub

' Teardown routine - called after each test
Sub TearDown()
    testCounter = 0
    testState = ""
End Sub

' Test 1: Verify SetUp initializes counter
Sub Test_SetupInitializesCounter(assert)
    assert.Assert testCounter, 0, "SetUp initializes counter"
End Sub

' Test 2: Verify SetUp sets state
Sub Test_SetupInitializesState(assert)
    assert.Assert testState, "initialized", "SetUp initializes state"
End Sub

' Test 3: Modify counter
Sub Test_CounterIncrement(assert)
    testCounter = testCounter + 1
    testCounter = testCounter + 1
    assert.Assert testCounter, 2, "Counter increment"
End Sub

' Test 4: String operations
Sub Test_StringConcat(assert)
    testState = testState & " and working"
    assert.Assert testState, "initialized and working", "String concatenation"
End Sub

' Test 5: Boolean operations
Sub Test_LogicalAnd(assert)
    Dim result As Boolean
    result = (testCounter = 0) And (testState = "initialized")
    assert.IsTrue result, "Logical And"
End Sub
