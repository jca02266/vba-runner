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
Sub Test_SetupInitializesCounter()
    Debug.Assert testCounter = 0, "SetUp failed to initialize counter"
End Sub

' Test 2: Verify SetUp sets state
Sub Test_SetupInitializesState()
    Debug.Assert testState = "initialized", "SetUp failed to initialize state"
End Sub

' Test 3: Modify counter
Sub Test_CounterIncrement()
    testCounter = testCounter + 1
    testCounter = testCounter + 1
    Debug.Assert testCounter = 2, "Counter increment failed"
End Sub

' Test 4: String operations
Sub Test_StringConcat()
    testState = testState & " and working"
    Debug.Assert testState = "initialized and working", "String concat failed"
End Sub

' Test 5: Boolean operations
Sub Test_LogicalAnd()
    Dim result As Boolean
    result = (testCounter = 0) And (testState = "initialized")
    Debug.Assert result, "Boolean operations failed"
End Sub
