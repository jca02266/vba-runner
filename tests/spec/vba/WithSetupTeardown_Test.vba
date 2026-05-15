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

' Test 1: Verify SetUp initializes state
Function TestSetupInitializesCounter() As Boolean
    TestSetupInitializesCounter = (testCounter = 0)
End Function

' Test 2: Verify SetUp sets state
Function TestSetupInitializesState() As Boolean
    TestSetupInitializesState = (testState = "initialized")
End Function

' Test 3: Modify counter
Function TestCounterIncrement() As Boolean
    testCounter = testCounter + 1
    testCounter = testCounter + 1
    TestCounterIncrement = (testCounter = 2)
End Function

' Test 4: String operations
Function TestStringConcat() As Boolean
    testState = testState & " and working"
    TestStringConcat = (testState = "initialized and working")
End Function

' Test 5: Boolean operations
Function TestLogicalAnd() As Boolean
    Dim result As Boolean
    result = (testCounter = 0) And (testState = "initialized")
    TestLogicalAnd = result
End Function
