' Test_circular_reference_terminate.vba
' VBA source tests for circular reference and Class_Terminate behavior
' Tests verify that Class_Terminate is called at most once per object instance

Sub Test_BasicTerminate()
    Dim helper As TestHelper
    Dim objA As CircularRefA

    Set helper = New TestHelper
    Set objA = New CircularRefA
    Set objA.Helper = helper

    Set objA = Nothing

    ' Class_Terminate should have been called once
    Debug.Assert helper.TerminateCount = 1, "Basic terminate failed"
    Debug.Print "Test_BasicTerminate passed"
End Sub

Sub Test_MutualReferences()
    Dim helper As TestHelper
    Dim objA As CircularRefA
    Dim objB As CircularRefB

    Set helper = New TestHelper
    Set objA = New CircularRefA
    Set objB = New CircularRefB

    Set objA.Helper = helper
    Set objB.Helper = helper
    Set objA.RefB = objB
    Set objB.RefA = objA

    ' Break the cycle
    Set objA = Nothing
    Set objB = Nothing

    ' Both should terminate exactly once
    Debug.Assert helper.TerminateCount = 2, "Mutual references terminate count failed: expected 2, got " & helper.TerminateCount
    Debug.Print "Test_MutualReferences passed"
End Sub

Sub Test_TerminateNotCalledTwice()
    Dim helper As TestHelper
    Dim objA As CircularRefA

    Set helper = New TestHelper
    Set objA = New CircularRefA
    Set objA.Helper = helper

    ' Set to Nothing twice - second call should not increment count
    Set objA = Nothing
    Set objA = Nothing

    Debug.Assert helper.TerminateCount = 1, "Terminate called more than once"
    Debug.Print "Test_TerminateNotCalledTwice passed"
End Sub

Sub Test_MultipleObjects()
    Dim helper As TestHelper
    Dim objA1 As CircularRefA
    Dim objA2 As CircularRefA
    Dim objA3 As CircularRefA

    Set helper = New TestHelper
    Set objA1 = New CircularRefA
    Set objA2 = New CircularRefA
    Set objA3 = New CircularRefA

    Set objA1.Helper = helper
    Set objA2.Helper = helper
    Set objA3.Helper = helper

    Set objA1 = Nothing
    Set objA2 = Nothing
    Set objA3 = Nothing

    ' All three should terminate
    Debug.Assert helper.TerminateCount = 3, "Multiple objects terminate failed"
    Debug.Print "Test_MultipleObjects passed"
End Sub

Sub Test_CircularChainCleanup()
    Dim helper As TestHelper
    Dim objA As CircularRefA
    Dim objB As CircularRefB
    Dim objA2 As CircularRefA

    Set helper = New TestHelper
    Set objA = New CircularRefA
    Set objB = New CircularRefB
    Set objA2 = New CircularRefA

    Set objA.Helper = helper
    Set objB.Helper = helper
    Set objA2.Helper = helper

    ' Create chain: A -> B -> A2 (via RefB/RefA properties)
    Set objA.RefB = objB
    Set objB.RefA = objA2
    Set objA2.RefB = objA

    ' Break chain
    Set objA = Nothing
    Set objB = Nothing
    Set objA2 = Nothing

    ' All three should terminate
    Debug.Assert helper.TerminateCount = 3, "Circular chain cleanup failed"
    Debug.Print "Test_CircularChainCleanup passed"
End Sub
