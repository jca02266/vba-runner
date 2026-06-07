Option Explicit

' VBA source tests for circular reference and Class_Terminate behavior
' Tests verify that Class_Terminate is called at most once per object instance

Sub Test_BasicTerminate(assert)
    Dim helper As Helper
    Dim objA As RefA

    Set helper = New Helper
    Set objA = New RefA
    Set objA.Helper = helper

    Set objA = Nothing

    assert.Assert helper.TerminateCount, 1, "Basic terminate"
End Sub

Sub Test_MutualReferences(assert)
    Dim helper As Helper
    Dim objA As RefA
    Dim objB As RefB

    Set helper = New Helper
    Set objA = New RefA
    Set objB = New RefB

    Set objA.Helper = helper
    Set objB.Helper = helper
    Set objA.ObjB = objB
    Set objB.ObjA = objA

    Set objA = Nothing
    Set objB = Nothing

    assert.Assert helper.TerminateCount, 2, "Mutual references"
End Sub

Sub Test_TerminateNotCalledTwice(assert)
    Dim helper As Helper
    Dim objA As RefA

    Set helper = New Helper
    Set objA = New RefA
    Set objA.Helper = helper

    Set objA = Nothing
    Set objA = Nothing

    assert.Assert helper.TerminateCount, 1, "Not called twice"
End Sub

Sub Test_MultipleObjects(assert)
    Dim helper As Helper
    Dim objA1 As RefA
    Dim objA2 As RefA
    Dim objA3 As RefA

    Set helper = New Helper
    Set objA1 = New RefA
    Set objA2 = New RefA
    Set objA3 = New RefA

    Set objA1.Helper = helper
    Set objA2.Helper = helper
    Set objA3.Helper = helper

    Set objA1 = Nothing
    Set objA2 = Nothing
    Set objA3 = Nothing

    assert.Assert helper.TerminateCount, 3, "Multiple objects"
End Sub

Sub Test_CircularChainCleanup(assert)
    Dim helper As Helper
    Dim objA As RefA
    Dim objB As RefB
    Dim objA2 As RefA

    Set helper = New Helper
    Set objA = New RefA
    Set objB = New RefB
    Set objA2 = New RefA

    Set objA.Helper = helper
    Set objB.Helper = helper
    Set objA2.Helper = helper

    Set objA.ObjB = objB
    Set objB.ObjA = objA2
    Set objA2.ObjB = objA

    Set objA = Nothing
    Set objB = Nothing
    Set objA2 = Nothing

    assert.Assert helper.TerminateCount, 3, "Circular chain cleanup"
End Sub
