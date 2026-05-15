Option Explicit

' circular_reference_terminate_test.vba
' VBA source tests for circular reference and Class_Terminate behavior
' Tests verify that Class_Terminate is called at most once per object instance

Sub Test_BasicTerminate(assert)
    Dim helper As CircularTestHelper
    Dim objA As CircularReferenceATest

    Set helper = New CircularTestHelper
    Set objA = New CircularReferenceATest
    Set objA.Helper = helper

    Set objA = Nothing

    assert.Assert helper.TerminateCount, 1, "Basic terminate"
End Sub

Sub Test_MutualReferences(assert)
    Dim helper As CircularTestHelper
    Dim objA As CircularReferenceATest
    Dim objB As CircularReferenceBTest

    Set helper = New CircularTestHelper
    Set objA = New CircularReferenceATest
    Set objB = New CircularReferenceBTest

    Set objA.Helper = helper
    Set objB.Helper = helper
    Set objA.RefB = objB
    Set objB.RefA = objA

    Set objA = Nothing
    Set objB = Nothing

    assert.Assert helper.TerminateCount, 2, "Mutual references"
End Sub

Sub Test_TerminateNotCalledTwice(assert)
    Dim helper As CircularTestHelper
    Dim objA As CircularReferenceATest

    Set helper = New CircularTestHelper
    Set objA = New CircularReferenceATest
    Set objA.Helper = helper

    Set objA = Nothing
    Set objA = Nothing

    assert.Assert helper.TerminateCount, 1, "Not called twice"
End Sub

Sub Test_MultipleObjects(assert)
    Dim helper As CircularTestHelper
    Dim objA1 As CircularReferenceATest
    Dim objA2 As CircularReferenceATest
    Dim objA3 As CircularReferenceATest

    Set helper = New CircularTestHelper
    Set objA1 = New CircularReferenceATest
    Set objA2 = New CircularReferenceATest
    Set objA3 = New CircularReferenceATest

    Set objA1.Helper = helper
    Set objA2.Helper = helper
    Set objA3.Helper = helper

    Set objA1 = Nothing
    Set objA2 = Nothing
    Set objA3 = Nothing

    assert.Assert helper.TerminateCount, 3, "Multiple objects"
End Sub

Sub Test_CircularChainCleanup(assert)
    Dim helper As CircularTestHelper
    Dim objA As CircularReferenceATest
    Dim objB As CircularReferenceBTest
    Dim objA2 As CircularReferenceATest

    Set helper = New CircularTestHelper
    Set objA = New CircularReferenceATest
    Set objB = New CircularReferenceBTest
    Set objA2 = New CircularReferenceATest

    Set objA.Helper = helper
    Set objB.Helper = helper
    Set objA2.Helper = helper

    Set objA.RefB = objB
    Set objB.RefA = objA2
    Set objA2.RefB = objA

    Set objA = Nothing
    Set objB = Nothing
    Set objA2 = Nothing

    assert.Assert helper.TerminateCount, 3, "Circular chain cleanup"
End Sub
