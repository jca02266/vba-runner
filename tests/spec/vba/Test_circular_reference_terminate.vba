' Test_circular_reference_terminate.vba
' VBA source tests for circular reference and Class_Terminate behavior
' Tests verify that Class_Terminate is called at most once per object instance

Class TestHelper
    Public TerminateCount As Integer
    Public TerminateLog As String

    Public Sub Reset()
        TerminateCount = 0
        TerminateLog = ""
    End Sub
End Class

Class SimpleObject
    Public Helper As Object

    Public Sub Class_Terminate()
        If Helper Is Nothing Then Exit Sub
        Helper.TerminateCount = Helper.TerminateCount + 1
    End Sub
End Class

Class CircularA
    Public RefB As Object
    Public Helper As Object

    Public Sub Class_Terminate()
        If Helper Is Nothing Then Exit Sub
        Helper.TerminateCount = Helper.TerminateCount + 1
        Helper.TerminateLog = Helper.TerminateLog & "A;"
        Set RefB = Nothing
    End Sub
End Class

Class CircularB
    Public RefA As Object
    Public Helper As Object

    Public Sub Class_Terminate()
        If Helper Is Nothing Then Exit Sub
        Helper.TerminateCount = Helper.TerminateCount + 1
        Helper.TerminateLog = Helper.TerminateLog & "B;"
        Set RefA = Nothing
    End Sub
End Class

Sub Test_BasicTerminate()
    Dim helper As TestHelper
    Dim obj As SimpleObject

    Set helper = New TestHelper
    Set obj = New SimpleObject
    Set obj.Helper = helper

    Set obj = Nothing

    ' Class_Terminate should have been called once
    Debug.Assert helper.TerminateCount = 1, "Basic terminate failed"
    Debug.Print "Test_BasicTerminate passed"
End Sub

Sub Test_MutualReferences()
    Dim helper As TestHelper
    Dim objA As CircularA
    Dim objB As CircularB

    Set helper = New TestHelper
    Set objA = New CircularA
    Set objB = New CircularB

    Set objA.Helper = helper
    Set objB.Helper = helper
    Set objA.RefB = objB
    Set objB.RefA = objA

    ' Break the cycle
    Set objA = Nothing
    Set objB = Nothing

    ' Both should terminate exactly once
    Debug.Assert helper.TerminateCount = 2, "Mutual references terminate count failed"
    Debug.Print "Test_MutualReferences passed"
End Sub

Sub Test_TerminateNotCalledTwice()
    Dim helper As TestHelper
    Dim obj1 As SimpleObject
    Dim obj2 As SimpleObject

    Set helper = New TestHelper
    Set obj1 = New SimpleObject
    Set obj2 = New SimpleObject

    Set obj1.Helper = helper
    Set obj2.Helper = helper

    ' Set to Nothing twice - second call should not increment count
    Set obj1 = Nothing
    Set obj1 = Nothing

    Debug.Assert helper.TerminateCount = 1, "Terminate called more than once"
    Debug.Print "Test_TerminateNotCalledTwice passed"
End Sub

Sub Test_MultipleObjects()
    Dim helper As TestHelper
    Dim obj1 As SimpleObject
    Dim obj2 As SimpleObject
    Dim obj3 As SimpleObject

    Set helper = New TestHelper
    Set obj1 = New SimpleObject
    Set obj2 = New SimpleObject
    Set obj3 = New SimpleObject

    Set obj1.Helper = helper
    Set obj2.Helper = helper
    Set obj3.Helper = helper

    Set obj1 = Nothing
    Set obj2 = Nothing
    Set obj3 = Nothing

    ' All three should terminate
    Debug.Assert helper.TerminateCount = 3, "Multiple objects terminate failed"
    Debug.Print "Test_MultipleObjects passed"
End Sub

Sub Test_CircularChainCleanup()
    Dim helper As TestHelper
    Dim objA As CircularA
    Dim objB As CircularB
    Dim objC As CircularA

    Set helper = New TestHelper
    Set objA = New CircularA
    Set objB = New CircularB
    Set objC = New CircularA

    Set objA.Helper = helper
    Set objB.Helper = helper
    Set objC.Helper = helper

    ' Create chain: A -> B -> C -> A (via RefB/RefA properties)
    Set objA.RefB = objB
    Set objB.RefA = objC
    Set objC.RefB = objA

    ' Break chain
    Set objA = Nothing
    Set objB = Nothing
    Set objC = Nothing

    ' All three should terminate
    Debug.Assert helper.TerminateCount = 3, "Circular chain cleanup failed"
    Debug.Print "Test_CircularChainCleanup passed"
End Sub

Sub Setup()
    ' Optional: Initialize test state if needed
End Sub

Sub TearDown()
    ' Optional: Clean up after each test
End Sub
