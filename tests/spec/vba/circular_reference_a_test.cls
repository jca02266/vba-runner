' circular_reference_a_test.cls
' Test class for circular reference testing

Class CircularReferenceATest
    Public RefB As Object
    Public Helper As Object

    Sub Class_Terminate()
        If Helper Is Nothing Then Exit Sub
        Helper.TerminateCount = Helper.TerminateCount + 1
        Helper.TerminateLog = Helper.TerminateLog & "A;"
        Set RefB = Nothing
    End Sub
End Class
