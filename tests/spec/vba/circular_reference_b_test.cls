' circular_reference_b_test.cls
' Test class for circular reference testing

Class CircularReferenceBTest
    Public RefA As Object
    Public Helper As Object

    Sub Class_Terminate()
        If Helper Is Nothing Then Exit Sub
        Helper.TerminateCount = Helper.TerminateCount + 1
        Helper.TerminateLog = Helper.TerminateLog & "B;"
        Set RefA = Nothing
    End Sub
End Class
