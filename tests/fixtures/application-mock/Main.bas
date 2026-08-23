Option Explicit

Function ReadPathSeparator() As String
    ReadPathSeparator = Application.PathSeparator
End Function

Function ToggleDisplayAlerts() As Boolean
    Application.DisplayAlerts = False
    ToggleDisplayAlerts = Application.DisplayAlerts
End Function

Function ReadWorksheetName() As String
    ReadWorksheetName = Application.Worksheets("Sheet1").Name
End Function
