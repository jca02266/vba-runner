' Word API スタブモードの動作確認用

' ActiveDocument.Content.Text の読み書き
Function TestContent() As String
    ActiveDocument.Content.Text = "Hello, World"
    TestContent = ActiveDocument.Content.Text
End Function

' Selection.TypeText
Function TestTypeText() As String
    Selection.TypeText "Hello"
    Selection.TypeText " World"
    TestTypeText = Selection.Text
End Function

' Bookmarks への読み書き
Function TestBookmark() As String
    ActiveDocument.Bookmarks("bmkTitle").Range.Text = "Report 2024"
    TestBookmark = ActiveDocument.Bookmarks("bmkTitle").Range.Text
End Function

' Paragraphs の For Each イテレーション
Function TestParagraphLoop() As Long
    Dim p As Object
    Dim count As Long
    count = 0
    For Each p In ActiveDocument.Paragraphs
        count = count + 1
    Next p
    TestParagraphLoop = count
End Function

' Tables(1).Cell(r, c) の読み書き
Function TestTableCell() As String
    ActiveDocument.Tables(1).Cell(1, 1).Range.Text = "Header"
    TestTableCell = ActiveDocument.Tables(1).Cell(1, 1).Range.Text
End Function

' Application プロパティ
Function TestAppName() As String
    Application.ScreenUpdating = False
    Application.ScreenUpdating = True
    TestAppName = Application.Name
End Function

' Find/Replace（スタブは常に Not Found）
Function TestFind() As Boolean
    With Selection.Find
        .Text = "foo"
        .Replacement.Text = "bar"
        .Execute
    End With
    TestFind = Selection.Find.Found
End Function

' Documents.Open（スタブ）
Function TestDocumentsOpen() As String
    Dim d As Object
    Set d = Documents.Open("C:\reports\sales.docx")
    TestDocumentsOpen = d.Name
End Function
