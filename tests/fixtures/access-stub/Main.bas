' Access API スタブモードの動作確認用

' DoCmd ノーオプ
Function TestDoCmd() As String
    DoCmd.SetWarnings False
    DoCmd.RunSQL "INSERT INTO t VALUES (1)"
    DoCmd.SetWarnings True
    DoCmd.OpenForm "frmTest"
    DoCmd.Close 2, "frmTest"
    TestDoCmd = "ok"
End Function

' CurrentDb → OpenRecordset（空テーブル）→ EOF = True
Function TestEmptyRecordset() As Boolean
    Dim db As Object
    Dim rs As Object
    Set db = CurrentDb()
    Set rs = db.OpenRecordset("TestTable")
    TestEmptyRecordset = rs.EOF
    rs.Close
End Function

' Recordset にデータを注入して Fields を読む
Function TestRecordsetRead() As String
    Dim db As Object
    Dim rs As Object
    Set db = CurrentDb()
    Set rs = db.OpenRecordset("Customers")
    If Not rs.EOF Then
        TestRecordsetRead = rs.Fields("Name").Value
    Else
        TestRecordsetRead = "(empty)"
    End If
    rs.Close
End Function

' Recordset のループ
Function TestRecordsetLoop() As Long
    Dim db As Object
    Dim rs As Object
    Dim count As Long
    Set db = CurrentDb()
    Set rs = db.OpenRecordset("Orders")
    count = 0
    Do While Not rs.EOF
        count = count + 1
        rs.MoveNext
    Loop
    rs.Close
    TestRecordsetLoop = count
End Function

' Forms("name").Control アクセス
Function TestFormControl() As String
    TestFormControl = Forms("frmCustomer").txtName.Value
End Function

' Application プロパティ
Function TestAppProps() As String
    Application.Echo False
    Application.ScreenUpdating = False
    Application.Echo True
    TestAppProps = Application.Name
End Function
