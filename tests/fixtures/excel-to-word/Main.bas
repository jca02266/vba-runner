' Excel VBA から Word オブジェクトを生成して操作するコード
' CreateObject("Word.Application") で新しい Word インスタンスを起動し、
' 文書を作成・編集・保存する典型的なパターン。

' --- 基本: CreateObject で Word を起動して文書を作成 ---
Function CreateWordDoc() As String
    Dim wordApp As Object
    Dim wordDoc As Object

    Set wordApp = CreateObject("Word.Application")
    wordApp.Visible = False

    Set wordDoc = wordApp.Documents.Add()
    wordDoc.Content.Text = "Hello from Excel VBA"

    CreateWordDoc = wordDoc.Content.Text
    wordDoc.Close False
    wordApp.Quit
End Function

' --- ブックマーク経由でテンプレートに値を差し込む ---
Function FillTemplate() As String
    Dim wordApp As Object
    Dim wordDoc As Object

    Set wordApp = CreateObject("Word.Application")
    Set wordDoc = wordApp.Documents.Add()

    ' ブックマーク位置に値を挿入
    wordDoc.Bookmarks("bmkCompany").Range.Text = "Acme Corp"
    wordDoc.Bookmarks("bmkDate").Range.Text = "2024-03-15"

    FillTemplate = wordDoc.Bookmarks("bmkCompany").Range.Text & " / " & _
                   wordDoc.Bookmarks("bmkDate").Range.Text

    wordApp.Quit
End Function

' --- 複数ドキュメントを開いて処理 ---
Function ProcessMultipleDocs() As Long
    Dim wordApp As Object
    Set wordApp = CreateObject("Word.Application")

    Dim doc1 As Object
    Dim doc2 As Object
    Set doc1 = wordApp.Documents.Add()
    Set doc2 = wordApp.Documents.Add()

    doc1.Content.Text = "Document 1"
    doc2.Content.Text = "Document 2"

    ProcessMultipleDocs = wordApp.Documents.Count
    wordApp.Quit
End Function

' --- 段落を追加して件数を確認 ---
Function AddParagraphs() As Long
    Dim wordApp As Object
    Dim wordDoc As Object

    Set wordApp = CreateObject("Word.Application")
    Set wordDoc = wordApp.Documents.Add()

    wordDoc.setContent "Line 1" & Chr(10) & "Line 2" & Chr(10) & "Line 3"

    AddParagraphs = wordDoc.Paragraphs.Count
    wordApp.Quit
End Function

' --- Selection を使ってテキストを入力 ---
Function TypeIntoDoc() As String
    Dim wordApp As Object
    Set wordApp = CreateObject("Word.Application")

    wordApp.Selection.TypeText "Hello"
    wordApp.Selection.TypeText ", World"

    TypeIntoDoc = wordApp.Selection.Text
    wordApp.Quit
End Function
