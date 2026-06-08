Option Explicit

' Tier 5（COM 型ライブラリ）よりも Tier 1-4 識別子が優先されることを検証するテスト

' Tier 1（ローカル変数）が Tier 5 の COM クラス名より優先される
Sub Test_LocalVarShadowsComClass(assert)
    Dim Dictionary As Long
    Dictionary = 42
    assert.Assert Dictionary, 42, "Tier1: Dim Dictionary As Long shadows Scripting.Dictionary"
End Sub

' Tier 1（ローカル変数）が Tier 5 の COM 名前空間より優先される
Sub Test_LocalVarShadowsComNamespace(assert)
    Dim Scripting As String
    Scripting = "local"
    assert.Assert Scripting, "local", "Tier1: Dim Scripting As String shadows COM namespace"
End Sub

' Tier 5: 影のない場合は COM オブジェクトが CreateObject 経由で使用可能
Sub Test_ComObjectAccessibleViaCreateObject(assert)
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "key", "value"
    assert.Assert d.Item("key"), "value", "Tier5: Scripting.Dictionary accessible via CreateObject"
    assert.Assert d.Count(), 1, "Tier5: Scripting.Dictionary.Count works"
End Sub

' Tier 5: 影のない場合は短縮クラス名 Dictionary が使用可能
Sub Test_ShortClassNameAccessible(assert)
    Dim d As Object
    Set d = New Dictionary
    d.Add "x", 99
    assert.Assert d.Item("x"), 99, "Tier5: New Dictionary (short name) accessible when not shadowed"
End Sub
