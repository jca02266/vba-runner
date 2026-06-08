Option Explicit

' Tier 2（このモジュール固有）として宣言
' CrossVar: TierHelper の Public CrossVar（Tier 3）と同名 → Tier 2 が優先
Private CrossVar As Long

' Tier 4 組み込みの Len と同名のモジュール変数 → Tier 2 が優先
Private Len As Long

' ----------------------------------------------------------------
' Tier 4: 組み込み関数が通常通り使える
' ----------------------------------------------------------------
Sub Test_Tier4BuiltinAccessible(assert)
    ' モジュールレベルの Len（Tier 2）は Long 型の変数
    ' ここでは組み込み Len を使いたい場合は VBA.Len で明示できる（本テストでは確認のみ）
    Dim s As String
    s = "hello"
    ' VBA. 修飾で Tier 4 組み込みに直接アクセス
    assert.Assert VBA.Len(s), 5, "Tier4: VBA.Len builtin accessible via VBA. qualifier"
End Sub

' ----------------------------------------------------------------
' Tier 2 > Tier 4: モジュールレベル変数が同名組み込みを隠蔽
' ----------------------------------------------------------------
Sub Test_Tier2OverTier4(assert)
    ' このモジュールの Len（Tier 2）は Long 型変数
    Len = 99
    assert.Assert Len, 99, "Tier2 > Tier4: module-level Dim Len shadows built-in Len function"
End Sub

' ----------------------------------------------------------------
' Tier 3: 他モジュールの Public 変数が Tier 3 として見える
' ----------------------------------------------------------------
Sub Test_Tier3CrossModuleAccess(assert)
    ' TierHelper.CrossVar は Tier 3
    ' このモジュールには CrossVar（Tier 2）があるので Tier 2 が優先されるが
    ' SharedName はこのモジュールに宣言がないので Tier 3 で解決される
    TierHelper.SharedName = 77  ' モジュール修飾で明示アクセス
    assert.Assert TierHelper.SharedName, 77, "Tier3: cross-module Public var accessible via qualification"
End Sub

' ----------------------------------------------------------------
' Tier 2 > Tier 3: 同名がある場合はローカルモジュールが優先
' ----------------------------------------------------------------
Sub Test_Tier2OverTier3(assert)
    ' TierHelper.CrossVar（Tier 3）と同名の CrossVar（Tier 2）がこのモジュールにある
    TierHelper.CrossVar = 10   ' Tier 3 側に書き込み
    CrossVar = 20              ' Tier 2 側（このモジュール）に書き込み
    assert.Assert CrossVar, 20, "Tier2 > Tier3: module CrossVar shadows TierHelper.CrossVar"
    assert.Assert TierHelper.CrossVar, 10, "Tier3: TierHelper.CrossVar unchanged by Tier2 assignment"
End Sub

' ----------------------------------------------------------------
' Tier 1 > Tier 2: ローカル変数がモジュール変数を隠蔽
' ----------------------------------------------------------------
Sub Test_Tier1OverTier2(assert)
    Dim CrossVar As Long   ' Tier 1: ローカル変数
    CrossVar = 42
    assert.Assert CrossVar, 42, "Tier1 > Tier2: local Dim CrossVar shadows module-level CrossVar"
End Sub
