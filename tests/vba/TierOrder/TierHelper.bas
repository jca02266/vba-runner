Option Explicit

' Tier 3（クロスモジュール Public）として公開される変数
' TierOrderTest.bas から Tier 3 として見える

Public CrossVar As Long    ' Tier 3: 他モジュールから Tier 3 として見える
Public SharedName As Long  ' Tier 3: Tier 4 組み込みと同名にはしない（汎用名）
