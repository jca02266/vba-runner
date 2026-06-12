Option Explicit

' ════════════════════════════════════════════════════════════
' ShopMain モジュール — 在庫管理デモのエントリポイント
'
' 主な構文デモ:
'   - Public Const / Private Const
'   - Public Enum (列挙型)
'   - Public Type (ユーザー定義型 UDT)
'   - モジュールレベル変数
'   - Public Sub / Private Sub / Private Function
'   - On Error GoTo / Err.Description
'   - 固定長配列 + UDT の組み合わせ
'   - Array() / UBound()
'   - String() 組み込み関数
' ════════════════════════════════════════════════════════════

' ─── パブリック定数 ────────────────────────────────────────
Public Const APP_NAME    As String = "SimpleShop"
Public Const APP_VERSION As String = "1.0.0"

' ─── プライベート定数 ──────────────────────────────────────
Private Const SEP_CHAR  As String = "="
Private Const SEP_WIDTH As Long   = 44

' ─── 列挙型 ────────────────────────────────────────────────
Public Enum SellResult
    srOk       = 0
    srNoStock  = 1
    srNotFound = 2
End Enum

' ─── ユーザー定義型 (UDT) ─────────────────────────────────
Public Type ProductSpec
    PName    As String
    Category As String
    Price    As Long
    Stock    As Long
End Type

' ─── モジュールレベル変数 ──────────────────────────────────
Private g_logBuf  As String
Private g_logLine As Long

' ════════════════════════════════════════════════════════════
' エントリポイント1: クラスを直接使うデモ
' ════════════════════════════════════════════════════════════

Public Sub RunDemo()
    g_logBuf  = ""
    g_logLine = 0

    Log String(SEP_WIDTH, SEP_CHAR)
    Log APP_NAME & " v" & APP_VERSION & "  — クラス直接使用デモ"
    Log String(SEP_WIDTH, SEP_CHAR)

    Dim inv As Inventory
    Set inv = New Inventory

    ' ── 商品登録 ─────────────────────────────────────────
    Log ""
    Log "■ 商品登録"
    LoadSampleData inv

    ' ── 在庫一覧 ─────────────────────────────────────────
    Log ""
    Log "■ 在庫一覧 (" & inv.Count & "件)"
    Log inv.AllSummary()

    ' ── 販売処理（正常・エラー・商品なし） ───────────────
    Log "■ 販売処理"
    Dim res As SellResult

    res = TrySell(inv, 1, 8)   ' → 低在庫アラート発生
    res = TrySell(inv, 2, 3)   ' → 正常
    res = TrySell(inv, 1, 99)  ' → 在庫不足エラー
    res = TrySell(inv, 99, 1)  ' → 商品が存在しない

    ' ── 価格変更（WithEvents → PriceAlert 発火） ─────────
    Log ""
    Log "■ 価格変更"
    AdjustPrice inv, 1, 150

    ' ── 在庫補充 ─────────────────────────────────────────
    Log ""
    Log "■ 在庫補充"
    AdjustStock inv, 1, 20

    ' ── 集計 ─────────────────────────────────────────────
    Log ""
    Log "■ 在庫評価額合計: ¥" & inv.TotalValue()

    ' ── アラートログ ─────────────────────────────────────
    Log ""
    Log "■ アラートログ:"
    Log inv.Alerts

    Log String(SEP_WIDTH, SEP_CHAR)
    Log "総ログ行数: " & g_logLine
    Log String(SEP_WIDTH, SEP_CHAR)

    Set inv = Nothing
End Sub

' ════════════════════════════════════════════════════════════
' エントリポイント2: フォームクラスを経由するデモ
' ════════════════════════════════════════════════════════════

Public Sub RunFormDemo()
    g_logBuf  = ""
    g_logLine = 0

    ' ShopForm は UserForm のコードビハインドに相当するクラス
    Dim frm As ShopForm
    Set frm = New ShopForm

    Log String(SEP_WIDTH, SEP_CHAR)
    Log frm.Title & " — フォームデモ"
    Log String(SEP_WIDTH, SEP_CHAR)

    ' フォームロード（UserForm が表示されるタイミング相当）
    frm.UserForm_Load

    ' リストボックスで商品を選択
    frm.ListBox_Change 1

    ' ボタン操作シミュレーション
    Log ""
    Log "■ ボタン操作"
    frm.BtnSell_Click 8     ' → 低在庫アラート
    frm.BtnSell_Click 99    ' → 在庫不足エラー
    frm.BtnRestock_Click 20
    frm.BtnChangePrice_Click 150

    frm.ListBox_Change 3
    frm.BtnSell_Click 2

    ' 一覧ボタン
    frm.BtnSummary_Click

    ' アラートログ
    Log ""
    Log "■ アラートログ:"
    Log frm.AlertLog

    Log String(SEP_WIDTH, SEP_CHAR)

    Set frm = Nothing
End Sub

' ════════════════════════════════════════════════════════════
' サンプルデータ投入 (UDT 配列を使用)
' ════════════════════════════════════════════════════════════

Private Sub LoadSampleData(ByVal inv As Inventory)
    Dim specs(4) As ProductSpec
    Dim i As Long

    specs(0).PName = "緑茶":       specs(0).Category = "飲料":   specs(0).Price = 120: specs(0).Stock = 10
    specs(1).PName = "クッキー":   specs(1).Category = "食品":   specs(1).Price = 250: specs(1).Stock = 20
    specs(2).PName = "コーヒー":   specs(2).Category = "飲料":   specs(2).Price = 180: specs(2).Stock = 5
    specs(3).PName = "シャンプー": specs(3).Category = "日用品": specs(3).Price = 680: specs(3).Stock = 15
    specs(4).PName = "チョコ":     specs(4).Category = "食品":   specs(4).Price = 300: specs(4).Stock = 8

    For i = 0 To UBound(specs)
        Dim p As Product
        Set p = inv.AddProduct(specs(i).PName, specs(i).Category, _
                               specs(i).Price, specs(i).Stock)
        Log "  登録: " & p.Summary()
    Next i
End Sub

' ════════════════════════════════════════════════════════════
' 販売処理 (On Error GoTo によるエラーハンドリング)
' ════════════════════════════════════════════════════════════

Private Function TrySell(ByVal inv As Inventory, ByVal id As Long, _
                          ByVal qty As Long) As SellResult
    Dim p As Product
    Set p = inv.GetProduct(id)

    If p Is Nothing Then
        Log "  商品 ID=" & id & " が見つかりません"
        TrySell = srNotFound
        Exit Function
    End If

    On Error GoTo ErrHandler
    p.Sell qty
    Log "  販売: " & p.Name & " ×" & qty & "  残:" & p.Stock
    TrySell = srOk
    Exit Function

ErrHandler:
    Log "  [エラー] " & Err.Description
    TrySell = srNoStock
End Function

' ════════════════════════════════════════════════════════════
' 価格・在庫変更
' ════════════════════════════════════════════════════════════

Private Sub AdjustPrice(ByVal inv As Inventory, ByVal id As Long, _
                         ByVal newPrice As Long)
    Dim p As Product
    Set p = inv.GetProduct(id)
    If p Is Nothing Then Exit Sub

    Dim oldPrice As Long
    oldPrice = p.Price

    On Error GoTo ErrHandler
    p.Price = newPrice
    Log "  " & p.Name & ": ¥" & oldPrice & " → ¥" & newPrice
    Exit Sub
ErrHandler:
    Log "  [価格エラー] " & Err.Description
End Sub

Private Sub AdjustStock(ByVal inv As Inventory, ByVal id As Long, _
                         ByVal qty As Long)
    Dim p As Product
    Set p = inv.GetProduct(id)
    If p Is Nothing Then Exit Sub

    On Error GoTo ErrHandler
    p.Restock qty
    Log "  " & p.Name & " 補充 +" & qty & "  在庫:" & p.Stock
    Exit Sub
ErrHandler:
    Log "  [補充エラー] " & Err.Description
End Sub

' ════════════════════════════════════════════════════════════
' ユーティリティ
' ════════════════════════════════════════════════════════════

Public Sub Log(ByVal msg As String)
    Debug.Print msg
    g_logBuf  = g_logBuf & msg & vbLf
    g_logLine = g_logLine + 1
End Sub
