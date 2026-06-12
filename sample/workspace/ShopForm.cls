VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShopForm"
Option Explicit

' ════════════════════════════════════════════════════════════
' ShopForm クラス — ユーザーフォームのコード部分を模倣
'
' 実際の Excel では UserForm モジュール (ShopForm.frm) に相当。
' UserForm のコードビハインドもクラスモジュールであり、
' ここでは .cls で同等の構造を示す。
'
' 主な構文デモ:
'   - フォームライフサイクル (Class_Initialize / Terminate)
'   - Private WithEvents m_inv As Inventory (在庫監視)
'   - WithEvents ハンドラーによる UI 更新シミュレーション
'   - ボタンクリックハンドラーの模倣 (BtnXxx_Click)
'   - On Error GoTo による個別エラー処理
'   - Private Property / Private Sub / Public Sub の混在
' ════════════════════════════════════════════════════════════

' ─── フォーム内部の Inventory を WithEvents で監視 ────────
Private WithEvents m_inv As Inventory

' ─── フォームの状態変数 ────────────────────────────────────
Private m_title       As String
Private m_selectedId  As Long
Private m_outputBuf   As String
Private m_isLoaded    As Boolean

' ════════════════════════════════════════════════════════════
' フォームライフサイクル
' ════════════════════════════════════════════════════════════

Private Sub Class_Initialize()
    ' UserForm_Initialize 相当: コントロールの初期化
    Set m_inv    = New Inventory
    m_title      = "在庫管理フォーム"
    m_selectedId = -1
    m_outputBuf  = ""
    m_isLoaded   = False
End Sub

Private Sub Class_Terminate()
    ' UserForm_Terminate 相当: リソース解放
    Set m_inv = Nothing
End Sub

' ════════════════════════════════════════════════════════════
' WithEvents ハンドラー (Inventory のイベントを受信)
' ════════════════════════════════════════════════════════════

Private Sub m_inv_ProductAdded(ByVal id As Long, ByVal productName As String)
    ' 商品追加時: リストボックス更新をシミュレート
    Notify "商品追加: [" & id & "] " & productName
End Sub

Private Sub m_inv_PriceAlert(ByVal productName As String, _
                             ByVal oldPrice As Long, ByVal newPrice As Long)
    ' 価格変更アラート: ラベル点滅などをシミュレート
    Notify "【価格変更】" & productName & ": ¥" & oldPrice & " → ¥" & newPrice
End Sub

Private Sub m_inv_StockAlert(ByVal productName As String, ByVal remaining As Long)
    ' 低在庫アラート: 警告ダイアログ表示をシミュレート
    Notify "【在庫警告】" & productName & " 残り " & remaining & " 個"
End Sub

' ════════════════════════════════════════════════════════════
' パブリックメソッド (ボタンクリック・操作の相当)
' ════════════════════════════════════════════════════════════

Public Sub UserForm_Load()
    ' フォームロード時に呼ぶ初期化処理
    ' (実際の UserForm では自動で呼ばれる)
    If m_isLoaded Then Exit Sub

    m_inv.AddProduct "緑茶",       "飲料",   120, 10
    m_inv.AddProduct "クッキー",   "食品",   250, 20
    m_inv.AddProduct "コーヒー",   "飲料",   180, 5
    m_inv.AddProduct "シャンプー", "日用品", 680, 15
    m_inv.AddProduct "チョコ",     "食品",   300, 8

    m_isLoaded = True
    Notify "データ読み込み完了 (" & m_inv.Count & " 件)"
End Sub

Public Sub ListBox_Change(ByVal selectedId As Long)
    ' リストボックスで商品を選択したとき相当
    m_selectedId = selectedId
    Dim p As Product
    Set p = m_inv.GetProduct(m_selectedId)
    If Not p Is Nothing Then
        Notify "選択: " & p.Summary()
    Else
        Notify "選択: ID=" & selectedId & " は存在しません"
    End If
End Sub

Public Sub BtnSell_Click(ByVal qty As Long)
    ' 「販売」ボタンクリック相当
    If Not AssertSelected() Then Exit Sub
    Dim p As Product
    Set p = m_inv.GetProduct(m_selectedId)

    On Error GoTo ErrHandler
    p.Sell qty
    Notify "販売完了: " & p.Name & " ×" & qty & "  残:" & p.Stock
    Exit Sub
ErrHandler:
    Notify "販売エラー: " & Err.Description
End Sub

Public Sub BtnRestock_Click(ByVal qty As Long)
    ' 「補充」ボタンクリック相当
    If Not AssertSelected() Then Exit Sub
    Dim p As Product
    Set p = m_inv.GetProduct(m_selectedId)

    On Error GoTo ErrHandler
    p.Restock qty
    Notify "補充完了: " & p.Name & " +" & qty & "  在庫:" & p.Stock
    Exit Sub
ErrHandler:
    Notify "補充エラー: " & Err.Description
End Sub

Public Sub BtnChangePrice_Click(ByVal newPrice As Long)
    ' 「価格変更」ボタンクリック相当
    If Not AssertSelected() Then Exit Sub
    Dim p As Product
    Set p = m_inv.GetProduct(m_selectedId)

    Dim oldPrice As Long
    oldPrice = p.Price

    On Error GoTo ErrHandler
    p.Price = newPrice
    Notify "価格変更: " & p.Name & "  ¥" & oldPrice & " → ¥" & newPrice
    Exit Sub
ErrHandler:
    Notify "価格エラー: " & Err.Description
End Sub

Public Sub BtnSummary_Click()
    ' 「一覧」ボタンクリック相当
    Notify ""
    Notify "─ 在庫一覧 ─"
    Notify m_inv.AllSummary()
    Notify "在庫評価額合計: ¥" & m_inv.TotalValue()
End Sub

' ════════════════════════════════════════════════════════════
' プロパティ
' ════════════════════════════════════════════════════════════

Public Property Get Title() As String:     Title     = m_title:    End Property
Public Property Get Output() As String:    Output    = m_outputBuf: End Property
Public Property Get AlertLog() As String:  AlertLog  = m_inv.Alerts: End Property

' ════════════════════════════════════════════════════════════
' プライベートヘルパー
' ════════════════════════════════════════════════════════════

Private Function AssertSelected() As Boolean
    If m_selectedId < 0 Then
        Notify "エラー: 商品を選択してください"
        AssertSelected = False
    Else
        AssertSelected = True
    End If
End Function

Private Sub Notify(ByVal msg As String)
    ' フォームへの通知 (ラベルテキスト更新 / Debug出力 をシミュレート)
    Debug.Print msg
    m_outputBuf = m_outputBuf & msg & vbLf
End Sub
