Option Explicit

' ════════════════════════════════════════════════════════════
' Inventory クラス — 商品コレクションを管理する
'
' 主な構文デモ:
'   - Private WithEvents (Product のイベントを受信)
'   - WithEvents イベントハンドラー (m_watched_XXX)
'   - Public Event / RaiseEvent (外向きイベントの中継)
'   - Scripting.Dictionary によるコレクション管理
'   - For Each ループ
'   - Class_Initialize / Class_Terminate
' ════════════════════════════════════════════════════════════

' ─── プライベート変数 ──────────────────────────────────────
Private m_items    As Object   ' Scripting.Dictionary: key=id(String), value=Product
Private m_nextId   As Long
Private m_alerts   As String   ' アラート蓄積バッファ

' ─── WithEvents — 最後に追加した商品の変化を監視 ──────────
Private WithEvents m_watched As Product

' ─── 外向きイベント ────────────────────────────────────────
Public Event ProductAdded(ByVal id As Long, ByVal productName As String)
Public Event PriceAlert(ByVal productName As String, _
                        ByVal oldPrice As Long, ByVal newPrice As Long)
Public Event StockAlert(ByVal productName As String, ByVal remaining As Long)

' ════════════════════════════════════════════════════════════
' ライフサイクル
' ════════════════════════════════════════════════════════════

Private Sub Class_Initialize()
    Set m_items = CreateObject("Scripting.Dictionary")
    m_nextId = 1
    m_alerts = ""
End Sub

Private Sub Class_Terminate()
    Set m_items   = Nothing
    Set m_watched = Nothing
End Sub

' ════════════════════════════════════════════════════════════
' WithEvents ハンドラー
' ════════════════════════════════════════════════════════════

Private Sub m_watched_PriceChanged(ByVal oldPrice As Long, ByVal newPrice As Long)
    Dim msg As String
    msg = "価格変更: " & m_watched.Name & "  ¥" & oldPrice & " → ¥" & newPrice
    AppendAlert msg
    RaiseEvent PriceAlert(m_watched.Name, oldPrice, newPrice)
End Sub

Private Sub m_watched_StockChanged(ByVal oldStock As Long, ByVal newStock As Long)
    Debug.Print "  [Inventory] 在庫更新: " & m_watched.Name & _
                "  " & oldStock & " → " & newStock
End Sub

Private Sub m_watched_StockLow(ByVal currentStock As Long)
    Dim msg As String
    msg = "低在庫警告: " & m_watched.Name & "  残り " & currentStock & " 個"
    AppendAlert msg
    RaiseEvent StockAlert(m_watched.Name, currentStock)
End Sub

' ════════════════════════════════════════════════════════════
' プロパティ
' ════════════════════════════════════════════════════════════

Public Property Get Count() As Long:    Count  = m_items.Count: End Property
Public Property Get Alerts() As String: Alerts = m_alerts:      End Property

' ════════════════════════════════════════════════════════════
' パブリックメソッド
' ════════════════════════════════════════════════════════════

Public Function AddProduct(ByVal pName As String, ByVal category As String, _
                           ByVal price As Long, ByVal stock As Long) As Product
    Dim p As Product
    Set p = New Product
    p.Id       = m_nextId
    p.Name     = pName
    p.Category = category
    p.Price    = price
    p.Stock    = stock

    m_items.Add CStr(m_nextId), p
    Set m_watched = p       ' WithEvents 監視対象を更新
    m_nextId = m_nextId + 1

    RaiseEvent ProductAdded(p.Id, p.Name)
    Set AddProduct = p
End Function

Public Function GetProduct(ByVal id As Long) As Product
    Dim key As String
    key = CStr(id)
    If m_items.exists(key) Then
        Set GetProduct = m_items(key)
    Else
        Set GetProduct = Nothing
    End If
End Function

Public Function AllSummary() As String
    Dim result As String
    Dim key As Variant
    For Each key In m_items.Keys
        Dim p As Product
        Set p = m_items(key)
        result = result & p.Summary() & vbLf
    Next key
    AllSummary = result
End Function

Public Function TotalValue() As Long
    Dim total As Long
    Dim key As Variant
    For Each key In m_items.Keys
        Dim p As Product
        Set p = m_items(key)
        total = total + p.Price * p.Stock
    Next key
    TotalValue = total
End Function

' ════════════════════════════════════════════════════════════
' プライベートメソッド
' ════════════════════════════════════════════════════════════

Private Sub AppendAlert(ByVal msg As String)
    m_alerts = m_alerts & msg & vbLf
End Sub
