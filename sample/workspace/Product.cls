VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Product"
Option Explicit

' ════════════════════════════════════════════════════════════
' Product クラス — 商品の名前・価格・在庫数を保持する
'
' 主な構文デモ:
'   - Private Const (クラス内定数)
'   - Private メンバー変数 (m_ プレフィックス)
'   - Public Event / RaiseEvent
'   - Class_Initialize / Class_Terminate
'   - Property Get / Let (引数検証付き)
'   - Public メソッド / Private ヘルパー
' ════════════════════════════════════════════════════════════

' ─── プライベート定数 ─────────────────────────────────────
Private Const LOW_STOCK_LEVEL As Long   = 5
Private Const MAX_STOCK       As Long   = 9999
Private Const MIN_PRICE       As Long   = 1
Private Const MAX_PRICE       As Long   = 1000000

' ─── プライベートメンバー変数 ─────────────────────────────
Private m_id       As Long
Private m_name     As String
Private m_category As String
Private m_price    As Long
Private m_stock    As Long

' ─── イベント宣言 ─────────────────────────────────────────
Public Event PriceChanged(ByVal oldPrice As Long, ByVal newPrice As Long)
Public Event StockChanged(ByVal oldStock As Long, ByVal newStock As Long)
Public Event StockLow(ByVal currentStock As Long)

' ════════════════════════════════════════════════════════════
' ライフサイクル
' ════════════════════════════════════════════════════════════

Private Sub Class_Initialize()
    m_id       = 0
    m_price    = MIN_PRICE
    m_stock    = 0
    m_name     = "(未設定)"
    m_category = ""
End Sub

Private Sub Class_Terminate()
    Debug.Print "  [~Product] " & m_name & " 解放"
End Sub

' ════════════════════════════════════════════════════════════
' プロパティ
' ════════════════════════════════════════════════════════════

Public Property Get Id() As Long:       Id = m_id:             End Property
Public Property Let Id(ByVal v As Long): m_id = v:             End Property

Public Property Get Name() As String:   Name = m_name:         End Property
Public Property Let Name(ByVal v As String): m_name = Trim(v): End Property

Public Property Get Category() As String:   Category = m_category:   End Property
Public Property Let Category(ByVal v As String): m_category = v:     End Property

Public Property Get Price() As Long
    Price = m_price
End Property

Public Property Let Price(ByVal v As Long)
    If Not ValidateRange(v, MIN_PRICE, MAX_PRICE) Then
        Err.Raise 1001, "Product.Price", "価格が範囲外です: " & v
    End If
    Dim old As Long
    old = m_price
    m_price = v
    If old <> v Then RaiseEvent PriceChanged(old, v)
End Property

Public Property Get Stock() As Long
    Stock = m_stock
End Property

Public Property Let Stock(ByVal v As Long)
    If Not ValidateRange(v, 0, MAX_STOCK) Then
        Err.Raise 1002, "Product.Stock", "在庫数が範囲外です: " & v
    End If
    Dim old As Long
    old = m_stock
    m_stock = v
    If old <> v Then
        RaiseEvent StockChanged(old, v)
        If IsLowStock() Then RaiseEvent StockLow(m_stock)
    End If
End Property

' ════════════════════════════════════════════════════════════
' パブリックメソッド
' ════════════════════════════════════════════════════════════

Public Sub Sell(ByVal qty As Long)
    If qty > m_stock Then
        Err.Raise 1003, "Product.Sell", _
            "在庫不足 (要求:" & qty & " 在庫:" & m_stock & ")"
    End If
    Me.Stock = m_stock - qty
End Sub

Public Sub Restock(ByVal qty As Long)
    Me.Stock = m_stock + qty
End Sub

Public Function Summary() As String
    Summary = "[" & m_id & "] " & m_name & _
              " (" & m_category & ")" & _
              "  ¥" & m_price & _
              "  在庫:" & m_stock
End Function

' ════════════════════════════════════════════════════════════
' プライベートヘルパー
' ════════════════════════════════════════════════════════════

Private Function IsLowStock() As Boolean
    IsLowStock = (m_stock <= LOW_STOCK_LEVEL)
End Function

Private Function ValidateRange(ByVal v As Long, ByVal minV As Long, _
                                ByVal maxV As Long) As Boolean
    ValidateRange = (v >= minV And v <= maxV)
End Function
