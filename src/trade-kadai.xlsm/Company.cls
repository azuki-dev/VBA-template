VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Company"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Option Explicit

' Companyクラス
' 各会社を表すクラスです。
' companyId、companyName、tradesプロパティを持っています。
' tradesプロパティはTradeオブジェクトのコレクションです。

Private pCompanyId As Integer     ' 会社のID
Private pCompanyName As String    ' 会社の名前
Private pTrades As Collection     ' 会社の取引のコレクション

' CompanyIdプロパティのゲッターとセッター
Public Property Get CompanyId() As Integer
    CompanyId = pCompanyId
End Property
Public Property Let CompanyId(Value As Integer)
    pCompanyId = Value
End Property

' CompanyNameプロパティのゲッターとセッター
Public Property Get CompanyName() As String
    CompanyName = pCompanyName
End Property
Public Property Let CompanyName(Value As String)
    pCompanyName = Value
End Property

' Tradesプロパティのゲッターとセッター
Public Property Get Trades() As Collection
    Set Trades = pTrades
End Property
Public Property Set Trades(Value As Collection)
    Set pTrades = Value
End Property

' Companyクラスの初期化
' Tradeオブジェクトのコレクションを新規作成します。
Private Sub Class_Initialize()
    Set pTrades = New Collection
End Sub

' Tradeオブジェクトを追加するメソッド
' 引数で与えられたTradeオブジェクトをTradesコレクションに追加します。
Public Sub AddTrade(objTrade As Trade)
    pTrades.Add objTrade
End Sub
