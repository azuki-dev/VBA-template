VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Trade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Sub Class_Initialize()


Option Explicit

' Tradeクラス
' 各取引を表すクラスです。
' id、name、amountプロパティを持っています。

Private pId As Integer       ' 取引のID
Private pName As String       ' 取引の名前
Private pAmount As Integer    ' 取引の量

' Idプロパティのゲッターとセッター
Public Property Get Id() As Integer
    Id = pId
End Property
Public Property Let Id(Value As Integer)
    pId = Value
End Property

' Nameプロパティのゲッターとセッター
Public Property Get Name() As String
    Name = pName
End Property
Public Property Let Name(Value As String)
    pName = Value
End Property

' Amountプロパティのゲッターとセッター
Public Property Get Amount() As Integer
    Amount = pAmount
End Property
Public Property Let Amount(Value As Integer)
    pAmount = Value
End Property

End Sub


