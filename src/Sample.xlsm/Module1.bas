Attribute VB_Name = "Module1"
Sub 問1()
    Range("A1") = "Hello World!"
End Sub
Sub 問2()
    Range("A1") = 1
    Range("A2") = 2
    Range("A3") = 3
    Range("A4") = 4
    Range("A5") = 5
    Range("A6") = 6
    Range("A7") = 7
    Range("A8") = 8
    Range("A9") = 9
    Range("A10") = 10
End Sub
Sub 問3()
    Range("B1").Value = "=SUM(A1:A10)"
End Sub
Sub 問4()
    Range("B2").Value = "=AVERAGE(A1:A10)"
End Sub
Sub 問5()
  Range("A1:A11").Sort _
        Key1:=Range("A1"), Order1:=xlAscending, _
        Header:=xlNo
End Sub
Sub 問6()
    Dim i As Integer
    Dim j As Integer
    
    j = 1
    
    For i = 1 To 10
        If Worksheets("Sheet1").Cells(i, 1).Value Mod 2 = 0 Then
            Worksheets("Sheet2").Cells(j, 1).Value = Worksheets("Sheet1").Cells(i, 1).Value
            j = j + 1
        End If
    Next i
    
End Sub
Sub 問7()
    Dim Names(1 To 5) As String
    
    Worksheets("Sheet1").Range("A1") = "保登"
    Worksheets("Sheet1").Range("A2") = "香風"
    Worksheets("Sheet1").Range("A3") = "天々座"
    Worksheets("Sheet1").Range("A4") = "切間"
    Worksheets("Sheet1").Range("A5") = "宇治"
    
    Names(1) = Sheets("Sheet1").Range("A1").Value
    Names(2) = Sheets("Sheet1").Range("A2").Value
    Names(3) = Sheets("Sheet1").Range("A3").Value
    Names(4) = Sheets("Sheet1").Range("A4").Value
    Names(5) = Sheets("Sheet1").Range("A5").Value
    
    Dim RandNum As Integer
    Dim Temp As String
    
    For i = 1 To 5
        RandNum = Int((5 - i + 1) * Rnd + i) ' ランダムな数を取得
        Temp = Names(i) ' 名前を一時的に保存
        Names(i) = Names(RandNum) ' 名前を入れ替える
        Names(RandNum) = Temp ' 一時的に保存した名前を代入
    Next i

    Sheets("Sheet2").Range("A1").Value = Names(1)
    Sheets("Sheet2").Range("A2").Value = Names(2)
    Sheets("Sheet2").Range("A3").Value = Names(3)
    Sheets("Sheet2").Range("A4").Value = Names(4)
    Sheets("Sheet2").Range("A5").Value = Names(5)
End Sub
Sub 問8()
    Range("A1") = 4513531
    Range("A2") = 4541554
    Range("A3") = 4524513
    Range("A4") = 5344513
    Range("A5") = 4863525
    Range("A6") = 3513515
    Range("A7") = 3512515
    Range("A8") = 5323515
    Range("A9") = 3512555
    Range("A10") = 4351254
    
    Range("B1").Value = "=MAX(A1:A10)"
    Range("B2").Value = "=MIN(A1:A10)"
End Sub
Sub 問9()
    Range("A1") = "1/1"
    Range("A2") = "4/23"
    Range("A3").Value = "=DAYS(A2,A1)"
End Sub
Sub 問10()
    Dim i As Integer
    Range("A1") = 1
    Range("A2") = 2
    Range("A3") = 3
    Range("A4") = 4
    Range("A5") = 5
    Range("A6") = 6
    Range("A7") = 7
    Range("A8") = 8
    Range("A9") = 9
    Range("A10") = 10
    
    
    For i = 1 To 10
        ' シート1のA列から数値を取得して10倍する
        Sheets("Sheet1").Activate
        Dim val As Double
        val = Range("A" & i).Value * 10
        
        ' シート2のA列に結果を表示する
        Sheets("Sheet2").Activate
        Range("A" & i).Value = val
    Next i

End Sub
