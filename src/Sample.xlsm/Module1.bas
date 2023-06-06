Attribute VB_Name = "Module1"
Sub ・ス・ス1()
    Range("A1") = "Hello World!"
End Sub
Sub ・ス・ス2()
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
Sub ・ス・ス3()
    Range("B1").Value = "=SUM(A1:A10)"
End Sub
Sub ・ス・ス4()
    Range("B2").Value = "=AVERAGE(A1:A10)"
End Sub
Sub ・ス・ス5()
  Range("A1:A11").Sort _
        Key1:=Range("A1"), Order1:=xlAscending, _
        Header:=xlNo
End Sub
Sub ・ス・ス6()
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
Sub ・ス・ス7()
    Dim Names(1 To 5) As String
    
    Worksheets("Sheet1").Range("A1") = "�ｿｽﾛ登"
    Worksheets("Sheet1").Range("A2") = "�ｿｽ�ｿｽ�ｿｽ�ｿｽ"
    Worksheets("Sheet1").Range("A3") = "�ｿｽV�ｿｽX�ｿｽ�ｿｽ"
    Worksheets("Sheet1").Range("A4") = "�ｿｽﾘ奇ｿｽ"
    Worksheets("Sheet1").Range("A5") = "�ｿｽF�ｿｽ�ｿｽ"
    
    Names(1) = Sheets("Sheet1").Range("A1").Value
    Names(2) = Sheets("Sheet1").Range("A2").Value
    Names(3) = Sheets("Sheet1").Range("A3").Value
    Names(4) = Sheets("Sheet1").Range("A4").Value
    Names(5) = Sheets("Sheet1").Range("A5").Value
    
    Dim RandNum As Integer
    Dim Temp As String
    
    For i = 1 To 5
        RandNum = Int((5 - i + 1) * Rnd + i) ' �ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽ_�ｿｽ�ｿｽ�ｿｽﾈ撰ｿｽ�ｿｽ�ｿｽ�ｿｽ謫ｾ
        Temp = Names(i) ' �ｿｽ�ｿｽ�ｿｽO�ｿｽ�ｿｽ�ｿｽ齊橸ｿｽI�ｿｽﾉ保托ｿｽ
        Names(i) = Names(RandNum) ' �ｿｽ�ｿｽ�ｿｽO�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽﾖゑｿｽ�ｿｽ�ｿｽ
        Names(RandNum) = Temp ' �ｿｽ齊橸ｿｽI�ｿｽﾉ保托ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽO�ｿｽ�ｿｽ�ｿｽ�ｿｽ
    Next i

    Sheets("Sheet2").Range("A1").Value = Names(1)
    Sheets("Sheet2").Range("A2").Value = Names(2)
    Sheets("Sheet2").Range("A3").Value = Names(3)
    Sheets("Sheet2").Range("A4").Value = Names(4)
    Sheets("Sheet2").Range("A5").Value = Names(5)
End Sub
Sub ・ス・ス8()
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
Sub ・ス・ス9()
    Range("A1") = "1/1"
    Range("A2") = "4/23"
    Range("A3").Value = "=DAYS(A2,A1)"
End Sub
Sub ・ス・ス10()
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
        ' �ｿｽV�ｿｽ[�ｿｽg1�ｿｽ�ｿｽA�ｿｽ�から数�ｿｽl�ｿｽ�ｿｽ�ｿｽ謫ｾ�ｿｽ�ｿｽ�ｿｽ�ｿｽ10�ｿｽ{�ｿｽ�ｿｽ�ｿｽ�ｿｽ
        Sheets("Sheet1").Activate
        Dim val As Double
        val = Range("A" & i).Value * 10
        
        ' �ｿｽV�ｿｽ[�ｿｽg2�ｿｽ�ｿｽA�ｿｽ�ｿｽﾉ鯉ｿｽ�ｿｽﾊゑｿｽ\�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽ�ｿｽ
        Sheets("Sheet2").Activate
        Range("A" & i).Value = val
    Next i

End Sub
