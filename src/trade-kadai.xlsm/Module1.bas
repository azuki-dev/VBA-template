Attribute VB_Name = "Module1"
    Option Explicit
 
Sub createExcelFileTest()
    Dim newBookName As String
    Dim newBookPath As String
    Dim newBook As Workbook
    
    '新しいファイルの名前を指定
    newBookName = "output.xlsx"
    
    '新しいファイルのフルパスを設定
    newBookPath = ThisWorkbook.Path & "\" & newBookName
    
    '指定したパスにファイルが作成済でないかを確認。
    If Dir(newBookPath) = "" Then
        '新しいファイルを作成
        Set newBook = Workbooks.Add
        
        '新しいファイルをVBAを実行したファイルと同じフォルダ保存
        newBook.SaveAs newBookPath
    
    Else
        '既に同名のファイルが存在する場合はメッセージを表示
        MsgBox "既に" & newBookName & "というファイルは存在します。"
    
    End If
End Sub 


End Sub
