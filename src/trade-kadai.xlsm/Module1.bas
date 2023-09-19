Attribute VB_Name = "Module1"
    Option Explicit
 
Sub Q1()
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

Sub Q2()

'----------ダイアログボックスで指定のエクセルファイルを開く----------

'エクセルファイルのファイルパスを格納する変数FilePathを宣言する
Dim FilePath As String

'ApplicationオブジェクトのGetOpenFilenameメソッドを使って、[ファイルを開く]ダイアログボックスを表示する
'[ファイルを開く]ダイアログボックスに、どんな拡張子のファイルを表示するかを引数FileFilterで設定する
'選択されたファイルのフルパスをFilePathに格納する
FilePath = Application.GetOpenFilename("output.xlsx")

'GetOpenFilenameメソッドによる[ファイルを開く]ダイアログボックスは選択されたファイルのフルパスを返すだけで、自動的には開かないため、Workbooks.Openを使ってファイルを開く
Workbooks.Open FilePath

'FilePathに代入されているフルパスから、ファイル名を抽出する
'Dir関数は引数に指定したファイルが存在したとき、そのファイル名を返す関数
"output.xlsx" = Dir(FilePath)

'----------値を入力する----------

'前の処理で抽出したファイルのシートとセルを指定して値を入力する




 Range("A3").EntireColumn.AutoFit


Workbooks("output.xlsx").Save
Workbooks("output.xlsx").Close

End Sub
