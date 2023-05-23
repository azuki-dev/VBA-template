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
FileName = Dir(FilePath)

'----------値を入力する----------

'前の処理で抽出したファイルのシートとセルを指定して値を入力する
Dim aaa As String
Sheets("sheet1").Select
Sheets("sheet1").Name = "エレクトロニクス"
Workbooks(FileName).Worksheets("エレクトロニクス").Range("A1").Value = "会社名"
Workbooks(FileName).Worksheets("エレクトロニクス").Range("A3").Value = "注文商品"
Workbooks(FileName).Worksheets("エレクトロニクス").Range("A4").Value = "メガスパンネジ"
Workbooks(FileName).Worksheets("エレクトロニクス").Range("A5").Value = "ハイパーロックボルト"
Workbooks(FileName).Worksheets("エレクトロニクス").Range("B1").Value = "エレクトロニクス"
Workbooks(FileName).Worksheets("エレクトロニクス").Range("B3").Value = "金額"
Workbooks(FileName).Worksheets("エレクトロニクス").Range("B4").Value = "9300"
Workbooks(FileName).Worksheets("エレクトロニクス").Range("B5").Value = "1700"
Workbooks(FileName).Worksheets("エレクトロニクス").Range("C3").Value = "数量"
Workbooks(FileName).Worksheets("エレクトロニクス").Range("C4").Value = "2"
Workbooks(FileName).Worksheets("エレクトロニクス").Range("C5").Value = "1"
Workbooks(FileName).Worksheets("エレクトロニクス").Range("D3").Value = "合計"
Workbooks(FileName).Worksheets("エレクトロニクス").Range("D4").Value = "=SUM(B4*C4)"
Workbooks(FileName).Worksheets("エレクトロニクス").Range("D5").Value = "=SUM(B5*C5)"
Workbooks(FileName).Worksheets("エレクトロニクス").Range("D6").Value = "=SUM(D4+D5)"
'シート2'
 With Sheets.Add(After:=Sheets(Sheets.Count))
Worksheets(2).Name = "プライムエンジニアリング"
End With
Workbooks(FileName).Worksheets("プライムエンジニアリング").Range("A1").Value = "会社名"
Workbooks(FileName).Worksheets("プライムエンジニアリング").Range("A3").Value = "注文商品"
Workbooks(FileName).Worksheets("プライムエンジニアリング").Range("A4").Value = "フレキシブルシャフトレンチ"
Workbooks(FileName).Worksheets("プライムエンジニアリング").Range("A5").Value = "メガパワーグラインダー"
Workbooks(FileName).Worksheets("プライムエンジニアリング").Range("A6").Value = "エクストラロングリーチレンチ"
Workbooks(FileName).Worksheets("プライムエンジニアリング").Range("B1").Value = "プライムエンジニアリング"
Workbooks(FileName).Worksheets("プライムエンジニアリング").Range("B3").Value = "金額"
Workbooks(FileName).Worksheets("プライムエンジニアリング").Range("B4").Value = "480"
Workbooks(FileName).Worksheets("プライムエンジニアリング").Range("B5").Value = "6100"
Workbooks(FileName).Worksheets("プライムエンジニアリング").Range("B6").Value = "8000"
Workbooks(FileName).Worksheets("プライムエンジニアリン").Range("C3").Value = "数量"
Workbooks(FileName).Worksheets("プライムエンジニアリング").Range("C4").Value = "10"
Workbooks(FileName).Worksheets("プライムエンジニアリング").Range("C5").Value = "1"
Workbooks(FileName).Worksheets("プライムエンジニアリング").Range("C6").Value = "3"
Workbooks(FileName).Worksheets("プライムエンジニアリング").Range("D3").Value = "合計"
Workbooks(FileName).Worksheets("プライムエンジニアリング").Range("D4").Value = "=SUM(B4*C4)"
Workbooks(FileName).Worksheets("プライムエンジニアリング").Range("D5").Value = "=SUM(B5*C5)"
Workbooks(FileName).Worksheets("プライムエンジニアリング").Range("D6").Value = "=SUM(B6*C6)"
Workbooks(FileName).Worksheets("プライムエンジニアリング").Range("D7").Value = "=SUM(D4+D5+D6)"










'----------列を自動調節-----------
 Range("A3").EntireColumn.AutoFit    '---(2)A列の一番長いセルのセル幅に自動調整
'----------上書き保存して閉じる----------

Workbooks(FileName).Save
Workbooks(FileName).Close

End Sub
