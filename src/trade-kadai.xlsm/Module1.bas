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

'----------シートを検索する----------
    Dim dicResult As Object
    Dim findStr As String
    Dim ws As Worksheet
    Dim resultRange As Range
    Dim address As String
    Dim arrAddress As Variant
    Dim resultRangeStr As String
    Dim key As Variant
    
    Set dicResult = CreateObject("Scripting.Dictionary")
    
    '検索対象の文字列を指定
    findStr = "エレクトロニクス"
    '検索対象のシートを指定
    Set ws = Worksheets("取引先マスタ")
    
    '検索を実行(1回目)
    Set resultRange = ws.Cells.Find(What:=findStr, LookIn:=xlValues, LookAt:=xlPart)
    
    '検索を実行(最後まで繰り返し)
    Do While Not (resultRange Is Nothing)
        '見つかったセルのアドレスを取得(例:C$3など)
        address = resultRange.address(RowAbsolute:=True, ColumnAbsolute:=False)
        '見つかったセルを取得(例:C3など)
        resultRangeStr = Split(address, "$")(0) & resultRange.Row
        '見つかったセルが既にDictionaryに設定済みの場合はLoopを抜ける
        If dicResult.Exists(resultRangeStr) Then
            Exit Do
        End If
        '見つかったセルの情報をDictionaryへ設定
        dicResult.Add resultRangeStr, resultRange.Value
        '次を検索
        Set resultRange = ws.Cells.FindNext(After:=resultRange)
    Loop
    
    '検索結果(=Dictionaryの内容)をイミディエイトウィンドウへ出力
    If dicResult.Count <> 0 Then
        For Each key In dicResult.Keys
            
        Next
    Else
        MsgBox "指定した文字列は存在しませんでした。"
    End If



'前の処理で抽出したファイルのシートとセルを指定して値を入力する
Dim aaa As String
Sheets("sheet1").Select
Sheets("sheet1").Name = "エレクトロニクス"
Workbooks("output.xlsx").Worksheets("エレクトロニクス").Range("A1").Value = "会社名"
Workbooks("output.xlsx").Worksheets("エレクトロニクス").Range("A3").Value = "注文商品"
Workbooks("output.xlsx").Worksheets("エレクトロニクス").Range("B3").Value = "金額"
Workbooks("output.xlsx").Worksheets("エレクトロニクス").Range("C3").Value = "数量"
Workbooks("output.xlsx").Worksheets("エレクトロニクス").Range("D3").Value = "合計"

'シート2'
 With Sheets.Add(After:=Sheets(Sheets.Count))
    .Name = "プライムエンジニアリング"
End With
Workbooks("output.xlsx").Worksheets("プライムエンジニアリング").Range("A1").Value = "会社名"
Workbooks("output.xlsx").Worksheets("プライムエンジニアリング").Range("A3").Value = "注文商品"
Workbooks("output.xlsx").Worksheets("プライムエンジニアリング").Range("B3").Value = "金額"
Workbooks("output.xlsx").Worksheets("プライムエンジニアリング").Range("C3").Value = "数量"
Workbooks("output.xlsx").Worksheets("プライムエンジニアリング").Range("D3").Value = "合計"



 Range("A3").EntireColumn.AutoFit 


Workbooks("output.xlsx").Save
Workbooks("output.xlsx").Close

End Sub
