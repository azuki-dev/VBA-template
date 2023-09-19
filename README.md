## command

エクセルファイルからマクロを抽出
```
cscript vbac.wsf decombine
```

記述したマクロをエクセルファイルに反映

1. エクスプローラーでこのディレクトリを開く
2. shift + 右クリック
3. Powershellウィンドウをここで開く
4. Powershell上で下記コマンドを実行
```
powershell.exe -ExecutionPolicy Bypass -File .\reload.ps1 <ファイル名>

例)
powershell.exe -ExecutionPolicy Bypass -File .\reload.ps1 Sample.xlsm
```
powershell.exe -ExecutionPolicy Bypass -File .\reload.ps1　trade-kadai.xlsm
```

[
    [29,78,972],
        [29,78,972],
    [29,78,972],

]
[
    [ "インダストリアルテクノロジーソリューションズ","ターボグラインダーアタッチメント",78,1900000],
    [ "インダストリアルテクノロジーソリューションズ","ターボグラインダーアタッチメント",78,1900000],
    [ "インダストリアルテクノロジーソリューションズ","ターボグラインダーアタッチメント",78,1900000]
]


```"トレードカダイの取引履歴シートからID->取引先商品数量を読み取る"
"商品IDを元に取引先マスタシートのA欄を参照してその横のB欄をコピー"
"取引先IDを元に製品マスタシートのA欄を参照してその横のB欄とC欄をコピー"
"コピー出来たらそれをoutputファイルの取引先をA欄に代入、商品をB欄に代入、Cは数量と価格(C欄でコピーしたもの)をかけて代入"