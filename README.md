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