$oData = New-Object -ComObject ExecCommand.SharedData
#標準出力に読み取ったデータを出力
#-NoNewLine:改行コードを出力しない
Write-Host -NoNewline $oData.GetData()
#データをセット
$oData.SetData("_RETURN_")
