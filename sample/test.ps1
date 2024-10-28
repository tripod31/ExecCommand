$oData = New-Object -ComObject ExecCommand.SharedData
$s = ""
// COMサーバー内で例外をcatchしているが、なぜか例外をcatchしない。なのでpowershell側で例外をcatchする
try {
    $s=$oData.GetData()
}catch{
    $s="GetDataでエラー"
}

#標準出力に読み取った共有データを出力
#-NoNewLine:改行コードを出力しない
Write-Host -NoNewline $s
#共有データをセット
$oData.SetData("_RETURN_")
