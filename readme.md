ExecCommand
=====

これは何か
-----
COMオブジェクトです。  
#### DOS窓を出さずにCUIコマンドを起動する  
Excel VBAからperlを起動して標準出力を得るためにWindowsScriptingHostのCOMオブジェクト(wscript.shell)を使用していましたが、これだと起動時にDOS窓が表示されるため、うっとおしい。  
そのため、C#でコマンドを起動するCOMオブジェクトを作成しました。  
コマンドの終了を待ち、標準出力・標準エラーを取得します。DOS窓は表示しません。

#### プロセス間通信
親プロセスから子プロセスに(ExcelからPerlなど）、コマンドラインで渡すのが難しい大きな文字列を渡すためにプロセス間通信をサポートしました。  
プロセスが親子関係である必要はありません。

開発・動作確認環境
-----
Windows10  
Microsoft VisualStudio Express 2015forWindowsDesktop  
.NET Framework 4.5(4.0向けにビルド)  
Excel2003

インストール
-----
#### ファイルをダウンロードして展開
execcommand.zipをクリック→Download

#### DLLのレジストリへの登録・登録解除
register.exeを実行する。DLLがregister.exeと違うディレクトリにある場合、パスを指定する。registerボタンをクリック。エラーが出ても、「型が正常に登録されました。」と表示されればOK。登録解除はunregisterボタン。  
64bitOSで32bitプログラムからDLLを使用するには、register32.exeを使用してDLLを登録する。

変更履歴
-----
#### 2018/03/17
64bitOSで32bitプログラムからDLLを使用するには、register32.exeを使用してDLLを登録する。

VBAからの使用例
-----
#### DOS窓を出さずにCUIコマンドを起動する
pingコマンドの出力を表示します。

	Sub test()
		Dim obj As Object
		Dim iRet As Integer

		Set obj = CreateObject("ExecCommand.ExecCommand")
		obj.exeFile = "ping.exe"
		obj.arg = "localhost -n 1"
		iRet = obj.Exec

		If iRet <> 0 Then
			'コマンド起動時エラー
			MsgBox obj.Err
			Exit Sub
		End If

		MsgBox "標準出力->" & obj.stdout & "<-" & vbCrLf & "標準エラー->" & obj.stderr & "<-"

	End Sub

#### プロセス間通信

	Sub test_remote()
		Dim obj As Object
		Dim iRet As Integer
		Dim msg As String
		Dim sPath As String
		Dim sDir As String
		Dim oServer
		Dim oExec
    
		sDir = get_curdir
		Set oServer = CreateObject("ExecCommand.RemoteServer")
    
		iRet = oServer.init
			If iRet <> 0 Then
			MsgBox oServer.Err
			Exit Sub
		End If
        
		oServer.Data = "ARGUMENT"
    
		'クライアント起動
		Set oExec = CreateObject("ExecCommand.ExecCommand")
		oExec.ExeFile = "perl"
		oExec.Arg = sDir & "\remote\test_client.pl"
		iRet = oExec.Exec
		If iRet <> 0 Then
			'コマンド起動時エラー
			MsgBox oExec.Err
			Exit Sub
		End If
    
		If oExec.stdout = "ARGUMENT" Then
			msg = "Argument was passed successfully."
		Else
			msg = "Argument was'nt passed."
		End If
    
		msg = msg & Chr(13)
		If oServer.Data = "RETURN" Then
			msg = msg & "Return value was passed successfully."
		Else
			msg = msg & "Return value was'nt passed."
		End If

		iRet = oServer.Close
		If iRet <> 0 Then
			MsgBox oServer.Err
			Exit Sub
		End If
		Set oServer = Nothing

		MsgBox msg
	End Sub

