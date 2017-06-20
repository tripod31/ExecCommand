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
親プロセスから子プロセスに(ExcelからPerlなど）、コマンドラインで渡すのが難しい大きな文字列を渡すためにプロセス間通信をサポートしました  
プロセスが親子関係である必要はありません。

開発・動作確認環境
-----
Windows10  
Microsoft VisualStudio Express 2015forWindowsDesktop  
.NET Framework 4.5(4.0向けにビルド)  
Excel2007

インストール
-----
#### ファイルをダウンロードして展開
「Download ZIP」のリンクからダウンロード

#### DLLのレジストリへの登録
regist_dll.batを実行する。WindowsVista以降の場合、管理者で実行。  
バッチファイル中のregasm.exeのパスを修正する必要があるかも。  

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

プロセス間通信

	Sub test_remote()
	    Dim obj As Object
	    Dim iRet As Integer
	    Dim msg As String

	    Set obj = CreateObject("ExecCommand.ExecCommand")

	    obj.InitRemoteServer
	    obj.ServerData = "[server]hello!"

	    'クライアント起動
	    obj.ExeFile = "perl"
	    obj.Arg = "test_client.pl"
	    obj.Exec

	    msg = "[server]client output:" & obj.stdout & vbCrLf & "[server]msg from client" & obj.ServerData

	    MsgBox msg

	End Sub
