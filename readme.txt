■これは何か
DOSコマンドを起動するCOMオブジェクトです。
Excel VBAからperlを起動して標準出力を得るためにWindowsScriptingHostのCOMオブジェクト(wscript.shell)を使用していましたが、これだと起動時にDOS窓が表示されるため、うっとおしい。
そのため、C#でコマンドを起動するCOMオブジェクトを作成しました。
コマンドの終了を待ち、標準出力・標準エラーを取得します。DOS窓は表示しません。

■開発・動作確認環境
Windows7
Microsoft VisualStudioExpress 2013forWindowsDesktop
.NET Framework 2
Excel2007

■インストール
・ClassLibraryForVBA.zipをインストールするフォルダに解凍します
・DLLのレジストリへの登録
Windows7の場合、管理者でregist_dll.batを実行する。
バッチファイル中のregasm.exeのパスを修正する必要があるかも。

【登録時にエラーが出る場合】
Windows7+IE9の環境で、次のエラーがでました。

RegAsm : error RA0000 : 入力アセンブリ 'ClassLibraryForVBA.dll' またはその依存関
係の 1 つが見つかりません。

IESHIMS.DLLを検索してパスを登録します。検索するといくつか出てきました。最新のファイルの場所は以下の場所でした。
C:\Windows\winsxs\x86_microsoft-windows-ie-ieshims_31bf3856ad364e35_9.4.8112.16476_none_5fdbc489b4a35eb0
これをシステムパスに追加したところ、エラーが出なくなり、登録できました。

■VBAからの使用例
pingコマンドの出力を表示します。

Sub test()
    Dim obj As Object
    Dim iRet As Integer
    
    Set obj = CreateObject("ClassLibraryForVBA.ExecCommand")
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
