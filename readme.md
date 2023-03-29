# ExecCommand

## これは何か
COMオブジェクトです。  
### **DOS窓を出さずにCUIコマンドを起動する**  
LibreOfficeのマクロからCUIコマンドを起動して標準出力を得るためにWindowsScriptingHostのCOMオブジェクト(wscript.shell)を使用していました。しかしこれだと起動時にDOS窓が表示されるため、うっとおしい。  
そのため、C#でコマンドを起動するCOMオブジェクトを作成しました。  
コマンドの終了を待ち、標準出力・標準エラーを取得します。DOS窓は表示しません。

### **プロセス間通信**
親プロセスから子プロセスに（ExcelからPerlなど）、コマンドラインで渡すのが難しい大きな文字列を渡すためにプロセス間通信をサポートしました。プロセスが親子関係である必要はありません。

## 開発・動作確認環境
Windows11 64bit  
Microsoft VisualStudio Community 2022 64bit  
.NET Framework 4.8
LibleOffice7.4.5.1 64bit  

## インストール
### **ファイルをダウンロードして展開**
execcommand.zipをクリック→Download

### **DLLのレジストリへの登録・登録解除**
register.exeを実行する。DLLがregister.exeと違うディレクトリにある場合、パスを指定する。  
64bitOSで32bitプログラムからDLLを使用するには、register32.exeを使用してDLLを登録する。32bit版のregasm.exeを使用して登録される。
#### **registerボタン**
DLLを登録します。エラーが出ても、「型が正常に登録されました。」と表示されればOK。  
#### **unregisterボタン**
DLLを登録解除します。
#### **regfileボタン**
登録用のレジストリファイルを出力します。

### サンプルファイル
#### **sample/test.ods**
LibreOfficeCalcファイル。COMオブジェクトを呼び出しています。  

#### **test.ps1**
powershellスクリプト。calcファイルから起動され、COMオブジェクトを呼び出しています。

変更履歴
-----
#### 2018/03/17
64bitOSで32bitプログラムからDLLを使用するには、register32.exeを使用してDLLを登録する。
