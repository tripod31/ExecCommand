using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;

using System.Runtime.Remoting;
using System.Runtime.Remoting.Channels;
using System.Runtime.Remoting.Channels.Ipc;

namespace ClassLibraryForVBA
{
    [System.Runtime.InteropServices.ComVisible(true)]
    //[System.Runtime.InteropServices.ClassInterface(System.Runtime.InteropServices.ClassInterfaceType.AutoDispatch)]
    [System.Runtime.InteropServices.ClassInterface(System.Runtime.InteropServices.ClassInterfaceType.AutoDual)]
    [Guid(ExecCommand.ClassId)]
    public class ExecCommand:IExecCommand
    {
        public const string ClassId = "098F1F11-4B55-4D66-A5FE-AD4EF0245D34";

        private string exeFile;
        private string arg;
        private string stdout;
        private string stderr;
        private string err;

        //実行ファイルパス
        public string ExeFile
        {
            set { this.exeFile = value; }
        }

        //引数
        public string Arg
        {
            set { this.arg = value; }
        }

        //標準出力
        public string StdOut
        {
            get { return this.stdout; }
        }

        //標準エラー
        public string StdErr
        {
            get { return this.stderr; }
        }

        //コマンド起動時エラー
        public string Err
        {
            get { return this.err; }
        }

        //戻り値
        //  0   正常
        //  -1  コマンド起動時エラー
        public int Exec()
        {
            System.Diagnostics.ProcessStartInfo psi =
                new System.Diagnostics.ProcessStartInfo();

            psi.FileName = this.exeFile;
            //出力を読み取れるようにする
            psi.RedirectStandardInput = false;
            psi.RedirectStandardOutput = true;
            psi.RedirectStandardError = true;
            psi.UseShellExecute = false;
            //ウィンドウを表示しないようにする
            psi.CreateNoWindow = true;

            psi.Arguments = this.arg;
            //起動
            this.stderr = "";
            this.stdout = "";
            this.err = "";
            System.Diagnostics.Process p;
            try
            {
                p = System.Diagnostics.Process.Start(psi);
            }
            catch (Exception e)
            {
                this.err = "EXE起動時エラー:" + e.Message;
                return -1;
            }
            //出力を読み取る
            string stdout = p.StandardOutput.ReadToEnd();
            string stderr = p.StandardError.ReadToEnd();
            //WaitForExitはReadToEndの後である必要がある
            //(親プロセス、子プロセスでブロック防止のため)
            p.WaitForExit();

            this.stdout = stdout;
            this.stderr = stderr;
            return 0;
        }

        // IPC関係
        public class RemoteObject : MarshalByRefObject
        {
            public string Data { get; set; }
        }
        RemoteObject m_remoteObject;    // 共有オブジェクト

        // サーバー側初期化
        public void InitRemoteServer()
        {

            //サーバサイドのチャンネルを生成.
            IpcServerChannel channel = new IpcServerChannel("vba-channel");
            
            //チャンネルを登録.
            ChannelServices.RegisterChannel(channel, true);

            //やり取りするオブジェクトを生成して登録.
            m_remoteObject= new RemoteObject();
            RemotingServices.Marshal(m_remoteObject,"vba-data");

        }

        // サーバー側データアクセスプロパティ
        public string RemoteData4Server
        {
            get
            {
                return m_remoteObject.Data;
            }
            set
            {
                m_remoteObject.Data = value;
            }
        }

        // クライアント側初期化
        public void InitRemoteClient()
        {
            //クライアントサイドのチャンネルを生成.
            IpcClientChannel channel = new IpcClientChannel();

            //チャンネルを登録
            ChannelServices.RegisterChannel(channel, true);
        }

        // クライアント側データアクセスプロパティ
        public string RemoteData4Client
        {
            get
            {
                //リモートオブジェクトを取得
                //URIは"/チャンネル名/公開名"になる.
                RemoteObject remoteObject = Activator.GetObject
                    (typeof(RemoteObject), "ipc://vba-channel/vba-data") as RemoteObject;

                return remoteObject.Data;
            }
            set
            {
                //リモートオブジェクトを取得
                //URIは"/チャンネル名/公開名"になる.
                RemoteObject remoteObject = Activator.GetObject
                    (typeof(RemoteObject), "ipc://vba-channel/vba-data") as RemoteObject;

                remoteObject.Data=value;
            }
        }

    }
}
