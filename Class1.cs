using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;

using System.Runtime.Remoting;
using System.Runtime.Remoting.Channels;
using System.Runtime.Remoting.Channels.Ipc;

namespace ExecCommand
{
    public class RemoteObject : MarshalByRefObject
    {
        public string Data { get; set; }
    }


    //[System.Runtime.InteropServices.ComVisible(true)]
    [System.Runtime.InteropServices.ClassInterface(System.Runtime.InteropServices.ClassInterfaceType.None)]
    //[System.Runtime.InteropServices.ClassInterface(System.Runtime.InteropServices.ClassInterfaceType.AutoDispatch)]
    //[System.Runtime.InteropServices.ClassInterface(System.Runtime.InteropServices.ClassInterfaceType.AutoDual)]
    [Guid("098F1F11-4B55-4D66-A5FE-AD4EF0245D34")]
    public class ExecCommand:IExecCommand
    {
        private string exeFile;
        private string arg;
        private string stdout="";
        private string stderr="";
        private string err="";

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

        ///////////////////////////////////////////
        // IPC関係

 
        private RemoteObject m_remoteObject;    // サーバー側共有オブジェクト
        private bool m_bInit = false;

        // サーバー側初期化
        public int InitRemoteServer()
        {
            if (m_bInit == false)
            {
                try
                {
                    //サーバサイドのチャンネルを生成.
                    IpcServerChannel channel = new IpcServerChannel("vba-channel");

                    //チャンネルを登録.
                    ChannelServices.RegisterChannel(channel, true);

                    //やり取りするオブジェクトを生成して登録.
                    m_remoteObject = new RemoteObject();
                    RemotingServices.Marshal(m_remoteObject, "vba-data");
                }
                catch (Exception e)
                {
                    this.err = "サーバー初期化時エラー:" + e.Message;
                    return -1;
                }
                m_bInit = true;
            }
            return 0;

        }

        // サーバー側データアクセスプロパティ
        public string ServerData
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
        public int InitRemoteClient()
        {
            try
            {
                //クライアントサイドのチャンネルを生成.
                IpcClientChannel channel = new IpcClientChannel();

                //チャンネルを登録
                ChannelServices.RegisterChannel(channel, true);
            }
            catch (Exception e)
            {
                this.err = "クライアント初期化時エラー:" + e.Message;
                return -1;
            }
            return 0;
        }

        // クライアント側データアクセスプロパティ
        public string RemoteData
        {
            get
            {
                RemoteObject remoteObject;
                try
                {
                    //リモートオブジェクトを取得
                    //URIは"/チャンネル名/公開名"になる.
                    remoteObject = Activator.GetObject
                        (typeof(RemoteObject), "ipc://vba-channel/vba-data") as RemoteObject;
                    return remoteObject.Data;
                }
                catch (Exception e)
                {
                    this.err = "リモートデータ取得時エラー:" + e.Message;
                    return "";
                }

            }
            set
            {
                RemoteObject remoteObject;
                try
                {
                    //リモートオブジェクトを取得
                    //URIは"/チャンネル名/公開名"になる.
                    remoteObject = Activator.GetObject
                        (typeof(RemoteObject), "ipc://vba-channel/vba-data") as RemoteObject;
                    remoteObject.Data=value;
                }
                catch (Exception e)
                {
                    this.err = "リモートデータ設定時エラー:" + e.Message;
                }

            }
        }

    }
}
