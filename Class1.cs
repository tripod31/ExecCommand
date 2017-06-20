using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;

using System.ServiceModel;

namespace ExecCommand
{

    [System.Runtime.InteropServices.ClassInterface(System.Runtime.InteropServices.ClassInterfaceType.None)]
    [Guid("098F1F11-4B55-4D66-A5FE-AD4EF0245D34")]
    public class ExecCommand : IExecCommand
    {
        private string exeFile;
        private string arg;
        private string stdout = "";
        private string stderr = "";
        private string err = "";

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

    }


    ///////////////////////////////////////////
    // IPC関係

    [ServiceBehavior(InstanceContextMode = InstanceContextMode.Single)]
    public class MyServer : IMyContract
    {
        private string m_str;

        public void SetData(string s)
        {
            m_str=s;
        }

        public string GetData()
        {
            return m_str;
        }
    }

    [System.Runtime.InteropServices.ClassInterface(System.Runtime.InteropServices.ClassInterfaceType.None)]
    [Guid("84cd29f0-c539-4b2d-82a8-345a54caac38")]
    public class RemoteServer:IRemoteServer
    {
        private string err = "";
        private ServiceHost  m_host=null;
        private MyServer m_server=null;

        public string Err
        {
            get { return this.err; }
        }

        // データアクセスプロパティ
        public string Data
        {
            get
            {
                return m_server.GetData();

            }
            set
            {
                m_server.SetData(value);
 
            }
        }

        // 初期化
        public int Init()
        {

            try
            {
                m_server = new MyServer();
                m_host = new ServiceHost(
                    m_server,
                    new Uri("net.pipe://localhost"));
                m_host.AddServiceEndpoint(
                    typeof(IMyContract),
                    new NetNamedPipeBinding(),
                    "InterProcessCommunication");
                m_host.Open();
            }
            catch (Exception e)
            {
                this.err = "サーバー初期化時エラー:" + e.Message;
                return -1;
            }
            return 0;

        }

        // 終了処理
        public int Close()
        {
            if (m_host != null) {
                try
                {
                    m_host.Close();
                }
                catch (Exception e)
                {
                    this.err = "サーバー終了時エラー:" + e.Message;
                    return -1;
                }

            }
            return 0;
        }

    }

    [System.Runtime.InteropServices.ClassInterface(System.Runtime.InteropServices.ClassInterfaceType.None)]
    [Guid("F7039B5B-965F-4A74-8826-91C71BB48AAB")]
    public class RemoteClient:IRemoteClient {
        private IMyContract m_server;
        private string err = "";

        public string Err
        {
            get { return this.err; }
        }

        // 初期化
        public int Init()
        {
            try
            {
                m_server = new ChannelFactory<IMyContract>(
                    new NetNamedPipeBinding(),
                    new EndpointAddress("net.pipe://localhost/InterProcessCommunication")
                    ).CreateChannel();
            }
            catch (Exception e)
            {
                this.err = "クライアント初期化時エラー:" + e.Message;
                return -1;
            }
            return 0;
        }

        // データアクセスプロパティ
        public string Data
        {
            get
            {
                try
                {
                    return m_server.GetData();
                }
                catch (Exception e)
                {
                    this.err = "リモートデータ取得時エラー:" + e.Message;
                    return "";
                }

            }
            set
            {
                try
                {
                    m_server.SetData(value);
                }
                catch (Exception e)
                {
                    this.err = "リモートデータ設定時エラー:" + e.Message;
                }

            }
        }

    }
}
