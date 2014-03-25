using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;


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
    }
}
