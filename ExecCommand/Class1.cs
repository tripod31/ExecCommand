using System;
using System.Runtime.InteropServices;
using System.IO.MemoryMappedFiles;

namespace ExecCommand
{

    [ClassInterface(ClassInterfaceType.None)]
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

    // IPC関係
    // 共有データ
    [ClassInterface(ClassInterfaceType.None)]
    [Guid("F30365B5-34E2-4B62-A282-1EA2D2C69212")]
    public class SharedData : ISharedData
    {
        private const string SHARED_MEMORY_NAME = "ExecCommand";
        private const int SHARED_MEMORY_SIZE = 1024 * 1024;

        // 共有メモリに文字列を書き込む
        // 先頭にデータ長（int32）、続いて文字列
        public void SetData(string str)
        {
            // 共有メモリをオープン。既に同じ名前のメモリがあればそれを開く
            MemoryMappedFile　share_mem = MemoryMappedFile.CreateOrOpen(SHARED_MEMORY_NAME, SHARED_MEMORY_SIZE);
            MemoryMappedViewAccessor accessor = share_mem.CreateViewAccessor();

            // 共有メモリに文字列を書き込む
            char[] data = str.ToCharArray();
            accessor.Write(0, data.Length);
            int length = data.Length;
            if (length + sizeof(int) > SHARED_MEMORY_SIZE){
                // データサイズが共有メモリサイズを超えた場合は、超えた分のデータを切り捨てる
                length = SHARED_MEMORY_SIZE - sizeof(int);
            }
            accessor.WriteArray<char>(sizeof(int), data, 0, length);

            accessor.Dispose();
        }

        // 共有メモリから文字列を取得
        // 戻り値：
        // 文字列。エラー発生時はエラー内容
        public string GetData()
        {
            string ret;
            try{           
                // 共有メモリをオープン
                MemoryMappedFile share_mem;
                share_mem = MemoryMappedFile.OpenExisting(SHARED_MEMORY_NAME);
                MemoryMappedViewAccessor accessor = share_mem.CreateViewAccessor();

                // データを読む
                int size = accessor.ReadInt32(0);
                char[] data = new char[size];
                accessor.ReadArray<char>(sizeof(int), data, 0, data.Length);
                ret = new string(data);

                accessor.Dispose();
            } catch (Exception e){
                ret = "GetDataでエラー："+e.Message;
            }
            return ret;
        }
    }
}
