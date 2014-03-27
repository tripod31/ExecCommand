using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;

namespace ClassLibraryForVBA
{
    //[System.Runtime.InteropServices.ComVisible(true)]
    //[System.Runtime.InteropServices.InterfaceType(System.Runtime.InteropServices.ComInterfaceType.InterfaceIsDual)]
    [System.Runtime.InteropServices.Guid("3F5CE45D-7F4F-4B49-9D9B-CE20732B8A1F")]
    public interface IExecCommand
    {
        //実行ファイルパス
        string ExeFile
        {
            set;
        }

        //引数
        string Arg
        {
            set;
        }

        //標準出力
        string StdOut
        {
            get;
        }

        //標準エラー
        string StdErr
        {
            get;
        }

        //コマンド起動時エラー
        string Err
        {
            get;
        }

        //戻り値
        //  0   正常
        //  -1  コマンド起動時エラー
        int Exec();

        void InitRemoteServer();
        string ServerData
        {
            get;
            set;
        }
        void InitRemoteClient();
        string RemoteData
        {
            get;
            set;
        }
    }
}
