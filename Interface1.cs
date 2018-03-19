using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.ServiceModel;

namespace ExecCommand
{
    [Guid("3F5CE45D-7F4F-4B49-9D9B-CE20732B8A1F")]
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

    }

    /// <summary>
    /// IPC関係
    /// </summary>

    [ServiceContract]
    [Guid("0A6098E2-A18B-44AB-854C-14A8053D8F8A")]
    public interface IMyContract
    {
        [OperationContract]
        void SetData(string s);

        [OperationContract]
        string GetData();

    }

    [Guid("28667748-9DCA-4D31-8A0F-FB4BEDBD08FF")]
    public interface IRemoteServer
    {
        string Err
        {
            get;
        }
        string Data
        {
            get;
            set;
        }
        int Init();
        int Close();

    }

    [Guid("AC4B6BC1-5DB1-411D-A8A1-8705CAD32B4E")]
    public interface IRemoteClient
    {


        string Err
        {
            get;
        }

        string Data
        {
            get;
            set;
        }
        int Init();
    }

}
