using System.ServiceModel;

namespace ExecCommand
{
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

    // IPC共有データ
    public interface ISharedData
    {
        [OperationContract]
        void SetData(string s);

        [OperationContract]
        string GetData();

    }

}
