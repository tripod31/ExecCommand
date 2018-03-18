using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;

namespace register
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            // RegAsm のパスを取得
            string path = System.IO.Path.Combine(RuntimeEnvironment.GetRuntimeDirectory(), "RegAsm.exe");
            this.textBox2.Text = path;
        }

        private string exec_regasm(string arg)
        {
            string msg="";
            try
            {
                // RegAsm のパスを取得
                string path = System.IO.Path.Combine(RuntimeEnvironment.GetRuntimeDirectory(), "RegAsm.exe");
                // パスをコンソールに出力
                Console.WriteLine("[" + path + "]");

                System.Diagnostics.Process p = new System.Diagnostics.Process();
                p.StartInfo.FileName = path;

                p.StartInfo.Arguments = arg;
                // 出力を取得できるようにする
                p.StartInfo.UseShellExecute = false;
                p.StartInfo.RedirectStandardOutput = true;
                p.StartInfo.RedirectStandardInput = false;
                p.StartInfo.RedirectStandardError = true;
                // ウィンドウを表示しない
                p.StartInfo.CreateNoWindow = true;

                // 起動
                p.Start();

                // 出力を取得
                string err = p.StandardError.ReadToEnd();
                if (err.Length == 0)
                {
                    err = "None";
                }
                msg = string.Format("output:\n{0}\nerror:\n{1}\n",p.StandardOutput.ReadToEnd(),err);
                
                // プロセス終了まで待機する
                p.WaitForExit();
                p.Close();
            } catch (Exception ex)
            {
                msg = ex.Message;
            }
            return msg;
        }

        // registerボタン
        private void button1_Click(object sender, EventArgs e)
        {
            string msg =exec_regasm(string.Format("/codebase \"{0}\"",this.textBox1.Text));
            MessageBox.Show(msg);
        }

        // unregisterボタン
        private void button2_Click(object sender, EventArgs e)
        {
            string msg = exec_regasm(string.Format("/unregister /verbose \"{0}\"", this.textBox1.Text));
            MessageBox.Show(msg);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.openFileDialog1.Filter = "DLL|*.dll";
            this.openFileDialog1.InitialDirectory = System.Environment.CurrentDirectory;

            DialogResult dr = this.openFileDialog1.ShowDialog();
            if (dr == DialogResult.OK)
            {
                this.textBox1.Text = this.openFileDialog1.FileName;
            }
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            // save settings
            Properties.Settings.Default.Save();
        }

        // regfileボタン
        private void button4_Click(object sender, EventArgs e)
        {
            this.saveFileDialog1.Filter = "reg|*.reg";
            this.saveFileDialog1.InitialDirectory = System.Environment.CurrentDirectory;

            DialogResult dr = this.saveFileDialog1.ShowDialog();
            if (dr != DialogResult.OK)
            {
                return;
            }

            string msg = exec_regasm(
                string.Format("/codebase \"{0}\" /regfile:\"{1}\"", 
                this.textBox1.Text,
                this.saveFileDialog1.FileName));
            MessageBox.Show(string.Format("{0}\n{1} created.\n",msg, this.saveFileDialog1.FileName));
        }
    }
}
