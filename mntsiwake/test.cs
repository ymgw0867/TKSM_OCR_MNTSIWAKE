using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;

namespace mntsiwake
{
    public partial class test : Form
    {
        public test()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Entity.InputRecord[] denData = dSet();
            //errCheck d = new errCheck();
            //d.dShow(denData);
        }

        //private Entity.InputRecord [] dSet()
        //{
        //    //伝票データのインスタンスを生成する
        //    Entity.InputRecord[] Den = new Entity.InputRecord[1];

        //    for (int i = 0; i < 5; i++)
        //    {
        //        if (i != 0)
        //        {
        //            //2件目以降なら要素数を追加
        //            Den.CopyTo(Den = new Entity.InputRecord[i + 1], 0);
        //        }

        //        //行データのインスタンスを生成する
        //        Den[i].Gyou = new Entity.Gyou[global.MAXGYOU];

        //        Den[i].Head.DenNo = (i + 1).ToString();

        //        for (int iX = 0; iX < 7; iX++)
        //        {
        //            Den[i].Gyou[iX].GyouNum = (iX + 1).ToString();
        //            Den[i].Gyou[iX].Tekiyou = (iX + 1).ToString() + "行目データ";
        //        }

        //    }

        //    return Den;
        //}

        private void test_Load(object sender, EventArgs e)
        {
            this.Tag = string.Empty;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Tag = "end";
            this.Close();
        }

        private void test_FormClosing(object sender, FormClosingEventArgs e)
        {
            switch ( e.CloseReason)
            {
                case CloseReason.ApplicationExitCall:
                    MessageBox.Show("Application.Exitによる");
                    break;
                case CloseReason.FormOwnerClosing:
                    MessageBox.Show("所有側のフォームが閉じられようとしている");
                    break;
                case CloseReason.MdiFormClosing:
                    MessageBox.Show("MDIの親フォームが閉じられようとしている");
                    break;
                case CloseReason.TaskManagerClosing:
                    MessageBox.Show("タスクマネージャによる");
                    break;
                case CloseReason.UserClosing:
                    if (Tag.ToString() == string.Empty)
                    {
                        MessageBox.Show("ユーザーインターフェイスによる：×ボタン");
                    }
                    else
                    {
                        MessageBox.Show("ユーザーインターフェイスによる：終了ボタン");
                    }
                    break;
                case CloseReason.WindowsShutDown:
                    MessageBox.Show("OSのシャットダウンによる");
                    break;
                case CloseReason.None:
                default:
                    MessageBox.Show("未知の理由");
                    break;
            }
        }
    }
}
