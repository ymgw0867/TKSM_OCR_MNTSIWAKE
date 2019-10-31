using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace mntsiwake
{
    public partial class frmFilSelect2 : Form
    {
        public struct MyStruct   //選択された中断ファイルを示す
        {
            public string fPath;
            public Boolean fFlg;
        }

        public MyStruct[] st = new MyStruct[1];

        public frmFilSelect2()
        {
            InitializeComponent();
        }

        private void frmFilSelect2_Load(object sender, EventArgs e)
        {
            //ウィンドウズ最小サイズ
            utility.WindowsMinSize(this, this.Size.Width, this.Size.Height);

            //ウィンドウズ最大サイズ
            utility.WindowsMaxSize(this, this.Size.Width, this.Size.Height);

            //中断伝票をリスト表示
            List1.Items.Clear();
            ListItemShow();

            //ボタンの表示状態
            button1.Enabled = true;
            button2.Enabled = true;
            button3.Enabled = true;
            button4.Enabled = true;

            //終了タグ初期化
            Tag = string.Empty;
        }

        /// <summary>
        /// リストボックスへファイルの内容を表示する
        /// </summary>
        private void ListItemShow()
        {
            int Y = 0;
            int r = 0;

            string readbuf = string.Empty;
            string sPath = global.WorkDir + global.DIR_BREAK + string.Format("{0:000}", global.pblComNo) + @"\";
            string fName = string.Empty;

            foreach (string files in System.IO.Directory.GetFiles(sPath, "*.csv"))
            {
                //ファイル名を配列にセット
                structSet(r, files);
                r++;

                //日付時間を取得
                fName = files.Replace(sPath, string.Empty).Substring(0, 10);

                //異なる日付時間ならListBoxへ表示する
                if ((readbuf != string.Empty) && (readbuf != fName))
                {
                    ListItemSet(Y, readbuf);
                    Y = 0;
                }

                Y++;

                readbuf = fName;
            }

            ListItemSet(Y, readbuf);
        }

        /// <summary>
        /// リストボックスへ表示する
        /// </summary>
        /// <param name="i">日付ごとの件数</param>
        /// <param name="tempName">ファイル名</param>
        private void ListItemSet(int i,string tempName)
        {
            //リストボックスへ表示する
            List1.Items.Add(tempName.Substring(0, 2) + "月" + tempName.Substring(2, 2) + "日" +
                            tempName.Substring(4, 2) + "時" + tempName.Substring(6, 2) + "分" +
                            tempName.Substring(8, 2) + "秒　" + i.ToString() + "件", false);

        }

        /// <summary>
        /// 配列へファイル名を格納する
        /// </summary>
        /// <param name="r">添え字</param>
        /// <param name="files">パスを含むファイル名</param>
        private void structSet(int r,string files)
        {
            //ファイル名を配列へ格納する
            if (r != 0)
            {
                //2件目以降なら要素数を追加
                st.CopyTo(st = new MyStruct[r + 1], 0);
            }

            st[r].fPath = files;
            st[r].fFlg = false;
        }

        /// <summary>
        /// リスト全てをチェック状態とする
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < List1.Items.Count; i++)
            {
                List1.SetItemChecked(i, true);
            }

            button3.Enabled = true;
        }

        /// <summary>
        /// リスト全てを未チェック状態とする
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < List1.Items.Count; i++)
            {
                List1.SetItemChecked(i, false);
            }

            button3.Enabled = false;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (List1.CheckedIndices.Count == 0)
            {
                MessageBox.Show("伝票が選択されていません。", "中断伝票未選択", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            if (MessageBox.Show("中断処理データを読み込みます。よろしいですか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No) return;

            this.Tag = "button3";


            //チェック日付に該当する配列フラグをtrueにする
            foreach (var item in List1.CheckedItems)
            {
                string stDate = item.ToString();

                stDate = stDate.ToString().Replace("月", string.Empty);    //"月"消去
                stDate = stDate.ToString().Replace("日", string.Empty);    //"日"消去
                stDate = stDate.ToString().Replace("時", string.Empty);    //"時"消去
                stDate = stDate.ToString().Replace("分", string.Empty);    //"分"消去
                stDate = stDate.ToString().Replace("秒　", string.Empty);  //"秒"消去
                stDate = stDate.ToString().Replace("件", string.Empty);    //"件"消去
                stDate = stDate.Substring(0, 10);

                for (int i = 0; i < st.Length; i++)
                {
                    //ファイル名に日付を含んでいたらFlgをtrueにする
                    if (st[i].fPath.IndexOf(stDate) >= 0) st[i].fFlg = true;                    
                }
            }
            
            this.Hide();
        }

        private void List1_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            //チェック項目がなければ確定ボタンをEnable=falseとする
            //if (List1.CheckedItems.Count > 0)
            //{
            //    button3.Enabled = true;
            //}
            //else
            //{
            //    button3.Enabled = false;
            //}
        }

        private void button4_Click(object sender, EventArgs e)
        {
            this.Tag = "button4";
            this.Close();
        }

        private void frmFilSelect2_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (this.Tag.ToString() != "button3")
            {
                if (MessageBox.Show("変換プログラムを終了します。よろしいですか？", "終了", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    //エラー終了処理
                    this.Dispose();
                    errEnd.Exit();
                }
                else
                {
                    e.Cancel = true;
                    return;
                }
            }

            this.Dispose();
        }

    }
}
