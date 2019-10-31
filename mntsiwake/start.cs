using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Drawing;

namespace mntsiwake
{
    class start
    {
        ///-------------------------------------------------------------------
        /// <summary>
        ///     ファイル有無検証 </summary>
        /// <param name="workDir">
        ///     インストールディレクトリパス</param>
        /// <returns>
        ///     true:処理ファイルあり、false:処理ファイルなし</returns>
        ///-------------------------------------------------------------------
        public Boolean FileExistChk(string workDir)
        {
            int pInFile = 1;
            int pDivFile = 1;

            //ローカルデータベースの存在を確認
            if (System.IO.File.Exists(workDir + global.DIR_CONFIG + global.CONFIGFILE) == false)
            {
                MessageBox.Show("設定データベースがありません。" + Environment.NewLine + "ソフトを再インストールしてください。","環境設定エラー",MessageBoxButtons.OK,MessageBoxIcon.Exclamation);
                Environment.Exit(0);
            }

            // 設定テーブル:Configより"sub2"フィールドの値を取得
            //Control.DataControl dc = new Control.DataControl(workDir + global.DIR_CONFIG, global.CONFIGFILE);
            //OleDbDataReader dr = dc.free_dsReader("SELECT * FROM Config");

            // ACCESSデータベースへ接続
            SysControl.SetDBConnect Con = new SysControl.SetDBConnect();
            OleDbCommand sCom = new OleDbCommand();
            sCom.Connection = Con.cnOpen();
            sCom.CommandText = "SELECT * FROM Config";
            OleDbDataReader dr = sCom.ExecuteReader();

            while (dr.Read())
            {
                global.pblSelFILE = int.Parse(dr["sub2"].ToString());
            }

            dr.Close();

            // データベース切断
            sCom.Connection.Close();

            if (System.IO.File.Exists(workDir + global.DIR_HENKAN + global.INFILE) == false) 
            {
                pInFile = 0;
                   
                //分割ファイル
                if (System.IO.File.Exists(workDir + global.DIR_INCSV + "*.csv") == false)
                {
                    pDivFile = 0;
                }
            }
            else
            {
                //入力ファイルがある場合、分割ファイルを削除する
                utility.FileDelete(workDir + global.DIR_INCSV,"*");
                pDivFile = 0;
                
            }
             
            //戻り値の判定
            if ((global.pblSelFILE != 1) && (pInFile == 0) && (pDivFile == 0))
            {
                return false;
            }

            return true;
        }

        ///-------------------------------------------------------------------
        /// <summary>
        ///     設定データの取得 </summary>
        /// <param name="workDir">
        ///     インストールディレクトリパス</param>
        ///-------------------------------------------------------------------
        public void InitialLoad(string workDir)
        {
            //Configテーブル
            //Control.DataControl dc = new Control.DataControl(workDir + global.DIR_CONFIG, global.CONFIGFILE);
            //OleDbDataReader dr = dc.free_dsReader("SELECT * FROM Config");

            // ACCESSデータベースへ接続
            SysControl.SetDBConnect Con = new SysControl.SetDBConnect();
            OleDbCommand sCom = new OleDbCommand();
            sCom.Connection = Con.cnOpen();
            sCom.CommandText = "SELECT * FROM Config";
            OleDbDataReader dr = sCom.ExecuteReader();

            try 
	        {

                int sR = 0;
                int sG = 0;
                int sB = 0;

                while (dr.Read())
                {
                    //カラーの取得

                    sR = 0;
                    sG = 0;
                    sB = 0;
                    if (utility.NumericCheck(dr["ErBkR"].ToString())) sR = int.Parse(dr["ErBkR"].ToString());
                    if (utility.NumericCheck(dr["ErBKG"].ToString())) sG = int.Parse(dr["ErBKG"].ToString());
                    if (utility.NumericCheck(dr["ErBkB"].ToString())) sB = int.Parse(dr["ErBkB"].ToString());
                    global.pblErrBackColor = Color.FromArgb(sR,sG,sB);
                    
                    sR = 0;
                    sG = 0;
                    sB = 0;
                    if (utility.NumericCheck(dr["ErFrR"].ToString())) sR = int.Parse(dr["ErFrR"].ToString());
                    if (utility.NumericCheck(dr["ErFrG"].ToString())) sG = int.Parse(dr["ErFrG"].ToString());
                    if (utility.NumericCheck(dr["ErFrB"].ToString())) sB = int.Parse(dr["ErFrB"].ToString());
                    global.pblErrForeColor = Color.FromArgb(sR, sG, sB);
                    
                    sR = 0;
                    sG = 0;
                    sB = 0;
                    if (utility.NumericCheck(dr["KinBkR"].ToString())) sR = int.Parse(dr["KinBkR"].ToString());
                    if (utility.NumericCheck(dr["KinBKG"].ToString())) sG = int.Parse(dr["KinBKG"].ToString());
                    if (utility.NumericCheck(dr["KinBkB"].ToString())) sB = int.Parse(dr["KinBkB"].ToString());
                    global.pblKinBackColor = Color.FromArgb(sR, sG, sB);

                    //画像表示設定
                    if (utility.NumericCheck(dr["ImgH"].ToString()))
                    {
                        global.pblImageHeight = int.Parse(dr["ImgH"].ToString());
                    }
                    else
                    {
                        global.pblImageHeight = 0;
                    }
                    
                    if (utility.NumericCheck(dr["ImgW"].ToString()))
                    {
                        global.pblImageWidth = int.Parse(dr["ImgW"].ToString());
                    }
                    else
                    {
                        global.pblImageWidth = 0;
                    }
                    
                    if (utility.NumericCheck(dr["ImgX"].ToString()))
                    {
                        global.pblImageX = int.Parse(dr["ImgX"].ToString());
                    }
                    else
                    {
                        global.pblImageX = 0;
                    }
                                        
                    //前回選択したデータベース名
                    global.pblBfDbName = dr["BfDb"].ToString().Trim();
                    
                    //接続関連
                    global.pblDsnPath = dr["DsnPath"].ToString();
                    global.pblDsnFlg = dr["DsnFlg"].ToString().Trim();
                    if (utility.NulltoStr(dr["sub1"].ToString()).Trim() == string.Empty)
                    {
                        global.pblKanjoVer = global.VER_21;
                    }
                    else
                    {
                        global.pblKanjoVer = utility.NulltoStr(dr["sub1"].ToString()).Trim();
                    }
                    
                    //メニューで指定した伝票の区分
                    if (utility.NumericCheck(dr["sub2"].ToString()))
                    {
                        global.pblSelFILE = int.Parse(dr["sub2"].ToString().Trim());
                    }
                    else
                    {
                        global.pblSelFILE = 0;
                    }
                    
                    //振替モード　固定部門、勘定科目,補助科目
                    global.pblHeadBumon = string.Empty;
                    global.pblHeadKamoku = string.Empty;
                    global.pblHeadHojo = string.Empty;

                    //カラーの設定                    
                    global.pblBackColor = Color.FromArgb(global.BACK_COLOR);
                    global.pblForeColor = Color.Blue;
                    global.pblNonColor = Color.FromArgb(global.NON_COLOR);
                   
                    //パスワードなし時対応 ↓
                    if (dr["DsnPassWord"] == DBNull.Value)
                    {
                        //NULLの場合
                        global.pblDsnPassWord = string.Empty;
                    }
                    else
                    {
                        global.pblDsnPassWord = dr["DsnPassWord"].ToString().Trim();
                    }
                    
                    //最大結合枚数セット
                    global.pblCombineMax = global.MAXDEN;

                }
	        }
	        catch (Exception e)
	        {
                MessageBox.Show(e.Message, "設定データ取得", MessageBoxButtons.OK);
	        }
            finally
            {
                dr.Close();
                sCom.Connection.Close();
            }
        }
    }
}
