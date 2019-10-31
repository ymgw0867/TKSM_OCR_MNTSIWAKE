using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Data.SqlClient;

namespace mntsiwake
{
    public partial class frmComSelect : Form
    {
        public frmComSelect()
        {
            InitializeComponent();
        }

        private void frmComSelect_Load(object sender, EventArgs e)
        {
            //ウィンドウズ最小サイズ
            utility.WindowsMinSize(this, this.Size.Width, this.Size.Height);

            //ウィンドウズ最大サイズ
            utility.WindowsMaxSize(this, this.Size.Width, this.Size.Height);

            //DataGridViewの設定
            GridViewSetting(dg1);
            GridViewSetting2(dg2);

            // 接続文字列取得 2016/10/01
            string sc = SqlControl.obcConnectSting.get(Properties.Settings.Default.sqlCurrentDB);     

            // データ表示
            GridViewShowData(sc, dg1);

            //終了時タグ初期化
            Tag = string.Empty;
        }

        /// <summary>
        /// データグリッドビューの定義を行います
        /// </summary>
        /// <param name="tempDGV">データグリッドビューオブジェクト</param>
        public void GridViewSetting(DataGridView tempDGV)
        {
            try
            {
                //フォームサイズ定義

                // 列スタイルを変更する

                tempDGV.EnableHeadersVisualStyles = false;

                // 列ヘッダー表示位置指定
                tempDGV.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;

                // 列ヘッダーフォント指定
                tempDGV.ColumnHeadersDefaultCellStyle.Font = new Font("ＭＳ Ｐゴシック", 9, FontStyle.Regular);

                // データフォント指定
                tempDGV.DefaultCellStyle.Font = new Font("ＭＳ Ｐゴシック", (float)9.5, FontStyle.Regular);

                // 行の高さ
                tempDGV.ColumnHeadersHeight = 18;
                tempDGV.RowTemplate.Height = 18;

                // 全体の高さ
                tempDGV.Height = 180;

                // 奇数行の色
                //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                // 各列幅指定
                tempDGV.Columns.Add("col1", "No");
                tempDGV.Columns.Add("col2", "期首");
                tempDGV.Columns.Add("col3", "決算期");
                tempDGV.Columns.Add("col4", "会社名");
                tempDGV.Columns.Add("col5", "dbnm");
                tempDGV.Columns.Add("col6", "taxmas");
                tempDGV.Columns.Add("col7", "reki");

                tempDGV.Columns[1].Visible = false; //期首は非表示
                tempDGV.Columns[2].Visible = false; //決算期は非表示
                tempDGV.Columns[4].Visible = false; //データベース名は非表示
                tempDGV.Columns[5].Visible = false; //税区分は非表示
                tempDGV.Columns[6].Visible = false; //暦は非表示

                tempDGV.Columns[0].Width = 80;
                tempDGV.Columns[1].Width = 100;
                tempDGV.Columns[2].Width = 100;
                tempDGV.Columns[3].Width = 200;

                tempDGV.Columns[3].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

                tempDGV.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                tempDGV.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                tempDGV.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                tempDGV.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;

                // 行ヘッダを表示しない
                tempDGV.RowHeadersVisible = false;

                // 選択モード
                tempDGV.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                tempDGV.MultiSelect = false;

                // 編集不可とする
                tempDGV.ReadOnly = true;

                // 追加行表示しない
                tempDGV.AllowUserToAddRows = false;

                // データグリッドビューから行削除を禁止する
                tempDGV.AllowUserToDeleteRows = false;

                // 手動による列移動の禁止
                tempDGV.AllowUserToOrderColumns = false;

                // 列サイズ変更禁止
                tempDGV.AllowUserToResizeColumns = false;

                // 行サイズ変更禁止
                tempDGV.AllowUserToResizeRows = false;

                // 行ヘッダーの自動調節
                //tempDGV.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders;

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "エラーメッセージ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void GridViewSetting2(DataGridView tempDGV)
        {
            try
            {
                //フォームサイズ定義

                // 列スタイルを変更する

                tempDGV.EnableHeadersVisualStyles = false;

                // 列ヘッダー表示位置指定
                tempDGV.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;

                // 列ヘッダーフォント指定
                tempDGV.ColumnHeadersDefaultCellStyle.Font = new Font("ＭＳ Ｐゴシック", 9, FontStyle.Regular);

                // データフォント指定
                tempDGV.DefaultCellStyle.Font = new Font("ＭＳ Ｐゴシック", (float)9.5, FontStyle.Regular);

                // 行の高さ
                tempDGV.ColumnHeadersHeight = 18;
                tempDGV.RowTemplate.Height = 18;

                // 全体の高さ
                tempDGV.Height = 180;

                // 奇数行の色
                //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                // 各列幅指定
                tempDGV.Columns.Add("col1", "pid");
                tempDGV.Columns.Add("col2", "決算期");
                tempDGV.Columns.Add("col3", "会計期首");
                tempDGV.Columns.Add("col4", "会計期末");
                tempDGV.Columns.Add("col5", "中間");

                tempDGV.Columns[0].Visible = false;
                tempDGV.Columns[1].Width = 100;
                tempDGV.Columns[2].Width = 110;
                tempDGV.Columns[3].Width = 110;
                tempDGV.Columns[4].Visible = false;

                tempDGV.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

                tempDGV.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                tempDGV.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                tempDGV.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

                // 行ヘッダを表示しない
                tempDGV.RowHeadersVisible = false;

                // 選択モード
                tempDGV.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                tempDGV.MultiSelect = false;

                // 編集不可とする
                tempDGV.ReadOnly = true;

                // 追加行表示しない
                tempDGV.AllowUserToAddRows = false;

                // データグリッドビューから行削除を禁止する
                tempDGV.AllowUserToDeleteRows = false;

                // 手動による列移動の禁止
                tempDGV.AllowUserToOrderColumns = false;

                // 列サイズ変更禁止
                tempDGV.AllowUserToResizeColumns = false;

                // 行サイズ変更禁止
                tempDGV.AllowUserToResizeRows = false;

                // 行ヘッダーの自動調節
                //tempDGV.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders;

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "エラーメッセージ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        ///----------------------------------------------------------------------
        /// <summary>
        ///     グリッドビューへ会社情報を表示する </summary>
        /// <param name="sConnect">
        ///     接続文字列</param>
        /// <param name="tempDGV">
        ///     DataGridViewオブジェクト名</param>
        ///----------------------------------------------------------------------
        private void GridViewShowData(string sConnect, DataGridView tempDGV)
        {
            string sqlSTRING = string.Empty;

            SqlControl.DataControl sdcon = new SqlControl.DataControl(sConnect);

            //データリーダーを取得する
            //OleDbDataReader dR;
            SqlDataReader dR;

            //////sqlSTRING += "SELECT sDbNm,siCorpNo,sDateKisyu,sHosei,siKsnKi,sCorpNm FROM wcompany";

            //////sqlSTRING += "SELECT DatabaseName, EntityCode,EntityName, CreateDate, CorpData ";
            //////sqlSTRING += "FROM tbCorpDatabaseContext ORDER BY EntityCode";

            sqlSTRING += "select * from ";
            sqlSTRING += "(select tbCorpDatabaseContext.EntityCode,tbCorpDatabaseContext.EntityName,";
            sqlSTRING += "tbCorpDatabaseContext.DatabaseName,tbCorpDatabaseContext.CreateDate,";
            sqlSTRING += "CorpData.value('(/ObcCorpData/Node[@key=\"InitializeAC\"])[1]','varchar(1)') as Type ";
            sqlSTRING += "from tbCorpDatabaseContext) as Corp ";
            sqlSTRING += "where Type is not null ";
            sqlSTRING += "order by EntityCode";

            dR = sdcon.free_dsReader(sqlSTRING);

            try
            {
                //グリッドビューに表示する
                int iX = 0;
                tempDGV.RowCount = 0;

                while (dR.Read())
                {
                    ////////中断データ処理のときは該当する会社しか表示しない 
                    //////if (global.pblSelFILE == 1)
                    //////{
                    //////    this.Text = "中断伝票処理　会社選択";
                    //////    global.pblBfDbName = string.Empty;

                    //////    //該当する会社の中断フォルダが存在するとき
                    //////    if (System.IO.Directory.Exists(global.WorkDir + global.DIR_BREAK + string.Format("{0:000}", int.Parse(dR["EntityCode"].ToString())) + @"\"))
                    //////    {
                    //////        //データグリッドにデータを表示する
                    //////        tempDGV.Rows.Add();
                    //////        GridViewCellData(tempDGV, iX, dR);
                    //////        break;
                    //////    }
                    //////}
                    //////else
                    //////{

                    //this.Text = string.Empty;

                    //データグリッドにデータを表示する
                    tempDGV.Rows.Add();
                    GridViewCellData(tempDGV, iX, dR);

                    iX++;

                    //////}
                }

                tempDGV.CurrentCell = null;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "会社情報を表示エラー", MessageBoxButtons.OK);
            }
            finally
            {
                dR.Close();
                sdcon.Close();
            }

            int sIx;

            //会社情報がないとき
            if (tempDGV.RowCount == 0) 
            {
                MessageBox.Show("会社情報が存在しません", "会社選択", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            else
            {
                //前回選択した会社を検索
                if (global.pblBfDbName == string.Empty)
                {
                    sIx = 0;
                }
                else
                {
                    sIx= -1;
        
                    for (int i = 0; i < tempDGV.RowCount; i++)
			        {
			            if (tempDGV[4,i].Value.ToString() == global.pblBfDbName)
                        {
                            sIx = i;
                            break;
                        }
			        }

                    //もし会社が見つからなければIndex=0
                    if (sIx == -1) sIx = 0;
                }
                
                //税抜別段
                if (tempDGV[5, sIx].Value.ToString() == "0")
                {
                    //表示名変更「会社を選択してください。」 (v6.0対応)--
                    this.lblMsg.Text = "会社を選択してください。";
                }
                //税込自動
                else
                {
                    if (tempDGV[5, sIx].Value.ToString() == "2")
                    {
                        this.lblMsg.Text = "会社を選択してください。";
                    }

                //税抜自動
                    else
                    {
                        this.lblMsg.Text = "会社及び税処理を選択してください。";
                    }
                }

                //対象の会社を選択状態にする
                //tempDGV.Rows[sIx].Selected = true;
            }
        }

        ///----------------------------------------------------------------------
        /// <summary>
        ///     データグリッドに表示データをセットする </summary>
        /// <param name="tempDGV">
        ///     datagridviewオブジェクト名</param>
        /// <param name="iX">
        ///     Row№</param>
        /// <param name="dR">
        ///     データリーダーオブジェクト名</param>
        ///----------------------------------------------------------------------
        private void GridViewCellData(DataGridView tempDGV, int iX, SqlDataReader dR)
        {
            string sKishudate;
            string sKessan;

            //会社№
            tempDGV[0, iX].Value = dR["EntityCode"].ToString();

            //会計期間のフォーマット
            GetKishu(dR["DatabaseName"].ToString(), out sKishudate, out sKessan);  //期首日付、決算期取得

            int yy = int.Parse(sKishudate.Substring(0, 4));
            int mm = int.Parse(sKishudate.Substring(5, 2));
            int dd = int.Parse(sKishudate.Substring(8, 2));

            //西暦・和暦の区分を取得
            tempDGV[6, iX].Value = GetReki(dR["DatabaseName"].ToString());

            if (tempDGV[6, iX].Value.ToString() == global.RWAREKI)
            {
                yy = yy -  Properties.Settings.Default.hosei;
            }
            else
            {
                yy = int.Parse(yy.ToString().Substring(2, 2));
            }

            tempDGV[1, iX].Value = string.Format("{0, 2}", yy) + "/" + string.Format("{0, 2}", mm) + "/" + string.Format("{0, 2}", dd);

            //決算期
            tempDGV[2, iX].Value = "第" + sKessan + "期";

            //会社名
            tempDGV[3, iX].Value = dR["EntityName"].ToString().Trim();

            //非表示項目
            tempDGV[4, iX].Value = dR["DatabaseName"].ToString().Trim();       //データベース名
            tempDGV[5, iX].Value = GetTaxMas(dR["DatabaseName"].ToString());   //税処理区分
        }

        ///----------------------------------------------------------------------
        /// <summary>
        ///     各社の消費税計算区分を取得 </summary>
        /// <param name="sDBName">
        ///     接続するデータベース名</param>
        /// <returns>
        ///     消費税計算区分</returns>
        ///----------------------------------------------------------------------
        private string GetTaxMas(string sDBName)
        {
            //接続文字列を取得する 2016/10/01
            string sc = SqlControl.obcConnectSting.get(sDBName);

            //データベースへ接続する 2016/10/01
            SqlControl.DataControl dCon = new SqlControl.DataControl(sc);

            //データリーダーを取得する
            //OleDbDataReader dR;
            SqlDataReader dR;
            dR = dCon.free_dsReader("SELECT AccountingPeriodID,ConsumptionTaxBasicEntryID,ConsumptionTaxCalcMethod,ConsumptionTaxSubtractionMethod,SalesConsumptionTaxPileCalc,BuyConsumptionTaxPileCalc,EnterpriseDivisionID,AutomaticCalculationTax,FractionWay,RowVersion,RatioApplyToTaxationSalesRatio FROM tbConsumptionTaxBasicEntry");

            string sZei = string.Empty;

            while (dR.Read())
            {
                sZei = dR["AutomaticCalculationTax"].ToString().Trim();
            }

            dR.Close();
            dCon.Close();

            //値を返す
            return sZei;
        }

        ///----------------------------------------------------------------------
        /// <summary>
        ///     各社の期首,決算期を取得  </summary>
        /// <param name="sDBName">
        ///     接続するデータベース名</param>
        /// <returns>
        ///     </returns>
        ///----------------------------------------------------------------------
        private bool GetKishu(string sDBName, out string rKishuDate, out string rKessan)
        {
            //接続文字列を取得する 2016/10/01
            string sc = SqlControl.obcConnectSting.get(sDBName);

            //データベースへ接続する
            SqlControl.DataControl dCon = new SqlControl.DataControl(sc);

            //有効なIDのデータリーダーを取得する
            //OleDbDataReader dR;
            SqlDataReader dR;
            dR = dCon.free_dsReader("SELECT AccountingPeriodID FROM tbAccountingPeriodConfig");

            try
            {
                dR.Read();
                string wID = dR["AccountingPeriodID"].ToString();
                dR.Close();

                //会計情報のデータリーダーを取得する
                dR = dCon.free_dsReader("SELECT AccountingPeriodID, PeriodStartDate,PeriodEndDate,FiscalTerm, CodeContext, RowVersion FROM tbAccountingPeriod WHERE AccountingPeriodID = " + wID + " ORDER BY AccountingPeriodID");

                string sKishuDate = string.Empty;

                rKishuDate = string.Empty;
                rKessan = string.Empty;

                while (dR.Read())
                {
                    rKishuDate = DateTime.Parse(dR["PeriodStartDate"].ToString()).ToShortDateString();
                    rKessan = dR["FiscalTerm"].ToString();
                }

                dR.Close();
                dCon.Close();

                //値を返す
                return true;
            }
            catch (Exception)
            {
                rKishuDate = string.Empty;
                rKessan = string.Empty;
                return false;
            }
            finally
            {
                if (dR.IsClosed == false) dR.Close();
                if (dCon.Cn.State == ConnectionState.Open) dCon.Close();
            }
        }
        
        ///----------------------------------------------------------------
        /// <summary>
        ///     西暦、和暦の区分を取得 </summary>
        /// <param name="sDBName">
        ///     接続するデータベース名</param>
        /// <returns>
        ///     </returns>
        ///----------------------------------------------------------------
        private string GetReki(string sDBName)
        {
            //接続文字列を取得する 2016/10/01
            string sc = SqlControl.obcConnectSting.get(sDBName);

            //データベースへ接続する
            SqlControl.DataControl dCon = new SqlControl.DataControl(sc);

            //データリーダーを取得する
            //OleDbDataReader dR;
            SqlDataReader dR;
            string sqlString = string.Empty;
            sqlString += "select CorpData.value('(/ObcCorpData/Node[@key=\"EraIndicate\"])[1]','varchar(1)') as reki from tbCorpDatabaseContext";
            string sReki = string.Empty;
            dR = dCon.free_dsReader(sqlString);

            try
            {
                dR.Read();
                sReki = dR["reki"].ToString();
                dR.Close();
                dCon.Close();

                //値を返す
                return sReki;
            }
            catch (Exception)
            {
                return sReki;
            }
            finally
            {
                if (dR.IsClosed == false) dR.Close();
                if (dCon.Cn.State == ConnectionState.Open) dCon.Close();
            }
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            //会社情報がないときはそのままクローズ
            if (dg1.RowCount == 0)
            {
                global.pblComNo = string.Empty;     //会社№
                global.pblDbName = string.Empty;    //データベース名
            }
            else
            {
                if (dg1.SelectedRows.Count == 0)
                {
                    MessageBox.Show("会社を選択してください", "会社未選択", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }

                if (dg2.SelectedRows.Count == 0)
                {
                    MessageBox.Show("会計期間を選択してください", "会計期間未選択", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }

                //選択した会社、会計期間情報を取得する
                global.pblComNo = dg1[0, dg1.SelectedRows[0].Index].Value.ToString();                   //会社№
                global.pblDbName = dg1[4, dg1.SelectedRows[0].Index].Value.ToString();                  //データベース名
                global.gsTaxMas = dg1[5, dg1.SelectedRows[0].Index].Value.ToString();                   //税処理区分
                global.pblReki = dg1[6, dg1.SelectedRows[0].Index].Value.ToString();                    //暦
                global.pblAccPID = dg2[0, dg2.SelectedRows[0].Index].Value.ToString();                  //AccountPeriodID

                company.Name = dg1[3, dg1.SelectedRows[0].Index].Value.ToString();                      //会社名    
                
                company.fromDate = DateTime.Parse(dg2[2, dg2.SelectedRows[0].Index].Value.ToString());  //期首年月日  
                company.FromYear = dg2[2, dg2.SelectedRows[0].Index].Value.ToString().Substring(0, 4);  //期首年
                company.FromMonth = dg2[2, dg2.SelectedRows[0].Index].Value.ToString().Substring(5, 2); //期首月
                company.FromDay = dg2[2, dg2.SelectedRows[0].Index].Value.ToString().Substring(8, 2);   //期首日

                company.ToDate = DateTime.Parse(dg2[3, dg2.SelectedRows[0].Index].Value.ToString());    //期末年月日
                company.ToYear = dg2[3, dg2.SelectedRows[0].Index].Value.ToString().Substring(0, 4);    //期末年
                company.ToMonth = dg2[3, dg2.SelectedRows[0].Index].Value.ToString().Substring(5, 2);   //期末月
                company.ToDay = dg2[3, dg2.SelectedRows[0].Index].Value.ToString().Substring(8, 2);     //期末日

                //////company.Kaisi = dR["tiKaisi"].ToString().Trim();

                company.Gengou = Properties.Settings.Default.gengou;
                company.Hosei = Properties.Settings.Default.hosei.ToString();
                company.Middle = dg2[4, dg2.SelectedRows[0].Index].Value.ToString();                    //中間

                //西暦のとき
                if (global.pblReki == global.RSEIREKI)
                {
                    company.Reki = "20";
                }
                else
                {
                    //和暦のとき
                    company.Reki = company.Gengou;
                }

                //////company.gsVersion = dR["tiIsVersion"].ToString().Trim();


                //選択されたデータベース名をローカルデータベースへを記録
                //Control.FreeSql sCon = new Control.FreeSql(global.WorkDir + global.DIR_CONFIG, global.CONFIGFILE);
                //sCon.Execute("update Config set BfDb = '" + global.pblDbName + "'");
                //sCon.Close();
                
                // ACCESSデータベースへ接続 2017/06/06
                SysControl.SetDBConnect Con = new SysControl.SetDBConnect();
                OleDbCommand sCom = new OleDbCommand();
                sCom.Connection = Con.cnOpen();
                sCom.CommandText = "update Config set BfDb = '" + global.pblDbName + "'";
                sCom.ExecuteNonQuery();
                sCom.Connection.Close();
            }

            //フォームを閉じる
            Tag = "btn";
            this.Close();
        }

        private void frmComSelect_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                if (Tag.ToString() == string.Empty)
                {
                    if (MessageBox.Show("変換プログラムを終了します。よろしいですか？", "終了", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == DialogResult.Yes)
                    {
                        //エンド処理
                        errEnd.Exit();
                    }
                    else
                    {
                        e.Cancel = true;
                        return;
                    }
                }
            }
            this.Dispose();
        }

        private void dg1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            //会社決算情報取得
            GetAccountPeriod(dg1[4, dg1.SelectedRows[0].Index].Value.ToString(), dg2);
            //会社の整理仕訳区分取得
            GetArrangeDivision(dg1[4, dg1.SelectedRows[0].Index].Value.ToString());
        }

        ///-----------------------------------------------------------------
        /// <summary>
        ///     会社決算情報取得グリッド表示 </summary>
        /// <param name="sDBName">
        ///     会社データベース名</param>
        /// <param name="tempDGV">
        ///     データグリッドオブジェクト名</param>
        ///-----------------------------------------------------------------
        private void GetAccountPeriod(string sDBName, DataGridView tempDGV)
        {
            //接続文字列を取得する 2016/10/01
            string sc = SqlControl.obcConnectSting.get(sDBName);

            //データベースへ接続する
            SqlControl.DataControl dCon = new SqlControl.DataControl(sc);

            //有効なIDのデータリーダーを取得する
            //OleDbDataReader dR;
            SqlDataReader dR;
            string strSql = string.Empty;

            strSql += "SELECT * FROM tbAccountingPeriod JOIN tbAccountingPeriodConfig ";
            strSql += "ON tbAccountingPeriod.AccountingPeriodID = tbAccountingPeriodConfig.AccountingPeriodID ";
            strSql += "WHERE tbAccountingPeriod.CodeContext = 0 ";
            strSql += "ORDER BY tbAccountingPeriod.AccountingPeriodID desc";

            dR = dCon.free_dsReader(strSql);

            int iX = 0;
            tempDGV.RowCount = 0;

            try
            {
                while (dR.Read())
                {
                    //データグリッドにデータを表示する
                    tempDGV.Rows.Add();
                    tempDGV[0, iX].Value = dR["AccountingPeriodID"].ToString();
                    tempDGV[1, iX].Value = dR["FiscalTerm"].ToString();
                    tempDGV[2, iX].Value = DateTime.Parse(dR["PeriodStartDate"].ToString()).ToShortDateString();
                    tempDGV[3, iX].Value = DateTime.Parse(dR["PeriodEndDate"].ToString()).ToShortDateString();
                    tempDGV[4, iX].Value = dR["FinancialClosingFrequency"].ToString();

                    iX++;
                }
                tempDGV.CurrentCell = null;

                dR.Close();
                dCon.Close();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "エラー", MessageBoxButtons.OK); 
            }
            finally
            {
                if (dR.IsClosed == false) dR.Close();
                if (dCon.Cn.State == ConnectionState.Open) dCon.Close();
            }
        }

        ///-----------------------------------------------------------------
        /// <summary>
        ///     整理仕訳区分を取得する </summary>
        /// <param name="sDBName">
        ///     会社データベース名</param>
        ///-----------------------------------------------------------------
        private void GetArrangeDivision(string sDBName)
        {
            //接続文字列を取得する 2016/10/01
            string sc = SqlControl.obcConnectSting.get(sDBName);

            //データベースへ接続する
            SqlControl.DataControl dCon = new SqlControl.DataControl(sc);

            //有効な整理仕訳区分の値を取得する
            //OleDbDataReader dR;
            SqlDataReader dR;
            string strSql = string.Empty;
            strSql += "SELECT ArrangeDivision FROM tbAccountingPeriodConfig ";
            dR = dCon.free_dsReader(strSql);

            company.Arrange = global.FLGOFF;

            try
            {
                while (dR.Read())
                {
                    company.Arrange = dR["ArrangeDivision"].ToString();
                }
                dR.Close();
                dCon.Close();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "エラー", MessageBoxButtons.OK);
            }
            finally
            {
                if (dR.IsClosed == false) dR.Close();
                if (dCon.Cn.State == ConnectionState.Open) dCon.Close();
            }
        }
    }
}
