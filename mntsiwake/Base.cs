using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Data.OleDb;
using System.Data.SqlClient;
using Microsoft.VisualBasic;
using Leadtools.Codecs;
using Leadtools;
using Leadtools.WinForms;
using Leadtools.ImageProcessing;
using GrapeCity.Win.MultiRow;

namespace mntsiwake
{
    ///----------------------------------------------------
    /// <summary>
    ///     徳島産業 : 勘定奉行i10 </summary>
    ///----------------------------------------------------
    public partial class Base : Form
    {
        Entity.InputRecord[] DenData;   //伝票データ配列
        errCheck.Errtbl[] eTbl;         //エラーデータ配列
        int DenIndex;                   //現在の伝票添え字

        Boolean bCngFlag = false;

        public Base()
        {
            InitializeComponent();
        }

        private void Base_Load(object sender, EventArgs e)
        {
            //ウィンドウズ最小サイズ
            utility.WindowsMinSize(this, this.Size.Width, this.Size.Height);

            //ウィンドウズ最大サイズ
            utility.WindowsMaxSize(this, this.Size.Width, this.Size.Height);

            //MessageBox.Show("ウィンドウズ最大サイズ");

            //インストールディレクトリを取得する
            //global.WorkDir = utility.GetPath();
            global.WorkDir = Properties.Settings.Default.appInstPath; // 2014/09/02

            if (global.WorkDir == "")
            {
                MessageBox.Show("インストールディレクトリが取得できませんでした" + Environment.NewLine + "プログラムを終了します", "レジストリ未登録", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                Environment.Exit(0);
            }

            //MessageBox.Show("インストールディレクトリを取得する");

            //フォルダ作成
            System.IO.Directory.CreateDirectory(global.WorkDir + global.DIR_OK);
            System.IO.Directory.CreateDirectory(global.WorkDir + global.DIR_INCSV);
            System.IO.Directory.CreateDirectory(global.WorkDir + global.DIR_BREAK);

            //ファイル有無チェック
            start s = new start();
            if (s.FileExistChk(global.WorkDir) == false)
            {
                MessageBox.Show("処理を行うデータがありません", "ＮＧ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                Environment.Exit(0);
            }

            //設定データ取得
            s.InitialLoad(global.WorkDir);

            //会社選択画面
            Form frm = new frmComSelect();
            frm.ShowDialog();

            //会社情報が存在しない場合はアプリケーションを終了する
            if (global.pblDbName == string.Empty) Environment.Exit(0);

            //結果ファイル削除
            utility.FileDelete(global.WorkDir + global.DIR_OK, global.OUTFILE);

            //入力ファイルをコピー
            if (System.IO.File.Exists(global.WorkDir + global.DIR_HENKAN + global.INFILE))
            {
                utility.FileDelete(global.WorkDir + global.DIR_HENKAN, global.TMPREAD);
                System.IO.File.Copy(global.WorkDir + global.DIR_HENKAN + global.INFILE, global.WorkDir + global.DIR_HENKAN + global.TMPREAD);

                //正しくないファイルの場合、アプリケーション終了
                if (DenKindJudge(global.WorkDir + global.DIR_HENKAN + global.TMPREAD) == false)
                {
                    System.IO.File.Copy(global.WorkDir + global.DIR_HENKAN + global.INFILE, global.WorkDir + global.DIR_HENKAN + global.LOGFILE);
                    MessageBox.Show("不正な伝票ファイルです。伝票データが存在しません。", "ＮＧ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    Environment.Exit(0);
                }
            }

            //マスター情報取得
            LoadMaster();

            //処理するファイルの読込
            switch (global.pblSelFILE)
            {
                case 0:
                    //伝票データ分割(TMPREAD→CSVFILE)
                    LoadCsvDivide();
                    break;

                //中断データリカバリ
                case 1:
                    frmFilSelect2 frmFil = new frmFilSelect2();
                    frmFil.ShowDialog();
                    BreakFilesMove(frmFil);
                    frmFil.Close();
                    frmFil.Dispose();

                    break;

                default:
                    break;
            }

            //伝票データロード
            DenData = LoadDataFurikae();

            //エラーチェック処理 
            Boolean x1;　
            eTbl = ChkMainNew(DenData,out x1);

            //エラーなしで処理を終了するとき：true、エラー有りまたはエラーなしで処理を終了しないとき：false
            if (x1 == true)
            {
                MainEnd(DenData,true);   //汎用データを作成して処理を終了する
            }
            
            //初めは会社情報を表示
            //tabData.SelectedIndex = global.TAB_COM;
    
            //キャプション
            this.Text = Application.ProductName + "Ver " + Application.ProductVersion;
    
            //フォームタグ初期化
            this.Tag = string.Empty;
    
            //初期設定はエラー項目カラー表示しない
            ChkErrColor.Checked = false;
            ErrColorChange();

            //エラーグリッドのカレントセルを無効にする(行選択状態にしない）
            tabData.SelectedIndex = global.TAB_ERR;
            fgNg.CurrentCell = null;

            //multirow編集モード
            gcMultiRow1.EditMode = EditMode.EditProgrammatically;

        }

        /// <summary>
        /// 選択した日付に該当する中断ファイルを分割ファイルへ移動する
        /// </summary>
        /// <param name="frmFil">フォームオブジェクト</param>
        private void BreakFilesMove(frmFilSelect2 frmFil)
        {
            string sPath = string.Empty;

            //選択した日付に該当する中断ファイルを分割ファイルへ移動する
            foreach (var item in frmFil.st)
            {
                if (item.fFlg == true)
                {
                    //現在のパスを含む中断ファイル名
                    sPath = global.WorkDir + global.DIR_BREAK + string.Format("{0:000}", global.pblComNo) + @"\";

                    //移動先のパスを含むファイル名
                    string reFile = global.WorkDir + global.DIR_INCSV + item.fPath.Replace(sPath, string.Empty);

                    //CSVファイル移動
                    File.Move(item.fPath, reFile);

                    //画像ファイル移動
                    File.Move(item.fPath.Replace("csv", "bmp"), reFile.Replace("csv", "bmp"));
                }
            }

            //Thumbs.dbは無条件に削除する
            if (File.Exists(sPath + "Thumbs.db"))
            {
                utility.FileDelete(sPath, "Thumbs.db");
            }

            //全てのファイルをリカバリーしたときはフォルダを削除
            if (System.IO.Directory.GetFiles(sPath).Count() == 0) Directory.Delete(sPath);
        }

        /// <summary>
        /// 入力ファイルをチェックする
        /// </summary>
        /// <param name="sDsnPath">Dファイルパス名</param>
        /// <returns>true:正常ファイル、false:不正なファイル</returns>
        private Boolean DenKindJudge(String sfPath)
        {
            int sX1 = 0;

            // StreamReader の新しいインスタンスを生成する
            StreamReader cReader = (new StreamReader(sfPath, Encoding.Default));

            // 1行読む
            if (cReader.Peek() >= 0)
            {
                string stBuffer = cReader.ReadLine();
                if (stBuffer.Substring(0, 1) != "*") sX1 = 1;
            }
            else
            {
                sX1 = 1;
            }

            // cReader を閉じる
            cReader.Close();

            if (sX1 == 0)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        private void LoadMaster()
        {
            //ステータスオン
            global.MASTERLOAD_STATUS = 1;

            //オーナーフォームを無効にする
            this.Enabled = false;

            //プログレスバーを表示する
            frmPrg frmP = new frmPrg();
            frmP.Owner = this;
            frmP.Show();

            //会社データ
            frmP.Text = "会社データロード中";
            frmP.progressValue = 10;
            frmP.ProgressStep();

            GridViewSetting_Company(fgCom);     //グリッドビュー設定
            GridViewShow_company(fgCom);        //グリッドにデータ表示

            ////伝票入力指定期間
            //frmP.Text = "伝票入力指定期間ロード中";
            //frmP.progressValue = 20;
            //frmP.ProgressStep();

            //company cp = new company();
            //cp.LimitDataLoad();

            ////入力制限期間を設定
            //frmP.Text = "入力制限期間を設定中";
            //frmP.progressValue = 50;
            //frmP.ProgressStep();
            
            //SetLimit();

            //勘定科目
            frmP.Text = "勘定科目をロード中";
            frmP.progressValue = 60;
            frmP.ProgressStep();

            GridViewSetting_Kamoku(fgKamoku);   //グリッドビュー設定
            GridViewShow_Kamoku(fgKamoku);      //グリッドにデータ表示

            //補助科目
            GridViewSetting_Hojo(fgHojo);       //グリッドビュー設定

            //部門
            frmP.Text = "部門データをロード中";
            frmP.progressValue = 70;
            frmP.ProgressStep();

            GridViewSetting_Bumon(fgBumon);     //グリッドビュー設定
            GridViewShow_Bumon(fgBumon);        //グリッドにデータ表示

            //税区分
            frmP.Text = "税区分をロード中";
            frmP.progressValue = 80;
            frmP.ProgressStep();

            GridViewSetting_Tax(fgTax);         //グリッドビュー設定
            GridViewShow_Tax(fgTax);            //グリッドにデータ表示

            //税処理
            frmP.Text = "税処理をロード中";
            frmP.progressValue = 90;
            frmP.ProgressStep();

            GridViewSetting_TaxMas(fgTaxMas);   //グリッドビュー設定
            GridViewShow_TaxMas(fgTaxMas);      //グリッドにデータ表示

            //摘要
            frmP.Text = "摘要をロード中";
            frmP.progressValue = 99;
            frmP.ProgressStep();

            GridViewSetting_Tekiyo(fgTekiyo);   //グリッドビュー設定
            GridViewShow_Tekiyo(fgTekiyo);      //グリッドにデータ表示

            // いったんオーナーをアクティブにする
            this.Activate();

            // 進行状況ダイアログを閉じる
            frmP.Close();

            // オーナーのフォームを有効に戻す
            this.Enabled = true;

            //ステータスオフ
            global.MASTERLOAD_STATUS = 0;
        }

        /// <summary>
        /// データグリッドビューの定義を行います
        /// </summary>
        /// <param name="tempDGV">データグリッドビューオブジェクト</param>
        private void GridViewSetting_Company(DataGridView tempDGV)
        {
            try
            {
                //フォームサイズ定義

                // 列スタイルを変更する

                tempDGV.EnableHeadersVisualStyles = false;

                // 列ヘッダー表示位置指定
                tempDGV.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;

                // 列ヘッダーフォント指定
                tempDGV.ColumnHeadersDefaultCellStyle.Font = new Font("Meiryo UI", 9, FontStyle.Regular);

                // データフォント指定
                tempDGV.DefaultCellStyle.Font = new Font("Meiryo UI", (float)9, FontStyle.Regular);

                // 行の高さ
                tempDGV.ColumnHeadersHeight = 18;
                tempDGV.RowTemplate.Height = 18;

                // 全体の高さ
                //tempDGV.Height = 180;

                // 奇数行の色
                //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                // 各列幅指定
                tempDGV.Columns.Add("col1", "項目名");
                tempDGV.Columns.Add("col2", "摘要");

                //tempDGV.Columns[4].Visible = false; //データベース名は非表示
                //tempDGV.Columns[5].Visible = false; //税区分は非表示

                tempDGV.Columns[0].Width = 100;
                //tempDGV.Columns[1].Width = 100;
                tempDGV.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

                tempDGV.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                tempDGV.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;

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

        /// <summary>
        /// グリッドビューへ会社情報を表示する
        /// </summary>
        /// <param name="sConnect">接続文字列</param>
        /// <param name="tempDGV">DataGridViewオブジェクト名</param>
        private void GridViewShow_company(DataGridView tempDGV)
        {

            string wrkGengou;
            //string wrkKikan;
            string wrkFromYear;
            string wrkFromMonth;
            string wrkFromDay;
            string wrkToYear;
            string wrkToMonth;
            string wrkToDay;
            string wrkKaisi;

            ////勘定奉行データベースより会社情報を取得する
            //company cp = new company();
            //cp.CompDataLoad();

            //会計期間のフォーマット
            //if (company.Hosei != "0")
            if (global.pblReki == global.RWAREKI)  //和暦ならば
            {
                wrkGengou = company.Gengou;
                wrkFromYear = (int.Parse(company.FromYear) - int.Parse(company.Hosei)).ToString();
                wrkFromYear = String.Format(string.Format("{0,2}", int.Parse(wrkFromYear)));
                wrkToYear = (int.Parse(company.ToYear) - int.Parse(company.Hosei)).ToString();
                wrkToYear = String.Format("{0,2}", int.Parse(wrkToYear));
            }
            else
            {
                wrkGengou = "  ";
                wrkFromYear = company.FromYear;
                wrkToYear = company.ToYear;
            }

            wrkFromMonth = String.Format("{0,2}", int.Parse(company.FromMonth));
            wrkFromDay = String.Format("{0,2}", int.Parse(company.FromDay));
            wrkToMonth = String.Format("{0,2}", int.Parse(company.ToMonth));
            wrkToDay = String.Format("{0,2}", int.Parse(company.ToDay));

            //入力開始月フォーマット
            wrkKaisi = int.Parse(company.FromMonth).ToString();
            if (int.Parse(wrkKaisi) > 12) wrkKaisi = (int.Parse(wrkKaisi) - 12).ToString();

            //取得方法追加「税処理を取得」 (v6.0対応)--
            if (global.gsTaxMas.Trim() == "2")
            {
                company.TaxMas = "1";
            }
            else
            {
                company.TaxMas = "0";
            }

            try
            {
                //グリッドビューに表示する
                tempDGV.RowCount = 6;

                //会社名
                tempDGV[0, 0].Value = "会社名";
                tempDGV[1, 0].Value = company.Name;

                //会計期間期首
                tempDGV[0, 1].Value = "会計期間・期首";
                tempDGV[1, 1].Value = wrkFromYear + "年" + wrkFromMonth + "月" + wrkFromDay + "日";

                //会計期間期末
                tempDGV[0, 2].Value = "会計期間・期末";
                tempDGV[1, 2].Value = wrkToYear + "年" + wrkToMonth + "月" + wrkToDay + "日";

                //入力開始月
                tempDGV[0, 3].Value = "入力開始月";
                tempDGV[1, 3].Value = string.Format("{0,2}", int.Parse(wrkKaisi)) + "月";

                //中間期決算
                tempDGV[0, 4].Value = "決算回数";
                if (company.Middle == global.FLGON)
                {
                    tempDGV[1, 4].Value = "する";
                }
                else
                {
                    tempDGV[1, 4].Value = "しない";
                }

                //決算回数
                switch (company.Middle)
                {
                    case "0":
                        tempDGV[1, 4].Value = "年1回";
                        break;

                    case "1":
                        tempDGV[1, 4].Value = "年2回（中間決算）";
                        break;

                    case "2":
                        tempDGV[1, 4].Value = "年4回（四半期決算）";
                        break;

                    default:
                        tempDGV[1, 4].Value = "不明";
                        break;
                }

                //税処理
                tempDGV[0, 5].Value = "税処理";

                if (global.gsTaxMas == "0")
                {
                    tempDGV[1, 5].Value = "税抜別段";
                }
                else
                {
                    tempDGV[1, 5].Value = "税込自動";
                }

                tabData.SelectedIndex = global.TAB_COM;
                tempDGV.CurrentCell = null;

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "エラー", MessageBoxButtons.OK);
            }

        }

        /// <summary>
        /// 日付入力範囲の設定
        /// </summary>
        private void SetLimit()
        {
            int wrkLock = int.Parse(company.LmLock);
            int wrkSt = int.Parse(company.LmStSoeji);
            int wrkEd = int.Parse(company.LmEdSoeji);
            int wrkKaisi = int.Parse(company.Kaisi);

            //通常仕訳の入力期間　とりあえずマスターの指定期間を入れておく
            Limit.LimitKikan s = new Limit.LimitKikan();
            s.initialDataSet();

            //最初の四半期決算期間
            Limit.BfQuaKessanDate1 kessan1 = new Limit.BfQuaKessanDate1();
            kessan1.GetKessanDate();

            //2度目の四半期決算期間
            Limit.BfQuaKessanDate1 kessan2 = new Limit.BfQuaKessanDate1();
            kessan2.GetKessanDate();

            //3度目の四半期決算期間
            Limit.BfQuaKessanDate1 kessan3 = new Limit.BfQuaKessanDate1();
            kessan3.GetKessanDate();

            //中間期決算期間
            Limit.MidKessanDate midKessan = new Limit.MidKessanDate();
            midKessan.GetKessanDate();

            //元の中間期決算期間
            Limit.BfMidKessan bfmidKessan = new Limit.BfMidKessan();
            bfmidKessan.GetKessanDate();

            //決算期間の取得
            Limit.KessanDate kessan = new Limit.KessanDate();
            kessan.GetKessanDate();

            //元の決算期間の取得
            Limit.BfKessan bfkessan = new Limit.BfKessan();
            bfkessan.GetKessanDate();

            //使用可のフラグON
            company.LmFlag = true;
            Limit.LimitKikan.Flag = true;
            Limit.MidKessanDate.Flag = true;
            Limit.KessanDate.Flag = true;

            DateTime sDate;

            switch (wrkLock)
            {
                //入力制限なしの場合
                case 0:
                    //入力開始月が中間期決算月以降の場合
                    if (wrkKaisi > 5) Limit.MidKessanDate.Flag = false; //中間期決算期間の入力を禁止

                    //入力期間表示
                    //Call ShowLimit(pblLimitKikan, 0, pblLimitKikan, 0)
                    break;

                //指定期間を入力禁止
                case 1:
                    if ((0 <= wrkEd) && (wrkEd <= 5))
                    {
                        //通常仕訳　指定期間の翌日から期末まで
                        sDate = DateTime.Parse(company.LmToYear + "/" + company.LmToMonth + "/" + company.LmToDay);
                        Limit.LimitKikan.FromYear = new Limit.GetNextDay(sDate).GetYear();
                        Limit.LimitKikan.FromMonth = new Limit.GetNextDay(sDate).GetMonth();
                        Limit.LimitKikan.FromDay = new Limit.GetNextDay(sDate).GetDay();
                        Limit.LimitKikan.ToYear = company.ToYear;
                        Limit.LimitKikan.ToMonth = company.ToMonth;
                        Limit.LimitKikan.ToDay = company.ToDay;

                        ////入力期間表示
                        //Call ShowLimit(pblLimitKikan, 0, pblKessanDate, 2)
                    }
                    else
                    {
                        if (wrkEd == 6)
                        {
                            //指定期間終了日が、中間期決算期間の終了日より後でなければ通常処理
                            if (JudgeDate(DateTime.Parse(company.LmToYear + "/" + company.LmToMonth + "/" + company.LmToDay), DateTime.Parse(Limit.MidKessanDate.ToYear + "/" + Limit.MidKessanDate.ToMonth + "/" + Limit.MidKessanDate.ToDay)))
                            {
                                //通常仕訳　中間期決算期間の翌日から期末まで
                                sDate = DateTime.Parse(Limit.MidKessanDate.ToYear + "/" + Limit.MidKessanDate.ToMonth + "/" + Limit.MidKessanDate.ToDay);
                                Limit.LimitKikan.FromYear = new Limit.GetNextDay(sDate).GetYear();
                                Limit.LimitKikan.FromMonth = new Limit.GetNextDay(sDate).GetMonth();
                                Limit.LimitKikan.FromDay = new Limit.GetNextDay(sDate).GetDay();
                                Limit.LimitKikan.ToYear = company.ToYear;
                                Limit.LimitKikan.ToMonth = company.ToMonth;
                                Limit.LimitKikan.ToDay = company.ToDay;

                                //中間期決算期間　指定期間の翌日から
                                sDate = DateTime.Parse(company.LmToYear + "/" + company.LmToMonth + "/" + company.LmToDay);
                                Limit.MidKessanDate.FromYear = new Limit.GetNextDay(sDate).GetYear();
                                Limit.MidKessanDate.FromMonth = new Limit.GetNextDay(sDate).GetMonth();
                                Limit.MidKessanDate.FromDay = new Limit.GetNextDay(sDate).GetDay();

                                //' 入力期間表示
                                //Call ShowLimit(pblMidKessanDate, 1, pblKessanDate, 2)
                            }
                            else
                            {
                                //通常仕訳　指定期間の翌日から期末まで
                                sDate = DateTime.Parse(company.LmToYear + "/" + company.LmToMonth + "/" + company.LmToDay);
                                Limit.LimitKikan.FromYear = new Limit.GetNextDay(sDate).GetYear();
                                Limit.LimitKikan.FromMonth = new Limit.GetNextDay(sDate).GetMonth();
                                Limit.LimitKikan.FromDay = new Limit.GetNextDay(sDate).GetDay();
                                Limit.LimitKikan.ToYear = company.ToYear;
                                Limit.LimitKikan.ToMonth = company.ToMonth;
                                Limit.LimitKikan.ToDay = company.ToDay;

                                //中間期決算を使用禁止
                                Limit.MidKessanDate.Flag = false;

                                ////入力期間表示
                                //Call ShowLimit(wrkNextDay, 0, pblKessanDate, 2)

                            }
                        }
                        else
                        {
                            if (7 <= wrkEd && wrkEd <= 12)
                            {
                                //通常仕訳　指定期間の翌日から期末まで
                                sDate = DateTime.Parse(company.LmToYear + "/" + company.LmToMonth + "/" + company.LmToDay);
                                Limit.LimitKikan.FromYear = new Limit.GetNextDay(sDate).GetYear();
                                Limit.LimitKikan.FromMonth = new Limit.GetNextDay(sDate).GetMonth();
                                Limit.LimitKikan.FromDay = new Limit.GetNextDay(sDate).GetDay();
                                Limit.LimitKikan.ToYear = company.ToYear;
                                Limit.LimitKikan.ToMonth = company.ToMonth;
                                Limit.LimitKikan.ToDay = company.ToDay;

                                //中間期決算を使用禁止
                                Limit.MidKessanDate.Flag = false;

                                ////入力期間表示
                                //Call ShowLimit(pblLimitKikan, 0, pblKessanDate, 2)
                            }
                            else
                            {
                                if (wrkEd == 13)
                                {
                                    //通常仕訳の使用禁止
                                    Limit.LimitKikan.Flag = false;

                                    //中間期決算の使用禁止
                                    Limit.MidKessanDate.Flag = false;

                                    //指定範囲末と期末が同じ場合
                                    if (company.LmToDay == company.ToDay)
                                    {
                                        //決算の使用禁止
                                        Limit.KessanDate.Flag = false;
                                    }
                                    else
                                    {
                                        //決算期間　指定期間の翌日から
                                        sDate = DateTime.Parse(company.LmToYear + "/" + company.LmToMonth + "/" + company.LmToDay);
                                        Limit.KessanDate.FromYear = new Limit.GetNextDay(sDate).GetYear();
                                        Limit.KessanDate.FromMonth = new Limit.GetNextDay(sDate).GetMonth();
                                        Limit.KessanDate.FromDay = new Limit.GetNextDay(sDate).GetDay();
                                    }

                                    ////入力期間表示
                                    //Call ShowLimit(pblKessanDate, 2, pblKessanDate, 2)
                                }
                            }
                        }
                    }
                    break;

                //指定期間のみ入力許可
                case 2:
                    if (0 <= wrkSt && wrkSt <= 5)
                    {
                        if (0 <= wrkEd && wrkEd <= 5)
                        {
                            //中間期決算の使用禁止
                            Limit.MidKessanDate.Flag = false;

                            //決算の使用禁止
                            Limit.KessanDate.Flag = false;

                            ////入力期間表示
                            //Call ShowLimit(pblLimitKikan, 0, pblLimitKikan, 0)
                        }
                        else
                        {
                            if (wrkEd == 6)
                            {
                                //通常仕訳　現時点の中間期決算末まで
                                Limit.LimitKikan.ToYear = Limit.MidKessanDate.ToYear;
                                Limit.LimitKikan.ToMonth = Limit.MidKessanDate.ToMonth;
                                Limit.LimitKikan.ToDay = Limit.MidKessanDate.ToDay;

                                //指定期間終了日が、中間期決算期間の終了日より後でなければ通常処理
                                if (JudgeDate(DateTime.Parse(company.LmToYear + "/" + company.LmToMonth + "/" + company.LmToDay), DateTime.Parse(Limit.MidKessanDate.FromYear + "/" + Limit.MidKessanDate.FromMonth + "/" + Limit.MidKessanDate.FromDay)))
                                {
                                    //中間期決算期間　指定期間まで
                                    Limit.MidKessanDate.ToYear = company.LmToYear;
                                    Limit.MidKessanDate.ToMonth = company.LmToMonth;
                                    Limit.MidKessanDate.ToDay = company.LmToDay;
                                }

                                //決算の使用禁止
                                Limit.KessanDate.Flag = false;

                                ////入力期間表示
                                //Call ShowLimit(pblLimitKikan, 0, pblMidKessanDate, 1)
                            }
                            else
                            {
                                if (7 <= wrkEd && wrkEd <= 12)
                                {
                                    //決算の使用禁止
                                    Limit.KessanDate.Flag = false;

                                    ////入力期間表示
                                    //Call ShowLimit(pblLimitKikan, 0, pblLimitKikan, 0)
                                }
                                else
                                {
                                    if (wrkEd == 13)
                                    {
                                        //通常仕訳　期末まで
                                        Limit.LimitKikan.ToYear = company.ToYear;
                                        Limit.LimitKikan.ToMonth = company.ToMonth;
                                        Limit.LimitKikan.ToDay = company.ToDay;

                                        //決算期間　指定期間まで
                                        Limit.KessanDate.ToYear = company.LmToYear;
                                        Limit.KessanDate.ToMonth = company.LmToMonth;
                                        Limit.KessanDate.ToDay = company.LmToDay;

                                        ////入力期間表示
                                        //Call ShowLimit(pblLimitKikan, 0, pblKessanDate, 2)
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        if (wrkSt == 6)
                        {
                            if (wrkEd == 6)
                            {
                                //通常仕訳の使用禁止
                                Limit.LimitKikan.Flag = false;

                                //指定期間開始日が、中間期決算期間の開始日より前でなければ通常処理
                                if (JudgeDate(DateTime.Parse(Limit.MidKessanDate.FromYear + "/" +
                                                             Limit.MidKessanDate.FromMonth + "/" +
                                                             Limit.MidKessanDate.FromDay),
                                              DateTime.Parse(company.LmFromYear + "/" +
                                                             company.LmFromMonth + "/" +
                                                             company.LmFromDay)))
                                {
                                    //中間期決算期間の開始日 = 指定期間開始日
                                    Limit.MidKessanDate.FromYear = company.LmFromYear;
                                    Limit.MidKessanDate.FromMonth = company.LmFromMonth;
                                    Limit.MidKessanDate.FromDay = company.LmFromDay;
                                }

                                //指定期間終了日が、中間期決算期間の終了日より後でなければ通常処理
                                if (JudgeDate(DateTime.Parse(company.LmToYear + "/" +
                                                             company.LmToMonth + "/" +
                                                             company.LmToDay),
                                              DateTime.Parse(Limit.MidKessanDate.ToYear + "/" +
                                                             Limit.MidKessanDate.ToMonth + "/" +
                                                             Limit.MidKessanDate.ToDay)))
                                {

                                    //中間期決算期間の終了日 = 指定期間終了日
                                    Limit.MidKessanDate.ToYear = company.LmToYear;
                                    Limit.MidKessanDate.ToMonth = company.LmToMonth;
                                    Limit.MidKessanDate.ToDay = company.LmToDay;

                                }

                                //決算の使用禁止
                                Limit.KessanDate.Flag = false;

                                ////入力期間表示
                                //Call ShowLimit(pblMidKessanDate, 1, pblMidKessanDate, 1)
                            }
                        }
                        else
                        {
                            if (7 <= wrkEd && wrkEd <= 12)
                            {
                                //通常仕訳　中間期決算期間の翌日から指定期間末まで
                                sDate = DateTime.Parse(company.LmToYear + "/" + company.LmToMonth + "/" + company.LmToDay);
                                Limit.LimitKikan.FromYear = new Limit.GetNextDay(sDate).GetYear();
                                Limit.LimitKikan.FromMonth = new Limit.GetNextDay(sDate).GetMonth();
                                Limit.LimitKikan.FromDay = new Limit.GetNextDay(sDate).GetDay();

                                //指定期間開始日が、中間期決算期間の開始日より前でなければ通常処理
                                if (JudgeDate(DateTime.Parse(Limit.MidKessanDate.FromYear + "/" +
                                                             Limit.MidKessanDate.FromMonth + "/" +
                                                             Limit.MidKessanDate.FromDay),
                                              DateTime.Parse(company.LmFromYear + "/" +
                                                             company.LmFromMonth + "/" +
                                                             company.LmFromDay)))
                                {
                                    //中間期決算期間の開始日 = 指定期間開始日
                                    Limit.MidKessanDate.FromYear = company.LmFromYear;
                                    Limit.MidKessanDate.FromMonth = company.LmFromMonth;
                                    Limit.MidKessanDate.FromDay = company.LmFromDay;
                                }

                                //決算の使用禁止
                                Limit.KessanDate.Flag = false;

                                ////入力期間表示
                                //Call ShowLimit(pblMidKessanDate, 1, pblLimitKikan, 0)
                            }
                            else
                            {
                                if (wrkEd == 13)
                                {
                                    //通常仕訳　中間期決算期間の翌日から期末まで
                                    sDate = DateTime.Parse(Limit.MidKessanDate.ToYear + "/" +
                                                           Limit.MidKessanDate.ToMonth + "/" +
                                                           Limit.MidKessanDate.ToDay);

                                    Limit.LimitKikan.FromYear = new Limit.GetNextDay(sDate).GetYear();
                                    Limit.LimitKikan.FromMonth = new Limit.GetNextDay(sDate).GetMonth();
                                    Limit.LimitKikan.FromDay = new Limit.GetNextDay(sDate).GetDay();
                                    Limit.LimitKikan.ToYear = company.ToYear;
                                    Limit.LimitKikan.ToMonth = company.ToMonth;
                                    Limit.LimitKikan.ToDay = company.ToDay;

                                    //指定期間開始日が、中間期決算期間の開始日より前でなければ通常処理
                                    if (JudgeDate(DateTime.Parse(Limit.MidKessanDate.FromYear + "/" +
                                                                 Limit.MidKessanDate.FromMonth + "/" +
                                                                 Limit.MidKessanDate.FromDay),
                                                  DateTime.Parse(company.LmFromYear + "/" +
                                                                 company.LmFromMonth + "/" +
                                                                 company.LmFromDay)))
                                    {
                                        //中間期決算期間　指定期間から
                                        Limit.MidKessanDate.FromYear = company.LmFromYear;
                                        Limit.MidKessanDate.FromMonth = company.LmFromMonth;
                                        Limit.MidKessanDate.FromDay = company.LmFromDay;
                                    }

                                    //決算期間　指定期間まで
                                    Limit.KessanDate.ToYear = company.LmToYear;
                                    Limit.KessanDate.ToMonth = company.LmToMonth;
                                    Limit.KessanDate.ToDay = company.LmToDay;

                                    ////入力期間表示
                                    //Call ShowLimit(pblMidKessanDate, 1, pblKessanDate, 2)

                                }
                                else
                                {
                                    if (7 <= wrkSt && wrkSt <= 12)
                                    {
                                        if (7 <= wrkEd && wrkEd <= 12)
                                        {
                                            //中間期決算の使用禁止
                                            Limit.MidKessanDate.Flag = false;

                                            //決算の使用禁止
                                            Limit.KessanDate.Flag = false;

                                            ////入力期間表示
                                            //Call ShowLimit(pblLimitKikan, 0, pblLimitKikan, 0)
                                        }
                                        else
                                        {
                                            if (wrkEd == 13)
                                            {
                                                //通常仕訳　指定期間開始から期末まで
                                                Limit.LimitKikan.ToYear = company.ToYear;
                                                Limit.LimitKikan.ToMonth = company.ToMonth;
                                                Limit.LimitKikan.ToDay = company.ToDay;

                                                //中間期決算の使用禁止
                                                Limit.MidKessanDate.Flag = false;

                                                //決算期間　指定期間まで
                                                Limit.KessanDate.ToYear = company.LmToYear;
                                                Limit.KessanDate.ToMonth = company.LmToMonth;
                                                Limit.KessanDate.ToDay = company.LmToDay;

                                                ////入力期間表示
                                                //Call ShowLimit(pblLimitKikan, 0, pblKessanDate, 2)
                                            }
                                            else
                                            {
                                                if (wrkSt == 13 && wrkEd == 13)
                                                {
                                                    //通常仕訳の使用禁止
                                                    Limit.LimitKikan.Flag = false;

                                                    //中間期決算の使用禁止
                                                    Limit.MidKessanDate.Flag = false;

                                                    //指定期間開始日が、中間期決算期間の開始日より前でなければ通常処理
                                                    if (JudgeDate(DateTime.Parse(Limit.KessanDate.FromYear + "/" +
                                                                                 Limit.KessanDate.FromMonth + "/" +
                                                                                 Limit.KessanDate.FromDay),
                                                                  DateTime.Parse(company.LmFromYear + "/" +
                                                                                 company.LmFromMonth + "/" +
                                                                                 company.LmFromDay)))
                                                    {
                                                        //決算期間 = 指定期間
                                                        Limit.KessanDate.FromYear = company.LmFromYear;
                                                        Limit.KessanDate.FromMonth = company.LmFromMonth;
                                                        Limit.KessanDate.FromDay = company.LmFromDay;
                                                        Limit.KessanDate.StSoeji = company.LmStSoeji;
                                                        Limit.KessanDate.ToYear = company.LmToYear;
                                                        Limit.KessanDate.ToMonth = company.LmToMonth;
                                                        Limit.KessanDate.ToDay = company.LmToDay;
                                                        Limit.KessanDate.EdSoeji = company.LmEdSoeji;
                                                        Limit.KessanDate.Flag = company.LmFlag;
                                                        Limit.KessanDate.Lock = company.LmLock;

                                                    }
                                                    ////入力期間表示
                                                    //Call ShowLimit(pblKessanDate, 2, pblKessanDate, 2)
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }

                    break;


                default:
                    break;
            }
        }

        /// <summary>
        /// 勘定科目データグリッドビューの定義を行います
        /// </summary>
        /// <param name="tempDGV">科目データグリッドビューオブジェクト</param>
        private void GridViewSetting_Kamoku(DataGridView tempDGV)
        {
            try
            {
                //フォームサイズ定義

                // 列スタイルを変更する

                tempDGV.EnableHeadersVisualStyles = false;

                // 列ヘッダー表示位置指定
                tempDGV.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;

                // 列ヘッダーフォント指定
                tempDGV.ColumnHeadersDefaultCellStyle.Font = new Font("Meiryo UI", 9, FontStyle.Regular);

                // データフォント指定
                tempDGV.DefaultCellStyle.Font = new Font("Meiryo UI", (float)9, FontStyle.Regular);

                // 行の高さ
                tempDGV.ColumnHeadersHeight = 18;
                tempDGV.RowTemplate.Height = 18;

                // 全体の高さ
                //tempDGV.Height = 180;

                // 奇数行の色
                //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                // 各列幅指定
                tempDGV.Columns.Add("col1", "コード");
                tempDGV.Columns.Add("col2", "勘定科目名");
                tempDGV.Columns.Add("col3", "勘定科目内部コード");
                tempDGV.Columns.Add("col4", "");
                tempDGV.Columns.Add("col5", "");

                tempDGV.Columns[2].Visible = false; //データベース名は非表示
                tempDGV.Columns[3].Visible = false; //データベース名は非表示
                tempDGV.Columns[4].Visible = false; //税区分は非表示

                tempDGV.Columns[0].Width = 60;
                //tempDGV.Columns[1].Width = 100;
                tempDGV.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

                tempDGV.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                tempDGV.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;

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

        ///------------------------------------------------------------------------
        /// <summary>
        ///     グリッドビューへ勘定科目を表示する : 2017/09/11 勘定奉行i10 </summary>
        /// <param name="tempDGV">DataGridViewオブジェクト名
        ///     </param>
        ///------------------------------------------------------------------------
        private void GridViewShow_Kamoku(DataGridView tempDGV)
        {
            //勘定奉行データベース接続文字列を取得する 2017/06/06
            string sc = SqlControl.obcConnectSting.get(global.pblDbName);

            //勘定奉行データベースへ接続する
            SqlControl.DataControl dCon = new SqlControl.DataControl(sc);

            //科目データ取得
            //データリーダーを取得する
            SqlDataReader dR;

            //string sqlSTRING = "SELECT sUcd,sNcd,sNm,tiIsTrk,tiIsZei FROM wkskm01 WHERE tiIsTrk = 1 ORDER BY sUcd";

            string sqlSTRING = string.Empty;
            sqlSTRING += "SELECT tbAccountItem.AccountItemID, tbAccountItem.AccountItemCode, tbAccountItem.AccountItemName, ";
            sqlSTRING += "tbAccountItem.IsUse, tbAccountItemAndConsumptionTaxDivisionRelation.AutomaticCalculationTax ";
            sqlSTRING += "FROM tbAccountItem inner join tbAccountItemAndConsumptionTaxDivisionRelation ";
            sqlSTRING += "on tbAccountItem.AccountItemID = tbAccountItemAndConsumptionTaxDivisionRelation.AccountItemID ";
            sqlSTRING += "WHERE (tbAccountItem.IsUse = 1) and (tbAccountItem.AccountingPeriodID = " + global.pblAccPID + ") ";
            sqlSTRING += "ORDER BY tbAccountItem.AccountItemCode";

            dR = dCon.free_dsReader(sqlSTRING);

            try
            {
                int iX = 0;

                //グリッドビューに表示する
                tempDGV.RowCount = 0;

                while (dR.Read())
                {
                    tempDGV.Rows.Add();

                    //コード
                    if (dR["AccountItemCode"].ToString().Trim().Length > global.LEN_KAMOKU)
                    {
                        tempDGV[0, iX].Value = dR["AccountItemCode"].ToString().Trim().Substring(dR["AccountItemCode"].ToString().Trim().Length - global.LEN_KAMOKU, global.LEN_KAMOKU);
                    }
                    else
                    {
                        tempDGV[0, iX].Value = dR["AccountItemCode"].ToString().Trim();
                    }

                    tempDGV[1, iX].Value = dR["AccountItemName"].ToString().Trim();     //名称
                    tempDGV[2, iX].Value = dR["AccountItemID"].ToString().Trim();       //勘定科目内部コード
                    tempDGV[3, iX].Value = dR["IsUse"].ToString();                      //
                    tempDGV[4, iX].Value = dR["AutomaticCalculationTax"].ToString();    //税

                    iX++;
                }

                dR.Close();
                dCon.Close();

                tabData.SelectedIndex = global.TAB_KAMOKU;
                tempDGV.CurrentCell = null;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "エラー", MessageBoxButtons.OK);
            }
        }

        ///----------------------------------------------------------------
        /// <summary>
        ///     補助科目データグリッドビューの定義を行います </summary>
        /// <param name="tempDGV">
        ///     補助科目データグリッドビューオブジェクト</param>
        ///----------------------------------------------------------------
        private void GridViewSetting_Hojo(DataGridView tempDGV)
        {
            try
            {
                //フォームサイズ定義

                // 列スタイルを変更する

                tempDGV.EnableHeadersVisualStyles = false;

                // 列ヘッダー表示位置指定
                tempDGV.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;

                // 列ヘッダーフォント指定
                tempDGV.ColumnHeadersDefaultCellStyle.Font = new Font("Meiryo UI", 9, FontStyle.Regular);

                // データフォント指定
                tempDGV.DefaultCellStyle.Font = new Font("Meiryo UI", (float)9, FontStyle.Regular);

                // 行の高さ
                tempDGV.ColumnHeadersHeight = 18;
                tempDGV.RowTemplate.Height = 18;

                // 全体の高さ
                //tempDGV.Height = 180;

                // 奇数行の色
                //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                // 各列幅指定
                tempDGV.Columns.Add("col1", "コード");
                tempDGV.Columns.Add("col2", "補助科目名");

                //tempDGV.Columns[3].Visible = false; //データベース名は非表示
                //tempDGV.Columns[4].Visible = false; //データベース名は非表示
                //tempDGV.Columns[5].Visible = false; //税区分は非表示

                tempDGV.Columns[0].Width = 60;
                //tempDGV.Columns[1].Width = 100;
                tempDGV.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

                tempDGV.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                tempDGV.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;

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

        ///---------------------------------------------------------------
        /// <summary>
        ///     グリッドビューへ補助科目を表示する : 2017/06/06 </summary>
        /// <param name="tempDGV">
        ///     DataGridViewオブジェクト名</param>
        ///---------------------------------------------------------------
        private void GridViewShow_Hojo(DataGridView tempDGV, string tempNcd)
        {
            string KanjoCode = string.Empty;
            string sonotaName = string.Empty;

            //勘定奉行データベース接続文字列を取得する
            string sc = SqlControl.obcConnectSting.get(global.pblDbName);

            //勘定奉行データベースへ接続する
            SqlControl.DataControl dCon = new SqlControl.DataControl(sc);

            //補助データ取得
            //データリーダーを取得する
            SqlDataReader dR;

            //勘定科目取得
            if (utility.NumericCheck(tempNcd))
            {
                KanjoCode = string.Format("{0:D10}", int.Parse(tempNcd));
            }
            else
            {
                KanjoCode = tempNcd;
            }

            string sqlSTRING = string.Empty;

            //補助コードがあるか？
            //sqlSTRING += "select sNcd,sUcd,sHjoUcd,wkhjm01.sNm ";
            //sqlSTRING += "from wkskm01 inner join wkhjm01 ";
            //sqlSTRING += "on wkskm01.sNcd = wkhjm01.sSknNcd ";
            //sqlSTRING += "where sUcd = '" + string.Format("{0,6}", tempNcd) + "'";

            sqlSTRING += "select tbAccountItem.AccountItemID,tbAccountItem.AccountItemCode,";
            sqlSTRING += "tbAccountItem.AccountItemName,tbSubAccountItem.SubAccountItemCode,";
            sqlSTRING += "tbSubAccountItem.SubAccountItemName ";
            sqlSTRING += "from tbAccountItem inner join tbSubAccountItem ";
            sqlSTRING += "on tbAccountItem.AccountItemID = tbSubAccountItem.AccountItemID ";
            sqlSTRING += "where tbAccountItem.AccountingPeriodID = " + global.pblAccPID + " and ";
            sqlSTRING += "tbAccountItem.AccountItemCode = '" + KanjoCode + "'";

            dR = dCon.free_dsReader(sqlSTRING);

            try
            {
                int iX = 0;

                //グリッドビューに表示する
                tempDGV.RowCount = 0;

                while (dR.Read())
                {
                    //最初のデータがコード「0」のときスキップする
                    if (iX == 0 && dR["SubAccountItemCode"].ToString().Trim() == "0000000000")
                    {
                        sonotaName = dR["SubAccountItemName"].ToString().Trim();
                    }
                    else
                    {
                        tempDGV.Rows.Add();

                        //コード
                        if (dR["SubAccountItemCode"].ToString().Trim().Length > global.LEN_HOJO)
                        {
                            tempDGV[0, iX].Value = dR["SubAccountItemCode"].ToString().Substring(dR["SubAccountItemCode"].ToString().Length - global.LEN_HOJO, global.LEN_HOJO);
                        }
                        else
                        {
                            tempDGV[0, iX].Value = dR["SubAccountItemCode"].ToString().Trim();
                        }

                        tempDGV[1, iX].Value = dR["SubAccountItemName"].ToString().Trim();  //名称

                        iX++;
                    }
                }

                dR.Close();
                dCon.Close();

                tempDGV.CurrentCell = null;

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "エラー", MessageBoxButtons.OK);
            }
        }

        /// <summary>
        /// 部門データグリッドビューの定義を行います
        /// </summary>
        /// <param name="tempDGV">部門データグリッドビューオブジェクト</param>
        private void GridViewSetting_Bumon(DataGridView tempDGV)
        {
            try
            {
                //フォームサイズ定義

                // 列スタイルを変更する

                tempDGV.EnableHeadersVisualStyles = false;

                // 列ヘッダー表示位置指定
                tempDGV.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;

                // 列ヘッダーフォント指定
                tempDGV.ColumnHeadersDefaultCellStyle.Font = new Font("Meiryo UI", 9, FontStyle.Regular);

                // データフォント指定
                tempDGV.DefaultCellStyle.Font = new Font("Meiryo UI", (float)9, FontStyle.Regular);

                // 行の高さ
                tempDGV.ColumnHeadersHeight = 18;
                tempDGV.RowTemplate.Height = 18;

                // 全体の高さ
                //tempDGV.Height = 180;

                // 奇数行の色
                //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                // 各列幅指定
                tempDGV.Columns.Add("col1", "コード");
                tempDGV.Columns.Add("col2", "部門名");

                //tempDGV.Columns[3].Visible = false; //データベース名は非表示
                //tempDGV.Columns[4].Visible = false; //データベース名は非表示
                //tempDGV.Columns[5].Visible = false; //税区分は非表示

                tempDGV.Columns[0].Width = 60;
                //tempDGV.Columns[1].Width = 100;
                tempDGV.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

                tempDGV.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                tempDGV.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;

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

        ///----------------------------------------------------------------
        /// <summary>
        ///     グリッドビューへ部門を表示する : 2017/06/06 勘定奉行i10</summary>
        /// <param name="tempDGV">
        ///     DataGridViewオブジェクト名</param>
        ///----------------------------------------------------------------
        private void GridViewShow_Bumon(DataGridView tempDGV)
        {
            //勘定奉行データベース接続文字列を取得する
            string sc = SqlControl.obcConnectSting.get(global.pblDbName);

            //勘定奉行データベースへ接続する
            SqlControl.DataControl dCon = new SqlControl.DataControl(sc);

            //部門データ取得
            //データリーダーを取得する
            SqlDataReader dR;

            //string sqlSTRING = "SELECT sUcd,sNm FROM wkbnm01 ORDER BY sUcd";

            string sqlSTRING = string.Empty;
            sqlSTRING += "select DepartmentID, DepartmentCode, DepartmentName from tbDepartment ";
            sqlSTRING += "order by DepartmentCode";

            dR = dCon.free_dsReader(sqlSTRING);

            try
            {
                int iX = 0;
                string sSonota = string.Empty;
                global.pblBumonFlg = false;

                //グリッドビューに表示する
                tempDGV.RowCount = 0;

                while (dR.Read())
                {
                    if (dR["DepartmentCode"].ToString() != "000000000000000")    //その他以外
                    {
                        tempDGV.Rows.Add();
                        //コード
                        if (dR["DepartmentCode"].ToString().Trim().Length > global.LEN_BUMON)
                        {
                            tempDGV[0, iX].Value = dR["DepartmentCode"].ToString().Trim().Substring(dR["DepartmentCode"].ToString().Trim().Length - global.LEN_BUMON, global.LEN_BUMON);
                        }
                        else
                        {
                            tempDGV[0, iX].Value = dR["DepartmentCode"].ToString().Trim();
                        }

                        tempDGV[1, iX].Value = dR["DepartmentName"].ToString().Trim();      //名称

                        iX++;

                        if (global.pblBumonFlg == false) global.pblBumonFlg = true;
                    }
                    else
                    {
                        sSonota = dR["DepartmentName"].ToString().Trim();     //名称
                    }
                }

                dR.Close();

                //その他取得
                sqlSTRING = string.Empty;
                sqlSTRING += "select DepartmentID,DepartmentCode,DepartmentName from tbDepartment ";
                sqlSTRING += "where DepartmentCode = '000000000000000' ";
                sqlSTRING += "order by DepartmentCode";

                dR = dCon.free_dsReader(sqlSTRING);

                while (dR.Read())
                {
                    sSonota = dR["DepartmentName"].ToString().Trim();     //名称
                }

                dR.Close();
                dCon.Close();

                //部門データありなら最終行に「その他」追加
                if (global.pblBumonFlg == true)
                {
                    tempDGV.Rows.Add();
                    tempDGV[0, iX].Value = "0";         //コード
                    tempDGV[1, iX].Value = sSonota;     //名称
                }

                tabData.SelectedIndex = global.TAB_TAXBUMON;
                tempDGV.CurrentCell = null;

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "エラー", MessageBoxButtons.OK);
            }
        }

        /// <summary>
        /// 税区分データグリッドビューの定義を行います
        /// </summary>
        /// <param name="tempDGV">税区分データグリッドビューオブジェクト</param>
        private void GridViewSetting_Tax(DataGridView tempDGV)
        {
            try
            {
                //フォームサイズ定義

                // 列スタイルを変更する

                tempDGV.EnableHeadersVisualStyles = false;

                // 列ヘッダー表示位置指定
                tempDGV.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;

                // 列ヘッダーフォント指定
                tempDGV.ColumnHeadersDefaultCellStyle.Font = new Font("Meiryo UI", 9, FontStyle.Regular);

                // データフォント指定
                tempDGV.DefaultCellStyle.Font = new Font("Meiryo UI", (float)9, FontStyle.Regular);

                // 行の高さ
                tempDGV.ColumnHeadersHeight = 18;
                tempDGV.RowTemplate.Height = 18;

                // 全体の高さ
                //tempDGV.Height = 180;

                // 奇数行の色
                //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                // 各列幅指定
                tempDGV.Columns.Add("col1", "コード");
                tempDGV.Columns.Add("col2", "税区分名");

                //tempDGV.Columns[3].Visible = false; //データベース名は非表示
                //tempDGV.Columns[4].Visible = false; //データベース名は非表示
                //tempDGV.Columns[5].Visible = false; //税区分は非表示

                tempDGV.Columns[0].Width = 60;
                //tempDGV.Columns[1].Width = 100;
                tempDGV.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

                tempDGV.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                tempDGV.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;

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

        ///--------------------------------------------------------------------
        /// <summary>
        ///     グリッドビューへ税区分を表示する : 2017/06/06  勘定奉行i10</summary>
        /// <param name="tempDGV">
        ///     DataGridViewオブジェクト名</param>
        ///--------------------------------------------------------------------
        private void GridViewShow_Tax(DataGridView tempDGV)
        {
            //勘定奉行データベース接続文字列を取得する
            string sc = SqlControl.obcConnectSting.get(global.pblDbName);

            //勘定奉行データベースへ接続する
            SqlControl.DataControl dCon = new SqlControl.DataControl(sc);

            //税区分データ取得
            //データリーダーを取得する
            SqlDataReader dR;
            //string sqlSTRING = "SELECT tiZeiCd,sZeiNm FROM wktax01 ORDER BY tiZeiCd";
            
            string sqlSTRING = string.Empty;
            sqlSTRING += "select TaxDivisionCode,TaxDivisionName from tbTaxDivision ";
            sqlSTRING += "WHERE AccountingPeriodID = " + global.pblAccPID + " ";
            sqlSTRING += "ORDER BY TaxDivisionCode";

            dR = dCon.free_dsReader(sqlSTRING);
            
            try
            {
                int iX = 0;

                //グリッドビューに表示する
                tempDGV.RowCount = 0;

                while (dR.Read())
                {
                    tempDGV.Rows.Add();
                    tempDGV[0, iX].Value = dR["TaxDivisionCode"].ToString().Trim();    //コード
                    tempDGV[1, iX].Value = dR["TaxDivisionName"].ToString().Trim();     //名称

                    iX++;
                }

                dR.Close();
                dCon.Close();

                tabData.SelectedIndex = global.TAB_TAXBUMON;
                tempDGV.CurrentCell = null;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "エラー", MessageBoxButtons.OK);
            }
        }

        /// <summary>
        /// 税区分データグリッドビューの定義を行います
        /// </summary>
        /// <param name="tempDGV">税区分データグリッドビューオブジェクト</param>
        private void GridViewSetting_TaxMas(DataGridView tempDGV)
        {
            try
            {
                //フォームサイズ定義

                // 列スタイルを変更する

                tempDGV.EnableHeadersVisualStyles = false;

                // 列ヘッダー表示位置指定
                tempDGV.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;

                // 列ヘッダーフォント指定
                tempDGV.ColumnHeadersDefaultCellStyle.Font = new Font("Meiryo UI", 9, FontStyle.Regular);

                // データフォント指定
                tempDGV.DefaultCellStyle.Font = new Font("Meiryo UI", (float)9, FontStyle.Regular);

                // 行の高さ
                tempDGV.ColumnHeadersHeight = 18;
                tempDGV.RowTemplate.Height = 18;

                // 全体の高さ
                //tempDGV.Height = 180;

                // 奇数行の色
                //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                // 各列幅指定
                tempDGV.Columns.Add("col1", "コード");
                tempDGV.Columns.Add("col2", "税処理名");

                //tempDGV.Columns[3].Visible = false; //データベース名は非表示
                //tempDGV.Columns[4].Visible = false; //データベース名は非表示
                //tempDGV.Columns[5].Visible = false; //税区分は非表示

                tempDGV.Columns[0].Width = 60;
                //tempDGV.Columns[1].Width = 100;
                tempDGV.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

                tempDGV.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                tempDGV.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;

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

        ///-----------------------------------------------------------------
        /// <summary>
        ///     グリッドビューへ税処理を表示する </summary>
        /// <param name="tempDGV">
        ///     DataGridViewオブジェクト名</param>
        ///-----------------------------------------------------------------
        private void GridViewShow_TaxMas(DataGridView tempDGV)
        {
            try
            {
                //グリッドビューに表示する
                tempDGV.RowCount = 2;

                //消費税計算区分をセット
                tempDGV[0, 0].Value = "0";
                tempDGV[1, 0].Value = "税抜別段";
                tempDGV[0, 1].Value = "1";
                tempDGV[1, 1].Value = "税込自動";

                tabData.SelectedIndex = global.TAB_TAXBUMON;
                tempDGV.CurrentCell = null;

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "エラー", MessageBoxButtons.OK);
            }
        }

        /// <summary>
        /// 摘要データグリッドビューの定義を行います
        /// </summary>
        /// <param name="tempDGV">摘要データグリッドビューオブジェクト</param>
        private void GridViewSetting_Tekiyo(DataGridView tempDGV)
        {
            try
            {
                //フォームサイズ定義

                // 列スタイルを変更する

                tempDGV.EnableHeadersVisualStyles = false;

                // 列ヘッダー表示位置指定
                tempDGV.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;

                // 列ヘッダーフォント指定
                tempDGV.ColumnHeadersDefaultCellStyle.Font = new Font("Meiryo UI", 9, FontStyle.Regular);

                // データフォント指定
                tempDGV.DefaultCellStyle.Font = new Font("Meiryo UI", (float)9, FontStyle.Regular);

                // 行の高さ
                tempDGV.ColumnHeadersHeight = 18;
                tempDGV.RowTemplate.Height = 18;

                // 全体の高さ
                //tempDGV.Height = 180;

                // 奇数行の色
                //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                // 各列幅指定
                tempDGV.Columns.Add("col1", "摘要名");

                //tempDGV.Columns[0].Width = 200;
                tempDGV.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

                tempDGV.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;

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

        ///-----------------------------------------------------------------
        /// <summary>
        ///     グリッドビューへ摘要名を表示する : 2017/06/06 勘定奉行i10</summary>
        /// <param name="tempDGV">
        ///     DataGridViewオブジェクト名</param>
        ///-----------------------------------------------------------------
        private void GridViewShow_Tekiyo(DataGridView tempDGV)
        {
            //勘定奉行データベース接続文字列を取得する
            string sc = SqlControl.obcConnectSting.get(global.pblDbName);

            //勘定奉行データベースへ接続する
            SqlControl.DataControl dCon = new SqlControl.DataControl(sc);

            //摘要名データ取得
            //データリーダーを取得する
            SqlDataReader dR;
            //string sqlSTRING = "SELECT sUcd,sNm FROM wktkm01 ORDER BY sUcd";
            string sqlSTRING = "SELECT SummaryContent FROM tbAC_Summary ORDER BY SummaryCode";            

            dR = dCon.free_dsReader(sqlSTRING);

            try
            {
                int iX = 0;

                //グリッドビューに表示する
                tempDGV.RowCount = 0;

                while (dR.Read())
                {
                    tempDGV.Rows.Add();

                    tempDGV[0, iX].Value = dR["SummaryContent"].ToString().Trim();     //名称

                    iX++;
                }

                dR.Close();
                dCon.Close();

                tabData.SelectedIndex = global.TAB_TEKIYOU;
                tempDGV.CurrentCell = null;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "エラー", MessageBoxButtons.OK);
            }
        }

        /// <summary>
        /// 決算日付と中間決算日付の比較
        /// </summary>
        /// <returns></returns>
        private Boolean JudgeDate(DateTime Date1, DateTime Date2)
        {
            //Date1の日付が後の場合、NG
            if (Date1 >= Date2)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        /// <summary>
        /// 伝票ＣＳＶデータを一枚ごとに分割する
        /// </summary>
        private void LoadCsvDivide()
        {
            string imgName = string.Empty;      //画像ファイル名
            string firstFlg = global.FLGON;
            global.pblDenNum = 0;               //伝票枚数を0にセット
            string[] stArrayData;               //CSVファイルを１行単位で格納する配列

            // StreamReader の新しいインスタンスを生成する
            //入力ファイル
            StreamReader inFile = new StreamReader(global.WorkDir + global.DIR_HENKAN + global.TMPREAD, Encoding.Default);

            // 読み込んだ結果をすべて格納するための変数を宣言する
            string stResult = string.Empty;
            string stBuffer;    

            // 読み込みできる文字がなくなるまで繰り返す
            while (inFile.Peek() >= 0)
            {
                // ファイルを 1 行ずつ読み込む
                stBuffer = inFile.ReadLine();

                // カンマ区切りで分割して配列に格納する
                stArrayData = stBuffer.Split(',');

                //先頭に「*」か「#」があったら新たな伝票なのでCSVファイル作成
                if ((stArrayData[0] == "*"))
                {
                    //最初の伝票以外のとき
                    if (firstFlg == global.FLGOFF)
                    {
                        //ファイル書き出し
                        outFileWrite(stResult, imgName);
                    }

                    //伝票枚数カウント
                    global.pblDenNum++;
                    firstFlg = global.FLGOFF;

                    //画像ファイル名を取得
                    imgName = stArrayData[1];

                    //文字列バッファをクリア
                    stResult = string.Empty;
                }

                // 読み込んだものを追加で格納する
                stResult = stResult + stBuffer + Environment.NewLine;
            }

            //伝票なし
            if (global.pblDenNum == 0)
            {
                MessageBox.Show("不正な伝票ファイルです。伝票データが存在しません。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                Environment.Exit(0);
            }
            else
            {
                //ファイル書き出し
                outFileWrite(stResult, imgName);

                // 入力ファイルを閉じる
                inFile.Close();

                //入力ファイルを削除する
                //一時ファイル : "kanjo2ktmpread.dat"
                utility.FileDelete(global.WorkDir + global.DIR_HENKAN, global.TMPREAD);

                //入力ファイル : "kanjo2kocr.csv"
                utility.FileDelete(global.WorkDir + global.DIR_HENKAN, global.INFILE);

                //画像ファイル
                utility.FileDelete(global.WorkDir + global.DIR_HENKAN, "WRH*.bmp");
            }
        }

        /// <summary>
        /// 分割ファイルを書き出す
        /// </summary>
        /// <param name="tempResult">書き出す文字列</param>
        /// <param name="tempImgName">画像名(拡張子含まない)</param>
        private void outFileWrite(string tempResult, string tempImgName)
        {
            //出力ファイル
            StreamWriter outFile = new StreamWriter(global.WorkDir + global.DIR_HENKAN + global.DIVFILE, false, System.Text.Encoding.GetEncoding(932));
            outFile.Write(tempResult);

            //ファイルクローズ
            outFile.Close();

            //分割一時ファイルをコピー(henkanフォルダ → 分割)
            System.IO.File.Copy(global.WorkDir + global.DIR_HENKAN + global.DIVFILE,
                                global.WorkDir + global.DIR_INCSV +
                                 string.Format("{0:00}", DateTime.Today.Month) +
                                 string.Format("{0:00}", DateTime.Today.Day) +
                                 string.Format("{0:00}", DateTime.Now.Hour) +
                                 string.Format("{0:00}", DateTime.Now.Minute) +
                                 string.Format("{0:00}", DateTime.Now.Second) +
                                 string.Format("{0:000}", global.pblDenNum) +
                                 tempImgName.Replace(".bmp",".csv"));

            //画像ファイルをコピー(henkanフォルダ → 分割)
            System.IO.File.Copy(global.WorkDir + global.DIR_HENKAN + tempImgName,
                                global.WorkDir + global.DIR_INCSV +
                                 string.Format("{0:00}", DateTime.Today.Month) +
                                 string.Format("{0:00}", DateTime.Today.Day) +
                                 string.Format("{0:00}", DateTime.Now.Hour) +
                                 string.Format("{0:00}", DateTime.Now.Minute) +
                                 string.Format("{0:00}", DateTime.Now.Second) +
                                 string.Format("{0:000}", global.pblDenNum) +
                                 tempImgName);

            //分割一時ファイルの削除
            utility.FileDelete(global.WorkDir + global.DIR_HENKAN, global.DIVFILE);

        }

        /// <summary>
        /// 仕訳伝票読み込み
        /// </summary>
        /// <returns>仕訳伝票配列データ</returns>
        private Entity.InputRecord[] LoadDataFurikae()
        {
            global.pblDenNum = 0;
            string firstFlg = global.FLGON;

            //伝票配列データのインスタンスを生成する
            Entity.InputRecord[] sDenpyo = new Entity.InputRecord[1];

            try
            {
                //分割後のＣＳＶファイル読込
                foreach (var item in Directory.GetFiles(global.WorkDir + global.DIR_INCSV, "*.csv"))
                {
                    // StreamReader の新しいインスタンスを生成する
                    StreamReader inFile = new StreamReader(item, Encoding.Default);

                    // 読み込んだ結果をすべて格納するための変数を宣言する
                    string stResult = string.Empty;
                    string stBuffer;

                    // 読み込みできる文字がなくなるまで繰り返す
                    while (inFile.Peek() >= 0)
                    {
                        // ファイルを 1 行ずつ読み込む
                        stBuffer = inFile.ReadLine();

                        if (stBuffer != string.Empty)
                        {
                            //先頭に「*」か「#」があったら新たな伝票なのでヘッダ格納
                            if ((stBuffer.Substring(0, 1) == "*"))
                            {
                                firstFlg = global.FLGOFF;

                                //伝票枚数加算
                                global.pblDenNum++;

                                //2件目以降なら要素数を追加
                                if (global.pblDenNum != 1)
                                    sDenpyo.CopyTo(sDenpyo = new Entity.InputRecord[global.pblDenNum], 0);

                                //行データのインスタンスを生成する
                                sDenpyo[global.pblDenNum - 1].Gyou = new Entity.Gyou[global.MAXGYOU];

                                ////伝票合計メモリクリア
                                //InitDenRec_Total(global.pblDenNum - 1);

                                //伝票データ配列クリア
                                InitDenRec(global.pblDenNum - 1, sDenpyo);

                                //ヘッダ取得
                                DataGetHead(global.pblDenNum - 1, stBuffer, item.Replace(global.WorkDir + global.DIR_INCSV, string.Empty), sDenpyo);
                            }
                            else
                            {
                                //行データ格納
                                DataGetGyou(global.pblDenNum - 1, stBuffer, sDenpyo);
                            }
                        }
                    }

                    // StreamReader 閉じる
                    inFile.Close();
                }

                //伝票データなし
                if (global.pblDenNum == 0)
                {
                    MessageBox.Show("不正な伝票ファイルです。伝票データが存在しません。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    Environment.Exit(0);
                }

                if (sDenpyo[0].Head.image.Trim() == string.Empty)
                {
                    MessageBox.Show("WinReaderHand S から画像が出力されていないため" + Environment.NewLine + "画像の表示はされません", "画像表示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("伝票データ取得中にエラーが発生しました。" + Environment.NewLine + e.Message, "エラー", MessageBoxButtons.OK);
            }

            return sDenpyo;
        }

        //伝票データ初期化
        private void InitDenRec_Total(int iX)
        {
        }

        //伝票データ初期化   
        private void InitDenRec(int iX, Entity.InputRecord[] sDenpyo)
        {
            sDenpyo[iX].Head.image = string.Empty;
            sDenpyo[iX].Head.CsvFile = string.Empty;
            sDenpyo[iX].Head.Year = string.Empty;
            sDenpyo[iX].Head.Month = string.Empty;
            sDenpyo[iX].Head.Day = string.Empty;
            sDenpyo[iX].Head.Kessan = string.Empty;
            sDenpyo[iX].Head.FukusuChk = string.Empty;
            sDenpyo[iX].Head.DenNo = string.Empty;

            for (int i = 0; i < global.MAXGYOU; i++)
            {
                sDenpyo[iX].Gyou[i].GyouNum = string.Empty;
                sDenpyo[iX].Gyou[i].GyouNum = string.Empty;
                sDenpyo[iX].Gyou[i].Kari.Bumon = string.Empty;
                sDenpyo[iX].Gyou[i].Kari.Kamoku = string.Empty;
                sDenpyo[iX].Gyou[i].Kari.Hojo = string.Empty;
                sDenpyo[iX].Gyou[i].Kari.Kin = string.Empty;
                sDenpyo[iX].Gyou[i].Kari.TaxMas = string.Empty;
                sDenpyo[iX].Gyou[i].Kari.TaxKbn = string.Empty;
                sDenpyo[iX].Gyou[i].Kashi.Bumon = string.Empty;
                sDenpyo[iX].Gyou[i].Kashi.Kamoku = string.Empty;
                sDenpyo[iX].Gyou[i].Kashi.Hojo = string.Empty;
                sDenpyo[iX].Gyou[i].Kashi.Kin = string.Empty;
                sDenpyo[iX].Gyou[i].Kashi.TaxMas = string.Empty;
                sDenpyo[iX].Gyou[i].Kashi.TaxKbn = string.Empty;
                sDenpyo[iX].Gyou[i].CopyChk = string.Empty;
                sDenpyo[iX].Gyou[i].Tekiyou = string.Empty;
            }

            sDenpyo[iX].KariTotal = 0;
            sDenpyo[iX].KashiTotal = 0;
            sDenpyo[iX].Head.Kari_T = 0;
            sDenpyo[iX].Head.Kashi_T = 0;
            sDenpyo[iX].Head.FukuMai = 0;
        }

        //伝票ヘッダ部格納
        private void DataGetHead(int iX, string readbuf, string csvf, Entity.InputRecord[] sDenpyo)
        {
            // カンマ区切りで分割して配列に格納する
            string[] stArrayData = readbuf.Split(',');

            //画像ファイル名                    
            sDenpyo[iX].Head.image = csvf.Replace(".csv", ".bmp").Trim();

            //CSVファイル名
            sDenpyo[iX].Head.CsvFile = csvf;

            //年
            sDenpyo[iX].Head.Year = stArrayData[2].Replace("-", string.Empty).Trim();

            //月
            sDenpyo[iX].Head.Month = stArrayData[3].Replace("-", string.Empty).Trim();

            //日
            sDenpyo[iX].Head.Day = stArrayData[4].Replace("-", string.Empty).Trim();

            //伝票No.
            sDenpyo[iX].Head.DenNo = stArrayData[5].Replace("-", string.Empty).Trim();

            //決算処理フラグ
            sDenpyo[iX].Head.Kessan = stArrayData[6].Trim();

            //複数枚チェック
            sDenpyo[iX].Head.FukusuChk = stArrayData[7].ToString().Trim();

        }


        //伝票行データ格納
        private void DataGetGyou(int iX, string readbuf, Entity.InputRecord[] sDenpyo)
        {
            int i = 0; //行カウント

            // カンマ区切りで分割して配列に格納する
            string[] stData = readbuf.Split(',');

            //空白行は対象としない
            if (utility.NumericCheck(stData[1]) == false) return;
            if ((int.Parse(stData[1]) < 1) || (int.Parse(stData[1]) > global.MAXGYOU)) return;

            //行数から配列の添え字を取得
            i = int.Parse(stData[1]) - 1;

            ////取消記号
            sDenpyo[iX].Gyou[i].Torikeshi = stData[0].Trim();

            ////行番号
            sDenpyo[iX].Gyou[i].GyouNum = stData[1].Trim();

            //借方明細
            //部門
            sDenpyo[iX].Gyou[i].Kari.Bumon = stData[2];

            //部門データ変換処理
            if (sDenpyo[iX].Gyou[i].Kari.Bumon.Trim() == string.Empty)
            {
                sDenpyo[iX].Gyou[i].Kari.Bumon = string.Empty;
            }
            else if (sDenpyo[iX].Gyou[i].Kari.Bumon.IndexOf("-", 0) >= 0)
            {
                sDenpyo[iX].Gyou[i].Kari.Bumon = sDenpyo[iX].Gyou[i].Kari.Bumon.Replace(" ", string.Empty);
                sDenpyo[iX].Gyou[i].Kari.Bumon = sDenpyo[iX].Gyou[i].Kari.Bumon.Replace("-", string.Empty);

                if (utility.NumericCheck(sDenpyo[iX].Gyou[i].Kari.Bumon))
                {
                    sDenpyo[iX].Gyou[i].Kari.Bumon = string.Format("-{0:##0}", int.Parse(sDenpyo[iX].Gyou[i].Kari.Bumon));
                }
                else
                {
                    sDenpyo[iX].Gyou[i].Kari.Bumon = ("-" + sDenpyo[iX].Gyou[i].Kari.Bumon).Trim();
                }
            }
            else
            {
                sDenpyo[iX].Gyou[i].Kari.Bumon = sDenpyo[iX].Gyou[i].Kari.Bumon.Replace(" ", string.Empty).Trim();
                if (utility.NumericCheck(sDenpyo[iX].Gyou[i].Kari.Bumon))
                {
                    sDenpyo[iX].Gyou[i].Kari.Bumon = string.Format("{0:###0}", int.Parse(sDenpyo[iX].Gyou[i].Kari.Bumon));
                }
            }

            //勘定科目
            sDenpyo[iX].Gyou[i].Kari.Kamoku = stData[3];
            //--勘定科目データ変換処理--
            if (sDenpyo[iX].Gyou[i].Kari.Kamoku.Trim() == string.Empty)
            {
                sDenpyo[iX].Gyou[i].Kari.Kamoku = string.Empty;
            }
            else if (sDenpyo[iX].Gyou[i].Kari.Kamoku.IndexOf("-", 0) >= 0)
            {
                sDenpyo[iX].Gyou[i].Kari.Kamoku = sDenpyo[iX].Gyou[i].Kari.Kamoku.Replace(" ", string.Empty).Trim();
                sDenpyo[iX].Gyou[i].Kari.Kamoku = sDenpyo[iX].Gyou[i].Kari.Kamoku.Replace("-", string.Empty).Trim();

                if (utility.NumericCheck(sDenpyo[iX].Gyou[i].Kari.Kamoku))
                {
                    sDenpyo[iX].Gyou[i].Kari.Kamoku = string.Format("-{0:##0}", int.Parse(sDenpyo[iX].Gyou[i].Kari.Kamoku));
                }
                else
                {
                    sDenpyo[iX].Gyou[i].Kari.Kamoku = ("-" + sDenpyo[iX].Gyou[i].Kari.Kamoku).Trim();
                }
            }
            else
            {
                sDenpyo[iX].Gyou[i].Kari.Kamoku = sDenpyo[iX].Gyou[i].Kari.Kamoku.Replace(" ", string.Empty).Trim();

                if (utility.NumericCheck(sDenpyo[iX].Gyou[i].Kari.Kamoku))
                {
                    sDenpyo[iX].Gyou[i].Kari.Kamoku = string.Format("{0:###0}", int.Parse(sDenpyo[iX].Gyou[i].Kari.Kamoku));
                }
            }

            //科目が設定されていて基本情報で「部門あり」で、部門が設定されていない場合は、部門に０を設定
            if ((sDenpyo[iX].Gyou[i].Kari.Kamoku != string.Empty) && (global.pblBumonFlg == true) && (sDenpyo[iX].Gyou[i].Kari.Bumon == string.Empty))
                sDenpyo[iX].Gyou[i].Kari.Bumon = "0";

            //補助コード
            sDenpyo[iX].Gyou[i].Kari.Hojo = stData[4];

            //--補助コードデータ変換処理--
            if (sDenpyo[iX].Gyou[i].Kari.Hojo.Trim() == string.Empty)
            {
                sDenpyo[iX].Gyou[i].Kari.Hojo = string.Empty;
            }
            else if (sDenpyo[iX].Gyou[i].Kari.Hojo.IndexOf("-", 0) >= 0)
            {
                sDenpyo[iX].Gyou[i].Kari.Hojo = sDenpyo[iX].Gyou[i].Kari.Hojo.Replace(" ", string.Empty).Trim();
                sDenpyo[iX].Gyou[i].Kari.Hojo = sDenpyo[iX].Gyou[i].Kari.Hojo.Replace("-", string.Empty).Trim();

                if (utility.NumericCheck(sDenpyo[iX].Gyou[i].Kari.Hojo))
                {
                    sDenpyo[iX].Gyou[i].Kari.Hojo = string.Format("-{0:##0}", int.Parse(sDenpyo[iX].Gyou[i].Kari.Hojo));
                }
                else
                {
                    sDenpyo[iX].Gyou[i].Kari.Hojo = ("-" + sDenpyo[iX].Gyou[i].Kari.Hojo).Trim();
                }
            }
            else
            {
                sDenpyo[iX].Gyou[i].Kari.Hojo = sDenpyo[iX].Gyou[i].Kari.Hojo.Replace(" ", string.Empty).Trim();
                if (utility.NumericCheck(sDenpyo[iX].Gyou[i].Kari.Hojo))
                {
                    sDenpyo[iX].Gyou[i].Kari.Hojo = string.Format("{0:###0}", int.Parse(sDenpyo[iX].Gyou[i].Kari.Hojo));
                }
            }

            //借方金額
            sDenpyo[iX].Gyou[i].Kari.Kin = stData[5];

            //--借方金額データ変換処理--
            if (sDenpyo[iX].Gyou[i].Kari.Kin.Trim() == string.Empty)
            {
                sDenpyo[iX].Gyou[i].Kari.Kin = string.Empty;
            }
            else if (sDenpyo[iX].Gyou[i].Kari.Kin.IndexOf("-", 0) >= 0)
            {
                sDenpyo[iX].Gyou[i].Kari.Kin = sDenpyo[iX].Gyou[i].Kari.Kin.Replace(" ", string.Empty).Trim();
                sDenpyo[iX].Gyou[i].Kari.Kin = sDenpyo[iX].Gyou[i].Kari.Kin.Replace("=", string.Empty).Trim();
                if (utility.NumericCheck(sDenpyo[iX].Gyou[i].Kari.Kin))
                {
                    sDenpyo[iX].Gyou[i].Kari.Kin = string.Format("-{0:##########}", int.Parse(sDenpyo[iX].Gyou[i].Kari.Kin));
                }
                else
                {
                    sDenpyo[iX].Gyou[i].Kari.Kin = "-9";
                }
            }
            else
            {
                sDenpyo[iX].Gyou[i].Kari.Kin = sDenpyo[iX].Gyou[i].Kari.Kin.Replace(" ", string.Empty);
                if (utility.NumericCheck(sDenpyo[iX].Gyou[i].Kari.Kin))
                {
                    sDenpyo[iX].Gyou[i].Kari.Kin = string.Format("{0:##########}", int.Parse(sDenpyo[iX].Gyou[i].Kari.Kin));
                }
                else
                {
                    sDenpyo[iX].Gyou[i].Kari.Kin = "-9";
                }
            }

            //税処理
            sDenpyo[iX].Gyou[i].Kari.TaxMas = stData[6];

            //税区分
            sDenpyo[iX].Gyou[i].Kari.TaxKbn = stData[7];

            //税区分データ変換処理
            if (sDenpyo[iX].Gyou[i].Kari.TaxKbn.Trim() == string.Empty)
            {
                sDenpyo[iX].Gyou[i].Kari.TaxKbn = string.Empty;
            }
            else if (sDenpyo[iX].Gyou[i].Kari.TaxKbn.IndexOf("-", 0) >= 0)
            {
                sDenpyo[iX].Gyou[i].Kari.TaxKbn = sDenpyo[iX].Gyou[i].Kari.TaxKbn.Replace(" ", string.Empty);
                sDenpyo[iX].Gyou[i].Kari.TaxKbn = sDenpyo[iX].Gyou[i].Kari.TaxKbn.Replace("-", string.Empty);

                if (utility.NumericCheck(sDenpyo[iX].Gyou[i].Kari.TaxKbn))
                {
                    sDenpyo[iX].Gyou[i].Kari.TaxKbn = string.Format("-{0:0}", int.Parse(sDenpyo[iX].Gyou[i].Kari.TaxKbn));
                }
                else
                {
                    sDenpyo[iX].Gyou[i].Kari.TaxKbn = ("-" + sDenpyo[iX].Gyou[i].Kari.TaxKbn);
                }
            }
            else
            {
                sDenpyo[iX].Gyou[i].Kari.TaxKbn = sDenpyo[iX].Gyou[i].Kari.TaxKbn.Replace(" ", string.Empty);

                if (utility.NumericCheck(sDenpyo[iX].Gyou[i].Kari.TaxKbn))
                {
                    sDenpyo[iX].Gyou[i].Kari.TaxKbn = string.Format("{0:#0}", int.Parse(sDenpyo[iX].Gyou[i].Kari.TaxKbn));
                }

            }

            //貸方明細 ------------------------------------------------------------------------------------------
            //部門
            sDenpyo[iX].Gyou[i].Kashi.Bumon = stData[8];

            //部門データ変換処理
            if (sDenpyo[iX].Gyou[i].Kashi.Bumon.Trim() == string.Empty)
            {
                sDenpyo[iX].Gyou[i].Kashi.Bumon = string.Empty;
            }
            else if (sDenpyo[iX].Gyou[i].Kashi.Bumon.IndexOf("-", 0) >= 0)
            {
                sDenpyo[iX].Gyou[i].Kashi.Bumon = sDenpyo[iX].Gyou[i].Kashi.Bumon.Replace(" ", string.Empty);
                sDenpyo[iX].Gyou[i].Kashi.Bumon = sDenpyo[iX].Gyou[i].Kashi.Bumon.Replace("-", string.Empty);

                if (utility.NumericCheck(sDenpyo[iX].Gyou[i].Kashi.Bumon))
                {
                    sDenpyo[iX].Gyou[i].Kashi.Bumon = string.Format("-{0:##0}", int.Parse(sDenpyo[iX].Gyou[i].Kashi.Bumon));
                }
                else
                {
                    sDenpyo[iX].Gyou[i].Kashi.Bumon = ("-" + sDenpyo[iX].Gyou[i].Kashi.Bumon).Trim();
                }
            }
            else
            {
                sDenpyo[iX].Gyou[i].Kashi.Bumon = sDenpyo[iX].Gyou[i].Kashi.Bumon.Replace(" ", string.Empty).Trim();
                if (utility.NumericCheck(sDenpyo[iX].Gyou[i].Kashi.Bumon))
                {
                    sDenpyo[iX].Gyou[i].Kashi.Bumon = string.Format("{0:###0}", int.Parse(sDenpyo[iX].Gyou[i].Kashi.Bumon));
                }
            }

            //勘定科目
            sDenpyo[iX].Gyou[i].Kashi.Kamoku = stData[9];
            //--勘定科目データ変換処理--
            if (sDenpyo[iX].Gyou[i].Kashi.Kamoku.Trim() == string.Empty)
            {
                sDenpyo[iX].Gyou[i].Kashi.Kamoku = string.Empty;
            }
            else if (sDenpyo[iX].Gyou[i].Kashi.Kamoku.IndexOf("-", 0) >= 0)
            {
                sDenpyo[iX].Gyou[i].Kashi.Kamoku = sDenpyo[iX].Gyou[i].Kashi.Kamoku.Replace(" ", string.Empty).Trim();
                sDenpyo[iX].Gyou[i].Kashi.Kamoku = sDenpyo[iX].Gyou[i].Kashi.Kamoku.Replace("-", string.Empty).Trim();

                if (utility.NumericCheck(sDenpyo[iX].Gyou[i].Kashi.Kamoku))
                {
                    sDenpyo[iX].Gyou[i].Kashi.Kamoku = string.Format("-{0:##0}", int.Parse(sDenpyo[iX].Gyou[i].Kashi.Kamoku));
                }
                else
                {
                    sDenpyo[iX].Gyou[i].Kashi.Kamoku = ("-" + sDenpyo[iX].Gyou[i].Kashi.Kamoku).Trim();
                }
            }
            else
            {
                sDenpyo[iX].Gyou[i].Kashi.Kamoku = sDenpyo[iX].Gyou[i].Kashi.Kamoku.Replace(" ", string.Empty).Trim();

                if (utility.NumericCheck(sDenpyo[iX].Gyou[i].Kashi.Kamoku))
                {
                    sDenpyo[iX].Gyou[i].Kashi.Kamoku = string.Format("{0:###0}", int.Parse(sDenpyo[iX].Gyou[i].Kashi.Kamoku));
                }
            }

            //科目が設定されていて基本情報で「部門あり」で、部門が設定されていない場合は、部門に０を設定
            if ((sDenpyo[iX].Gyou[i].Kashi.Kamoku != string.Empty) && (global.pblBumonFlg == true) && (sDenpyo[iX].Gyou[i].Kashi.Bumon == string.Empty))
                sDenpyo[iX].Gyou[i].Kashi.Bumon = "0";

            //補助コード
            sDenpyo[iX].Gyou[i].Kashi.Hojo = stData[10];

            //--補助コードデータ変換処理--
            if (sDenpyo[iX].Gyou[i].Kashi.Hojo.Trim() == string.Empty)
            {
                sDenpyo[iX].Gyou[i].Kashi.Hojo = string.Empty;
            }
            else if (sDenpyo[iX].Gyou[i].Kashi.Hojo.IndexOf("-", 0) >= 0)
            {
                sDenpyo[iX].Gyou[i].Kashi.Hojo = sDenpyo[iX].Gyou[i].Kashi.Hojo.Replace(" ", string.Empty).Trim();
                sDenpyo[iX].Gyou[i].Kashi.Hojo = sDenpyo[iX].Gyou[i].Kashi.Hojo.Replace("-", string.Empty).Trim();

                if (utility.NumericCheck(sDenpyo[iX].Gyou[i].Kashi.Hojo))
                {
                    sDenpyo[iX].Gyou[i].Kashi.Hojo = string.Format("-{0:##0}", int.Parse(sDenpyo[iX].Gyou[i].Kashi.Hojo));
                }
                else
                {
                    sDenpyo[iX].Gyou[i].Kashi.Hojo = ("-" + sDenpyo[iX].Gyou[i].Kashi.Hojo).Trim();
                }

            }
            else
            {
                sDenpyo[iX].Gyou[i].Kashi.Hojo = sDenpyo[iX].Gyou[i].Kashi.Hojo.Replace(" ", string.Empty).Trim();
                if (utility.NumericCheck(sDenpyo[iX].Gyou[i].Kashi.Hojo))
                {
                    sDenpyo[iX].Gyou[i].Kashi.Hojo = string.Format("{0:###0}", int.Parse(sDenpyo[iX].Gyou[i].Kashi.Hojo));
                }
            }

            //貸方金額
            sDenpyo[iX].Gyou[i].Kashi.Kin = stData[11];

            //--貸方金額データ変換処理--
            if (sDenpyo[iX].Gyou[i].Kashi.Kin.Trim() == string.Empty)
            {
                sDenpyo[iX].Gyou[i].Kashi.Kin = string.Empty;
            }
            else if (sDenpyo[iX].Gyou[i].Kashi.Kin.IndexOf("-", 0) >= 0)
            {
                sDenpyo[iX].Gyou[i].Kashi.Kin = sDenpyo[iX].Gyou[i].Kashi.Kin.Replace(" ", string.Empty).Trim();
                sDenpyo[iX].Gyou[i].Kashi.Kin = sDenpyo[iX].Gyou[i].Kashi.Kin.Replace("=", string.Empty).Trim();
                if (utility.NumericCheck(sDenpyo[iX].Gyou[i].Kashi.Kin))
                {
                    sDenpyo[iX].Gyou[i].Kashi.Kin = string.Format("-{0:##########}", int.Parse(sDenpyo[iX].Gyou[i].Kashi.Kin));
                }
                else
                {
                    sDenpyo[iX].Gyou[i].Kashi.Kin = "-9";
                }
            }
            else
            {
                sDenpyo[iX].Gyou[i].Kashi.Kin = sDenpyo[iX].Gyou[i].Kashi.Kin.Replace(" ", string.Empty);
                if (utility.NumericCheck(sDenpyo[iX].Gyou[i].Kashi.Kin))
                {
                    sDenpyo[iX].Gyou[i].Kashi.Kin = string.Format("{0:##########}", int.Parse(sDenpyo[iX].Gyou[i].Kashi.Kin));
                }
                else
                {
                    sDenpyo[iX].Gyou[i].Kashi.Kin = "-9";
                }
            }

            //税処理
            sDenpyo[iX].Gyou[i].Kashi.TaxMas = stData[12];

            //税区分
            sDenpyo[iX].Gyou[i].Kashi.TaxKbn = stData[13];

            //税区分データ変換処理
            if (sDenpyo[iX].Gyou[i].Kashi.TaxKbn.Trim() == string.Empty)
            {
                sDenpyo[iX].Gyou[i].Kashi.TaxKbn = string.Empty;
            }
            else if (sDenpyo[iX].Gyou[i].Kashi.TaxKbn.IndexOf("-", 0) >= 0)
            {
                sDenpyo[iX].Gyou[i].Kashi.TaxKbn = sDenpyo[iX].Gyou[i].Kashi.TaxKbn.Replace(" ", string.Empty);
                sDenpyo[iX].Gyou[i].Kashi.TaxKbn = sDenpyo[iX].Gyou[i].Kashi.TaxKbn.Replace("-", string.Empty);

                if (utility.NumericCheck(sDenpyo[iX].Gyou[i].Kashi.TaxKbn))
                {
                    sDenpyo[iX].Gyou[i].Kashi.TaxKbn = string.Format("-{0:0}", int.Parse(sDenpyo[iX].Gyou[i].Kashi.TaxKbn));
                }
                else
                {
                    sDenpyo[iX].Gyou[i].Kashi.TaxKbn = ("-" + sDenpyo[iX].Gyou[i].Kashi.TaxKbn);
                }
            }
            else
            {
                sDenpyo[iX].Gyou[i].Kashi.TaxKbn = sDenpyo[iX].Gyou[i].Kashi.TaxKbn.Replace(" ", string.Empty);

                if (utility.NumericCheck(sDenpyo[iX].Gyou[i].Kashi.TaxKbn))
                {
                    sDenpyo[iX].Gyou[i].Kashi.TaxKbn = string.Format("{0:#0}", int.Parse(sDenpyo[iX].Gyou[i].Kashi.TaxKbn));
                }
            }

            //摘要複写
            if (i == 0)
            {
                sDenpyo[iX].Gyou[i].CopyChk = "0";
            }
            else
            {
                sDenpyo[iX].Gyou[i].CopyChk = stData[14];
            }

            //摘要
            sDenpyo[iX].Gyou[i].Tekiyou = stData[15].TrimEnd();

        }

        ///---------------------------------------------------------------
        /// <summary>
        ///     エラーチェックルーチン </summary>
        /// <param name="dData">
        ///     伝票配列データ</param>
        /// <param name="x1">
        ///     エラーなしで処理を終了するとき：true、
        ///     エラー有りまたはエラーなしで処理を終了しないとき：false</param>
        /// <returns>
        ///     エラー配列データ</returns>
        ///---------------------------------------------------------------
        private errCheck.Errtbl[] ChkMainNew(Entity.InputRecord[] dData, out Boolean x1)
        {
            //オーナーフォームを無効にする
            this.Enabled = false;

            //プログレスバーを表示する
            frmPrg frmP = new frmPrg();
            frmP.Owner = this;
            frmP.Show();

            //エラーチェックインスタンス作成
            errCheck ec = new errCheck();

            //伝票データを順次読み込みエラーチェックを実施する
            for (int i = 0; i < global.pblDenNum; i++)
            {
                //結合チェック
                frmP.Text = "ＮＧチェック中・・・結合";
                frmP.progressValue = 10;
                frmP.ProgressStep();

                ec.ChkCombineNEW(i, dData);          //結合枚数
                ec.ChkCombineItem(i, dData);         //結合行数
                ec.ChkCombineDateNEW(i, dData);      //結合日付
                //ec.ChkCombineDenNoNEW(i, dData);     //結合伝票
                ec.ChkCombineKessanNEW(i, dData);    //決算

                // 日付チェック
                frmP.Text = "ＮＧチェック中・・・日付";
                frmP.progressValue = 20;
                frmP.ProgressStep();
                
                if (ec.ChkDateNEW(i, dData) == true)
                {
                    // 決算日付チェック
                    frmP.Text = "ＮＧチェック中・・・決算日付";
                    frmP.progressValue = 34;
                    frmP.ProgressStep();

                    if (ec.ChkDateKessanNEW(i, dData) == true)
                    {
                        // 会計期間チェック
                        frmP.Text = "ＮＧチェック中・・・会計期間";
                        frmP.progressValue = 38;
                        frmP.ProgressStep();
                        ec.ChkDateKikanNEW(i, dData);

                        //if (ec.ChkDateKikanNEW(i, dData) == true)
                        //{
                        //    // 日付入力範囲チェック
                        //    frmP.Text = "ＮＧチェック中・・・日付入力範囲";
                        //    frmP.progressValue = 42;
                        //    frmP.ProgressStep();
                        //    ec.ChkDateLimitNEW(i, dData);
                        //}
                    }
                }

                // 入力不備チェック
                frmP.Text = "ＮＧチェック中・・・入力不備";
                frmP.progressValue = 46;
                frmP.ProgressStep();
                ec.ChkDataPoorNEW(i, dData);

                // 勘定科目コードチェック
                frmP.Text = "ＮＧチェック中・・・勘定科目コード";
                frmP.progressValue = 50;
                frmP.ProgressStep();
                ec.ChkKamokuNEW(i, dData);

                // 補助コードチェック
                frmP.Text = "ＮＧチェック中・・・補助科目コード";
                frmP.progressValue = 60;
                frmP.ProgressStep();
                ec.ChkHojoNEW(i, dData);

                // 部門コードチェック
                frmP.Text = "ＮＧチェック中・・・部門コード";
                frmP.progressValue = 70;
                frmP.ProgressStep();
                ec.ChkBumonNEW(i, dData);

                // 消費税計算区分（略名：税処理）コードチェック
                frmP.Text = "ＮＧチェック中・・・税処理コード";
                frmP.progressValue = 80;
                frmP.ProgressStep();
                ec.ChkOtherNEW(i, dData);

                // 税区分コードチェック
                frmP.Text = "ＮＧチェック中・・・税区分コード";
                frmP.progressValue = 85;
                frmP.ProgressStep();
                ec.ChkTaxKbnNEW(i, dData);

                // 貸借差額チェック
                frmP.Text = "ＮＧチェック中・・・貸借差額";
                frmP.progressValue = 90;
                frmP.ProgressStep();
                ec.ChkSumNEW(i, dData);

                //相手科目未記入チェック
                frmP.Text = "ＮＧチェック中・・・相手科目";
                frmP.progressValue = 95;
                frmP.ProgressStep();
                ec.ChkAiteNEW(i, dData);

                // 摘要複写
                frmP.Text = "ＮＧチェック中・・・摘要文字数";
                frmP.progressValue = 96;
                frmP.ProgressStep();
                ec.ChkTekiyou(i, dData);

                // 有効明細
                frmP.Text = "ＮＧチェック中・・・有効明細";
                frmP.progressValue = 98;
                frmP.ProgressStep();
                ec.ChkYukoMeisai(i, dData);

                // 有効明細：日付摘要のみの場合は、エラーとする
                frmP.Text = "ＮＧチェック中・・・摘要のみ";
                frmP.progressValue = 99;
                frmP.ProgressStep();
                ec.ChkTekiyouOnly(i, dData);
            }

            // いったんオーナーをアクティブにする
            this.Activate();

            // 進行状況ダイアログを閉じる
            frmP.Close();

            // オーナーのフォームを有効に戻す
            this.Enabled = true;

            x1 = false;

            //エラー有り
            if (ec.eTbl[0].Count > 0)
            {
                DenIndex = ec.eTbl[0].DenNo;                //現在の伝票添え字
                DataShow(DenIndex, DenData, ec.eTbl);       //データ表示
            }
            //エラーなし
            else
            {
                DenIndex = 0;                               //インデックスは最初の伝票とする
                DataShow(DenIndex, DenData, ec.eTbl);       //データ表示

                if (MessageBox.Show("ＮＧは見つかりませんでした。処理を終了しますか？", Application.ProductName, MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == DialogResult.Yes)
                {
                    //終了処理
                    x1 = true;
                }
            }

            return ec.eTbl;
        }

        ///-------------------------------------------------------------
        /// <summary>
        ///     スーパーコレクトデータ画面表示 </summary>
        /// <param name="iX">
        ///     現在の伝票添え字</param>
        /// <param name="st">
        ///     伝票配列データ</param>
        /// <param name="err">
        ///     エラー配列データ</param>
        ///-------------------------------------------------------------
        private void DataShow(int iX, Entity.InputRecord[] st, errCheck.Errtbl[] err)
        {
            gcMultiRow1.EditMode = EditMode.EditOnShortcutKey;

            //伝票ページ表示
            this.lblNowDen.Text = "（ " + (iX + 1).ToString() + "／" + global.pblDenNum.ToString() + " ）";
            
            //伝票ヘッダ表示
            this.txtYear.Text =st[iX].Head.Year;
            this.txtMonth.Text = st[iX].Head.Month;
            this.txtDay.Text = st[iX].Head.Day;
            this.txtDenNo.Text = st[iX].Head.DenNo;

            //西暦のときは二桁表示
            this.lblGengo.Text = company.Reki;
    
            if (company.Hosei == "0") 
            {
                if (utility.NumericCheck(st[iX].Head.Year))
                {
                    this.txtYear.Text = string.Format("{0:00}", int.Parse(st[iX].Head.Year));
                }
                else
                {
                    this.txtYear.Text = st[iX].Head.Year;
                }
            }
            //和暦のときは一桁表示
            else
            {
                this.txtYear.Text = st[iX].Head.Year;
            }
    
            //決算処理フラグ
            if (st[iX].Head.Kessan == "1")
            {
                this.ChkKessan.Checked = true;
            }
            else
            {
                this.ChkKessan.Checked = false;
            }
    
            //複数枚チェック
            if (st[iX].Head.FukusuChk == "1")
            {
                this.chkFukusuChk.Checked = true;
            }
            else
            {
                this.chkFukusuChk.Checked = false;
            }

            //伝票行表示
            this.gcMultiRow1.AllowUserToAddRows = false;                    //手動による行追加を禁止する
            this.gcMultiRow1.AllowUserToDeleteRows = false;                 //手動による行削除を禁止する
            this.gcMultiRow1.Rows.Clear();                                  //行数をクリア
            this.gcMultiRow1.RowCount = global.MAXGYOU;                     //行数を設定
            this.gcMultiRow1.RowsDefaultCellStyle.ForeColor = Color.Blue;   //テキストカラーの設定

            //取消欄をクリア
            chkDel0.Checked = false;
            chkDel1.Checked = false;
            chkDel2.Checked = false;
            chkDel3.Checked = false;
            chkDel4.Checked = false;
            chkDel5.Checked = false;
            chkDel6.Checked = false;

            for (int i = 0; i < global.MAXGYOU; i++)
            {
                //　借　方　項　目　の　表　示　/////////////////////////////////////////////////////////////////
                gcMultiRow1[i, MultiRow.DP_KARI_CODEB].Value = st[iX].Gyou[i].Kari.Bumon;       //部門コード
                gcMultiRow1[i, MultiRow.DP_KARI_CODE].Value = st[iX].Gyou[i].Kari.Kamoku;       //科目コード
                gcMultiRow1[i, MultiRow.DP_KARI_CODEH].Value = st[iX].Gyou[i].Kari.Hojo;        //補助コード
                gcMultiRow1[i, MultiRow.DP_KARI_KIN].Value = st[iX].Gyou[i].Kari.Kin;           //金額
                gcMultiRow1[i, MultiRow.DP_KARI_ZEI_S].Value = st[iX].Gyou[i].Kari.TaxMas;      //税処理
                gcMultiRow1[i, MultiRow.DP_KARI_ZEI].Value = st[iX].Gyou[i].Kari.TaxKbn;        //税区分

                //　貸　方　項　目　の　表　示　/////////////////////////////////////////////////////////////////
                gcMultiRow1[i, MultiRow.DP_KASHI_CODEB].Value = st[iX].Gyou[i].Kashi.Bumon;     //部門コード
                gcMultiRow1[i, MultiRow.DP_KASHI_CODE].Value = st[iX].Gyou[i].Kashi.Kamoku;     //科目コード
                gcMultiRow1[i, MultiRow.DP_KASHI_CODEH].Value = st[iX].Gyou[i].Kashi.Hojo;      //補助コード
                gcMultiRow1[i, MultiRow.DP_KASHI_KIN].Value = st[iX].Gyou[i].Kashi.Kin;         //金額
                gcMultiRow1[i, MultiRow.DP_KASHI_ZEI_S].Value = st[iX].Gyou[i].Kashi.TaxMas;    //税処理
                gcMultiRow1[i, MultiRow.DP_KASHI_ZEI].Value = st[iX].Gyou[i].Kashi.TaxKbn;      //税区分

                //摘要複写チェックボックス
                if (st[iX].Gyou[i].CopyChk == "0")
                {
                    if (i == 0) chkFu0.Checked = false;
                    if (i == 1) chkFu1.Checked = false;
                    if (i == 2) chkFu2.Checked = false;
                    if (i == 3) chkFu3.Checked = false;
                    if (i == 4) chkFu4.Checked = false;
                    if (i == 5) chkFu5.Checked = false;
                    if (i == 6) chkFu6.Checked = false;
                }
                else
                {
                    if (i == 0) chkFu0.Checked = true;
                    if (i == 1) chkFu1.Checked = true;
                    if (i == 2) chkFu2.Checked = true;
                    if (i == 3) chkFu3.Checked = true;
                    if (i == 4) chkFu4.Checked = true;
                    if (i == 5) chkFu5.Checked = true;
                    if (i == 6) chkFu6.Checked = true;
                }

                //摘要
                gcMultiRow1[i, MultiRow.DP_TEKIYOU].Value = st[iX].Gyou[i].Tekiyou;
                gcMultiRow1[i, MultiRow.DP_TEKIYOU].Style.ForeColor = Color.Blue;

                //取消区分
                if (st[iX].Gyou[i].Torikeshi == "0")
                {
                    if (i == 0) chkDel0.Checked = false;
                    if (i == 1) chkDel1.Checked = false;
                    if (i == 2) chkDel2.Checked = false;
                    if (i == 3) chkDel3.Checked = false;
                    if (i == 4) chkDel4.Checked = false;
                    if (i == 5) chkDel5.Checked = false;
                    if (i == 6) chkDel6.Checked = false;
                }
                else
                {
                    if (i == 0) chkDel0.Checked = true;
                    if (i == 1) chkDel1.Checked = true;
                    if (i == 2) chkDel2.Checked = true;
                    if (i == 3) chkDel3.Checked = true;
                    if (i == 4) chkDel4.Checked = true;
                    if (i == 5) chkDel5.Checked = true;
                    if (i == 6) chkDel6.Checked = true;
                }
            }

            //摘要複写
            for (int i = 0; i < gcMultiRow1.RowCount; i++)
            {
                ShowTekiyou1gyo(i, 0, st);
            }
                
            //頁合計
            gcMultiRow1.ColumnFooters[0].Cells[MultiRow.DP_KARI_P].Value = st[iX].KariTotal;        //借方合計
            gcMultiRow1.ColumnFooters[0].Cells[MultiRow.DP_KASHI_P].Value = st[iX].KashiTotal;      //貸方合計
            gcMultiRow1.ColumnFooters[0].Cells[MultiRow.DP_SAGAKU_P].Value = st[iX].KariTotal - st[iX].KashiTotal;  //差額合計
                
            //差額があれば赤表示
            if (gcMultiRow1.ColumnFooters[0].Cells[MultiRow.DP_SAGAKU_P].Value.ToString() != "0")
            {
                gcMultiRow1.ColumnFooters[0].Cells[MultiRow.DP_SAGAKU_P].Style.ForeColor = Color.Red;
            }
            else
            {
                gcMultiRow1.ColumnFooters[0].Cells[MultiRow.DP_SAGAKU_P].Style.ForeColor = Color.Black;
            }

            //伝票合計
            gcMultiRow1.ColumnFooters[0].Cells[MultiRow.DP_KARI_T].Value = st[iX].Head.Kari_T;      //借方合計
            gcMultiRow1.ColumnFooters[0].Cells[MultiRow.DP_KASHI_T].Value = st[iX].Head.Kashi_T;    //貸方合計
            gcMultiRow1.ColumnFooters[0].Cells[MultiRow.DP_SAGAKU_T].Value = st[iX].Head.Kari_T - st[iX].Head.Kashi_T;  //差額合計

            //差額があれば赤表示
            if (gcMultiRow1.ColumnFooters[0].Cells[MultiRow.DP_SAGAKU_T].Value.ToString() != "0")
            {
                gcMultiRow1.ColumnFooters[0].Cells[MultiRow.DP_SAGAKU_T].Style.ForeColor = Color.Red;
            }
            else
            {
                gcMultiRow1.ColumnFooters[0].Cells[MultiRow.DP_SAGAKU_T].Style.ForeColor = Color.Black;
            }
                
            //スクロールバー設定
            hScrollBar1.Minimum = 0;
            hScrollBar1.Maximum = global.pblDenNum - 1;
            hScrollBar1.Value = iX;
            double d = global.pblDenNum / 10;
            hScrollBar1.LargeChange = (int)System.Math.Floor(d) + 1;
                
            //伝票めくりボタン可能・不可設定
            btnFirst.Enabled = true;
            btnBefore.Enabled = true;
            btnNext.Enabled = true;
            btnEnd.Enabled = true;
        
            //先頭の伝票のとき
            if (iX == 0)
            {
                btnFirst.Enabled = false;
                btnBefore.Enabled = false;
            }
        
            //最終の伝票のとき
            if ((iX + 1) == global.pblDenNum)
            {
                btnNext.Enabled = false;
                btnEnd.Enabled = false;
            }
                    
            //伝票ヘッダの色初期化
            txtYear.BackColor = Color.White;
            txtYear.ForeColor = Color.Blue;
            txtMonth.BackColor = Color.White;
            txtMonth.ForeColor = Color.Blue;
            txtDay.BackColor = Color.White;
            txtDay.ForeColor = Color.Blue;
            txtDenNo.BackColor = Color.White;
            txtDenNo.ForeColor = Color.Blue;

            //エラー情報表示
            ShowNG_Grid(fgNg, err);

            //エラー箇所バックカラー
            global.pblKessanColor = System.Drawing.SystemColors.Control;
            ChkKessan.BackColor = System.Drawing.SystemColors.Control;
        
            global.pblFukuColor = System.Drawing.SystemColors.Control;
            chkFukusuChk.BackColor = System.Drawing.SystemColors.Control;
        
            global.pblSagakuColor = global.pblBackColor;
            gcMultiRow1.ColumnFooters[0].Cells[MultiRow.DP_SAGAKU_T].Style.BackColor = global.pblBackColor;

            Show_NGColor(iX, err);

            //画像表示
            ShowImage(iX, st[iX].Head.image);

            //multirow選択解除
            gcMultiRow1.ClearSelection();

            //カーソルを戻す
            btnOk.Focus();
            gcMultiRow1.EditMode = EditMode.EditProgrammatically;

        }

        /// <summary>
        /// 摘要複写
        /// </summary>
        /// <param name="Nowcnt">行index</param>
        /// <param name="Mode">0:初期表示時、1:複数チェックON・OFFの時</param>
        /// <param name="st">伝票配列データ</param>
        private void ShowTekiyou1gyo(int Nowcnt, int Mode, Entity.InputRecord[] st)
        {
            int iSpacePos;
            int sLen;
            string sTekiyo;
            string sTekiyoW;

            //全行表示未完了の場合は終了
            if (gcMultiRow1[6, MultiRow.DP_TEKIYOU].Value == null) return;

            //各行の処理
            if (chkFu0.Checked == false) gcMultiRow1[0, MultiRow.DP_TEKIYOU].Value += "　";
            if (chkFu1.Checked == false) gcMultiRow1[1, MultiRow.DP_TEKIYOU].Value += "　";
            if (chkFu2.Checked == false) gcMultiRow1[2, MultiRow.DP_TEKIYOU].Value += "　";
            if (chkFu3.Checked == false) gcMultiRow1[3, MultiRow.DP_TEKIYOU].Value += "　";
            if (chkFu4.Checked == false) gcMultiRow1[4, MultiRow.DP_TEKIYOU].Value += "　";
            if (chkFu5.Checked == false) gcMultiRow1[5, MultiRow.DP_TEKIYOU].Value += "　";
            if (chkFu6.Checked == false) gcMultiRow1[6, MultiRow.DP_TEKIYOU].Value += "　";

            //先頭行は摘要複写機能はない
            if (Nowcnt != 0)
            {
                if (((Nowcnt == 0 && chkDel0.Checked == false) || (Nowcnt == 1 && chkDel1.Checked == false) || 
                     (Nowcnt == 2 && chkDel0.Checked == false) || (Nowcnt == 3 && chkDel1.Checked == false) || 
                     (Nowcnt == 4 && chkDel0.Checked == false) || (Nowcnt == 5 && chkDel1.Checked == false) || 
                     (Nowcnt == 6 && chkDel0.Checked == false)) && 
                    gcMultiRow1[Nowcnt - 1, MultiRow.DP_TEKIYOU].Value.ToString() != string.Empty)
                {
                    //前行が取り消し行でなく、かつ摘要記述がある場合、かつ現行の摘要入力がある場合のみ摘要複写が有効
                    if ((Nowcnt == 0 && chkFu0.Checked == true) ||(Nowcnt == 1 && chkFu1.Checked == true) || 
                        (Nowcnt == 2 && chkFu2.Checked == true) ||(Nowcnt == 3 && chkFu3.Checked == true) || 
                        (Nowcnt == 4 && chkFu4.Checked == true) ||(Nowcnt == 5 && chkFu5.Checked == true) || 
                        (Nowcnt == 6 && chkFu6.Checked == true))
                    {
                        //前行が取消行でない場合のみ適用複写が有効
                        //右のスペースは削除する
                        //摘要複写の対象は、１文字目から次のスペースまでとする(全角チェック)
                        iSpacePos = gcMultiRow1[Nowcnt - 1, MultiRow.DP_TEKIYOU].Value.ToString().IndexOf(global.TEKIYO_SPACE_ZEN, 0);
                        if (iSpacePos == -1)
                        {
                            //半角スペースチェック
                            iSpacePos = gcMultiRow1[Nowcnt - 1, MultiRow.DP_TEKIYOU].Value.ToString().IndexOf(global.TEKIYO_SPACE_HAN, 0);
                        }

                        if (iSpacePos == -1)
                        {
                            //スペースが見つからない場合は、摘要すべてが複写対象
                            iSpacePos = gcMultiRow1[Nowcnt - 1, MultiRow.DP_TEKIYOU].Value.ToString().Length;
                        }

                        if (iSpacePos > 0)
                        {
                            sTekiyo = gcMultiRow1[Nowcnt - 1, MultiRow.DP_TEKIYOU].Value.ToString().Substring(0, iSpacePos);
                        }
                        else
                        {
                            sTekiyo = string.Empty;
                        }

                        if (sTekiyo.Trim() != string.Empty)
                        {
                            sTekiyoW = gcMultiRow1[Nowcnt, MultiRow.DP_TEKIYOU].Value.ToString();

                            if (sTekiyoW.Length < sTekiyo.Length)
                            {
                                sTekiyoW = sTekiyo;
                            }
                            else
                            {
                                sLen = sTekiyo.Length;
                                sTekiyoW = sTekiyo + sTekiyoW.Remove(0, sLen);
                            }

                            gcMultiRow1[Nowcnt, MultiRow.DP_TEKIYOU].Value = sTekiyoW;
                        }
                    }
                    else
                    {
                        if (Mode == 1)        //複写をON/OFFしたときのみ有効
                        {
                            //摘要複写のチェックを解除した場合
                            iSpacePos = gcMultiRow1[Nowcnt - 1, MultiRow.DP_TEKIYOU].Value.ToString().IndexOf(global.TEKIYO_SPACE_ZEN, 0);
                            if (iSpacePos == -1)
                            {
                                //半角スペースチェック
                                iSpacePos = gcMultiRow1[Nowcnt - 1, MultiRow.DP_TEKIYOU].Value.ToString().IndexOf(global.TEKIYO_SPACE_HAN, 0);
                            }

                            if (iSpacePos == -1)
                            {
                                //スペースが見つからない場合は、摘要すべてが複写対象
                                iSpacePos = gcMultiRow1[Nowcnt - 1, MultiRow.DP_TEKIYOU].Value.ToString().Length;
                            }

                            if (iSpacePos > 0)
                            {
                                sTekiyo = gcMultiRow1[Nowcnt - 1, MultiRow.DP_TEKIYOU].Value.ToString().Substring(0, iSpacePos);
                            }
                            else
                            {
                                sTekiyo = string.Empty;
                            }

                            //スペース埋めをする
                            if (sTekiyo.Trim() != string.Empty)
                            {
                                sLen = sTekiyo.Length;
                                sTekiyo = string.Empty;
                                sTekiyo = sTekiyo.PadLeft(sLen, '　');
                                sTekiyoW = gcMultiRow1[Nowcnt, MultiRow.DP_TEKIYOU].Value.ToString();

                                if (sTekiyoW.Length < sTekiyo.Length)
                                {
                                    sTekiyoW = sTekiyo;
                                }
                                else
                                {
                                    sLen = sTekiyo.Length;
                                    sTekiyoW = sTekiyo + sTekiyoW.Remove(0, sLen);
                                }

                                gcMultiRow1[Nowcnt, MultiRow.DP_TEKIYOU].Value = sTekiyoW;
                            }
                        }
                    }
                }
            }

            //摘要にスペースが７個設定されてしまう不具合の対策
            for (int Cnt = 0; Cnt < gcMultiRow1.RowCount; Cnt++)
            {
                //右のスペースを削除する
                if (gcMultiRow1[Cnt, MultiRow.DP_TEKIYOU].Value != null)
                {
                    gcMultiRow1[Cnt, MultiRow.DP_TEKIYOU].Value = gcMultiRow1[Cnt, MultiRow.DP_TEKIYOU].Value.ToString().TrimEnd();
                }
            }
        }

        /// <summary>
        /// エラーリストグリッド表示
        /// </summary>
        /// <param name="tempDGV">エラー表示グリッド</param>
        /// <param name="err">エラーテーブル</param>
        private void ShowNG_Grid(DataGridView tempDGV, errCheck.Errtbl[] err)
        {
            // 列スタイルを変更する
            tempDGV.EnableHeadersVisualStyles = false;

            // 列ヘッダー表示位置指定
            tempDGV.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;

            // 列ヘッダーフォント指定
            tempDGV.ColumnHeadersDefaultCellStyle.Font = new Font("Meiryo UI", 9, FontStyle.Regular);

            // データフォント指定
            tempDGV.DefaultCellStyle.Font = new Font("Meiryo UI", (float)9, FontStyle.Regular);

            // 行の高さ
            tempDGV.ColumnHeadersHeight = 18;
            tempDGV.RowTemplate.Height = 18;

            // 全体の高さ
            //tempDGV.Height = 253;

            // 奇数行の色
            //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

            // 各列設定
            tempDGV.Rows.Clear();
            tempDGV.Columns.Clear();
            tempDGV.Columns.Add("col1", "頁");
            tempDGV.Columns.Add("col2", "行");
            tempDGV.Columns.Add("col3", "貸借");
            tempDGV.Columns.Add("col4", "データ");
            tempDGV.Columns.Add("col5", "エラー内容");
            tempDGV.Columns.Add("col6", "エラー箇所");

            tempDGV.Columns[5].Visible = false; //エラー箇所は非表示

            tempDGV.Columns[0].Width = 35;
            tempDGV.Columns[1].Width = 35;
            tempDGV.Columns[2].Width = 60;
            tempDGV.Columns[3].Width = 200;
            tempDGV.Columns[4].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

            tempDGV.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            tempDGV.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            tempDGV.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            tempDGV.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            tempDGV.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;

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

            //エラーがなければ終了
            if (err[0].Count == 0)
            {
                lblErr.Text = "エラー件数 ： 0件";
                return;
            }

            //エラー内容をグリッドに表示
            int iX = 0;
            for (int pCnt = 0; pCnt < global.pblDenNum; pCnt++)
            {
                for (int LineCnt = 0; LineCnt < global.MAXGYOU; LineCnt++)
                {
                    for (int j = 0; j < err.Length; j++)
                    {
                        if (err[j].DenNo == pCnt && err[j].LINE == LineCnt)
                        {
                            tempDGV.Rows.Add();
                            tempDGV[0, iX].Value = err[j].DenNo + 1;

                            if (err[j].Field != "借" && err[j].Field != "貸")
                            {
                                tempDGV[1, iX].Value = err[j].LINE;
                            }
                            else
                            {
                                tempDGV[1, iX].Value = err[j].LINE + 1;
                            }

                            tempDGV[2, iX].Value = err[j].Field;
                            tempDGV[3, iX].Value = err[j].Data;
                            tempDGV[4, iX].Value = err[j].Notes;
                            tempDGV[5, iX].Value = err[j].DpPos;

                            iX++;
                        }
                    }
                }
            }

            //表示タブ
            tempDGV.CurrentCell = null;
            tabData.SelectedIndex = global.TAB_ERR;

            //エラーデータグリッド表示
            lblErr.Text = "エラー件数 ： " + tempDGV.Rows.Count.ToString() + "件";
        }

        /// <summary>
        /// エラー項目バックカラー切替
        /// </summary>
        /// <param name="iX">現在の伝票</param>
        /// <param name="err">エラーテーブル</param>
        private void Show_NGColor(int iX, errCheck.Errtbl[] err)
        {

            if (err[0].Count == 0) return;

            for (int i = 0; i < err.Length; i++)
            {
                if (err[i].DenNo == iX)
                {
                    if (err[i].LINE == 0)
                    {
                        switch (err[i].DpPos)
                        {
                                //日付関連エラー
                            case "txtYear":

                                if (ChkErrColor.Checked == true)
                                {
                                    this.txtYear.BackColor = Color.Yellow;
                                    this.txtMonth.BackColor = Color.Yellow;
                                    this.txtDay.BackColor = Color.Yellow;
                                }
                                else
                                {
                                    this.txtYear.BackColor = Color.Empty;
                                    this.txtMonth.BackColor = Color.Empty;
                                    this.txtDay.BackColor = Color.Empty;
                                }

                                break;

                                //伝票関連エラー
                            case "txtDenNo":

                                if (ChkErrColor.Checked == true)
                                {
                                    this.txtDenNo.BackColor = Color.Yellow;
                                }
                                else
                                {
                                    this.txtDenNo.BackColor = Color.Empty;
                                }

                                break;

                                //複数枚エラー
                            case "fukusu":
                                global.pblFukuColor = Color.Yellow;

                                if (ChkErrColor.Checked == true)
                                {
                                    chkFukusuChk.BackColor = Color.Yellow;
                                }
                                else
                                {
                                    chkFukusuChk.BackColor = SystemColors.Control;
                                }

                                break;

                                //決算エラー
                            case "kessan":
                                global.pblKessanColor = Color.Yellow;

                                if (ChkErrColor.Checked == true)
                                {
                                    ChkKessan.BackColor = Color.Yellow;
                                }
                                else
                                {
                                    ChkKessan.BackColor = SystemColors.Control;
                                }

                                break;

                                //差額エラー
                            case "txtSagaku_T":
                                global.pblSagakuColor = Color.Yellow;

                                if (ChkErrColor.Checked == true)
                                {
                                    gcMultiRow1.ColumnFooters[0].Cells[MultiRow.DP_SAGAKU_T].Style.BackColor = Color.Yellow;
                                }
                                else
                                {
                                    gcMultiRow1.ColumnFooters[0].Cells[MultiRow.DP_SAGAKU_T].Style.BackColor = SystemColors.Control;
                                }

                                break;

                            default:

                                //その他エラー
                                if (ChkErrColor.Checked == true)
                                {
                                    gcMultiRow1[err[i].LINE, err[i].DpPos].Style.BackColor = Color.Yellow;
                                }
                                else
                                {
                                    gcMultiRow1[err[i].LINE, err[i].DpPos].Style.BackColor = Color.Empty;
                                }

                                break;
                        }
                    }
                    else
                    {
                        //その他エラー
                        if (ChkErrColor.Checked == true)
                        {
                            gcMultiRow1[err[i].LINE, err[i].DpPos].Style.BackColor = Color.Yellow;
                        }
                        else
                        {
                            gcMultiRow1[err[i].LINE, err[i].DpPos].Style.BackColor = Color.Empty;
                        }
                    }
                }
            }
        }

        /// <summary>
        /// 伝票画像表示
        /// </summary>
        /// <param name="iX">現在の伝票</param>
        /// <param name="tempImgName">画像名</param>
        public void ShowImage(int iX, string tempImgName)
        {
            string wrkFileName;

            //修正画面へ組み入れた画像フォームの表示    
            //画像の出力が無い場合は、画像表示をしない。
            if (tempImgName == string.Empty)
            {
                leadImg.Visible = false;
                global.pblImageFile = string.Empty;
                return;
            }
      
            //画像ファイルがあるときのみ表示　---> ＣＳＶ分割後のフォルダから取得
            wrkFileName = global.WorkDir + global.DIR_INCSV + tempImgName;
            if (File.Exists(wrkFileName))
            { 
                leadImg.Visible = true;

                //画像ロード
                RasterCodecs.Startup();
                RasterCodecs cs = new RasterCodecs();

                // 描画時に使用される速度、品質、およびスタイルを制御します。 
                RasterPaintProperties prop = new RasterPaintProperties();
                prop = RasterPaintProperties.Default;
                prop.PaintDisplayMode = RasterPaintDisplayModeFlags.Resample;
                leadImg.PaintProperties = prop;

                leadImg.Image = cs.Load(wrkFileName, 0,CodecsLoadByteOrder.BgrOrGray,1,1);

                //画像表示倍率設定
                if (global.miMdlZoomRate == 0f)
                {
                    leadImg.ScaleFactor *= global.ZOOM_RATE;
                }
                else
                {
                    leadImg.ScaleFactor *= global.miMdlZoomRate;
                }

                //画像のマウスによる移動を可能とする
                leadImg.InteractiveMode = RasterViewerInteractiveMode.Pan;

                // グレースケールに変換
                GrayscaleCommand grayScaleCommand = new GrayscaleCommand();
                grayScaleCommand.BitsPerPixel = 8;
                grayScaleCommand.Run(leadImg.Image);
                leadImg.Refresh();

                cs.Dispose();
                RasterCodecs.Shutdown();
                global.pblImageFile = wrkFileName;
            }
            else
            {
                //画像ファイルがないとき
                leadImg.Visible = false;
                global.pblImageFile = string.Empty;
            }
        }

        /// <summary>
        /// 終了処理
        /// </summary>
        /// <param name="st">伝票配列データ</param>
        /// <param name="SaveDataStatus">汎用データ作成実行ステータス true:実行する、false:実行しない</param>
        private void MainEnd(Entity.InputRecord[] st,Boolean SaveDataStatus)
        {
            //frmProg.Caption = "データ変換中・・・"
            //frmProg.prgBar.Value = 60

            //引数のステータスがtrueなら汎用データを作成する
            if (SaveDataStatus == true) SaveData(st);

            //frmProg.prgBar.Value = 100
            //Unload frmProg

            //分割ファイル削除
            utility.FileDelete(global.WorkDir + global.DIR_INCSV,"*");

            //終了メッセージ表示
            if (global.pblDenNum > 0 && SaveDataStatus == true)　
            {
                MessageBox.Show("処理が終了しました。" + Environment.NewLine + "勘定奉行でデータの受入れを行ってください。",Application.ProductName,MessageBoxButtons.OK,MessageBoxIcon.Information);
            }

            this.Tag = "MainEnd";
            this.Close();
        }

        ///------------------------------------------------------
        /// <summary>
        ///     データ出力処理 </summary>
        /// <param name="st">
        ///     伝票配列データ</param>
        ///------------------------------------------------------
        private void SaveData(Entity.InputRecord[] st)
        {
            string wrkOutputData;
            Boolean iniFlg = true;             
            Boolean pblFirstGyouFlg;
            Entity.OutputRecord OutData = new Entity.OutputRecord(); 

            //出力ファイルインスタンス作成
            StreamWriter outFile = new StreamWriter(global.WorkDir + global.DIR_OK + global.tmpFile, false, System.Text.Encoding.GetEncoding(932));

            try
            {
                //伝票データを読み出す
                for (int iX = 0; iX < global.pblDenNum; iX++)
                {
                    //プログレスバー表示
                    //frmProg.Caption = "データ変換中・・・(" & CStr(DenCnt) & "/" & CStr(pblDenNum) & ")"
                    //frmProg.prgBar.Value = CInt((DenCnt / pblDenNum) * 100)
        
                    //伝票最初行フラグ
                    pblFirstGyouFlg = true;

                    for (int i = 0; i < global.MAXGYOU; i++)
                    {
                        //取消行は対象外
                        if (st[iX].Gyou[i].Torikeshi == "0")
                        {
                            //空白行は出力しない（借方貸方両方の科目コードがないとき、又は摘要チェック、摘要がないとき）
                            if (st[iX].Gyou[i].Kari.Kamoku != string.Empty || st[iX].Gyou[i].Kashi.Kamoku != string.Empty || 
                                st[iX].Gyou[i].CopyChk != string.Empty || st[iX].Gyou[i].Tekiyou != string.Empty) 
                            {
                                //出力データ初期化
                                InitOutRec(OutData);

                                //ヘッダファイル出力：2017/09/11
                                if (iniFlg)
                                {
                                    wrkOutputData = string.Empty;
                                    wrkOutputData += Entity.OutPutHeader.dn01 + ",";
                                    wrkOutputData += Entity.OutPutHeader.hd03 + ",";    //整理区分
                                    wrkOutputData += Entity.OutPutHeader.hd02 + ",";
                                    wrkOutputData += Entity.OutPutHeader.hd04 + ",";

                                    wrkOutputData += Entity.OutPutHeader.kr01 + ",";
                                    wrkOutputData += Entity.OutPutHeader.kr02 + ",";
                                    wrkOutputData += Entity.OutPutHeader.kr03 + ",";
                                    wrkOutputData += Entity.OutPutHeader.kr08 + ",";
                                    wrkOutputData += Entity.OutPutHeader.kr10 + ",";
                                    wrkOutputData += Entity.OutPutHeader.kr05 + ",";
                                    wrkOutputData += Entity.OutPutHeader.kr04 + ",";
                                    wrkOutputData += Entity.OutPutHeader.kr11 + ",";

                                    wrkOutputData += Entity.OutPutHeader.ks01 + ",";
                                    wrkOutputData += Entity.OutPutHeader.ks02 + ",";
                                    wrkOutputData += Entity.OutPutHeader.ks03 + ",";
                                    wrkOutputData += Entity.OutPutHeader.ks08 + ",";
                                    wrkOutputData += Entity.OutPutHeader.ks10 + ",";
                                    wrkOutputData += Entity.OutPutHeader.ks05 + ",";
                                    wrkOutputData += Entity.OutPutHeader.ks04 + ",";
                                    wrkOutputData += Entity.OutPutHeader.ks11 + ",";

                                    wrkOutputData += Entity.OutPutHeader.tk01;

                                    outFile.WriteLine(wrkOutputData);
                                    iniFlg = false;


                                    //sb.Append(OutData.Kugiri).Append(",");
                                    //sb.Append(OutData.Kessan).Append(",");
                                    //sb.Append(OutData.Date).Append(",");
                                    //sb.Append(OutData.DenNo).Append(",");

                                    //sb.Append(OutData.Kari.Bumon).Append(",");
                                    //sb.Append(OutData.Kari.Kamoku).Append(",");
                                    //sb.Append(OutData.Kari.Hojo).Append(",");
                                    //sb.Append(OutData.Kari.Kin).Append(",");
                                    //sb.Append(OutData.Kari.Tax).Append(",");
                                    //sb.Append(OutData.Kari.TaxMas).Append(",");
                                    //sb.Append(OutData.Kari.TaxKbn).Append(",");
                                    //sb.Append(OutData.Kari.JigyoKbn).Append(",");

                                    //sb.Append(OutData.Kashi.Bumon).Append(",");
                                    //sb.Append(OutData.Kashi.Kamoku).Append(",");
                                    //sb.Append(OutData.Kashi.Hojo).Append(",");
                                    //sb.Append(OutData.Kashi.Kin).Append(",");
                                    //sb.Append(OutData.Kashi.Tax).Append(",");
                                    //sb.Append(OutData.Kashi.TaxMas).Append(",");
                                    //sb.Append(OutData.Kashi.TaxKbn).Append(",");
                                    //sb.Append(OutData.Kashi.JigyoKbn).Append(",");

                                    //sb.Append(OutData.Tekiyou);
                                }

                                //出力データ作成
                                wrkOutputData = SetData(iX, i, st,pblFirstGyouFlg,OutData);
        
                                //一時ファイルへ出力            
                                outFile.WriteLine(wrkOutputData);
                                pblFirstGyouFlg = false;
                            }
                        }
                    }
                }

                //ファイルクローズ
                outFile.Close();

                //出力ファイル削除
                utility.FileDelete(global.WorkDir + global.DIR_OK, global.OUTFILE);

                //一時ファイルを出力ファイルにコピー
                File.Copy(global.WorkDir + global.DIR_OK + global.tmpFile, global.WorkDir + global.DIR_OK + global.OUTFILE);

                //一時ファイル削除
                utility.FileDelete(global.WorkDir + global.DIR_OK, global.tmpFile);
            }
            catch (Exception e)
            {
                MessageBox.Show("データ変換中" + Environment.NewLine + e.Message, "エラー", MessageBoxButtons.OK);
            }
        }

        /// <summary>
        /// 出力用データ初期化
        /// </summary>
        /// <param name="OutData">出力用データ</param>
        private void InitOutRec(Entity.OutputRecord OutData)
        {
            OutData.Kugiri = string.Empty;
            OutData.Kessan = string.Empty;
            OutData.Date = string.Empty;
            OutData.DenNo = string.Empty;
            OutData.Kari.Bumon = string.Empty;
            OutData.Kari.Kamoku = string.Empty;
            OutData.Kari.Hojo = string.Empty;
            OutData.Kari.Kin = string.Empty;
            OutData.Kari.Tax = string.Empty;
            OutData.Kari.TaxMas = string.Empty;
            OutData.Kari.TaxKbn = string.Empty;
            OutData.Kari.JigyoKbn = string.Empty;
            OutData.Kashi.Bumon = string.Empty;
            OutData.Kashi.Kamoku = string.Empty;
            OutData.Kashi.Hojo = string.Empty;
            OutData.Kashi.Kin = string.Empty;
            OutData.Kashi.Tax = string.Empty;
            OutData.Kashi.TaxMas = string.Empty;
            OutData.Kashi.TaxKbn = string.Empty;
            OutData.Kashi.JigyoKbn = string.Empty;
            OutData.Tekiyou = string.Empty;
            OutData.Arrange = string.Empty;
        }

        ///----------------------------------------------------
        /// <summary>
        ///     出力データ作成 </summary>
        /// <param name="iX">
        ///     伝票添え字</param>
        /// <param name="i">
        ///     行添え字</param>
        /// <param name="st">
        ///     伝票配列データ</param>
        /// <param name="fFlg">
        ///     最初データフラグ</param>
        /// <param name="OutData">
        ///     出力データ</param>
        /// <returns>
        ///     出力データ文字列</returns>
        ///----------------------------------------------------
        private string SetData(int iX, int i, Entity.InputRecord[] st, Boolean fFlg, Entity.OutputRecord OutData)
        {
            //伝票区切
            //複数チェックなし　かつ　伝票最初の行のみ
            if (st[iX].Head.FukusuChk == "0" && fFlg == true)
            {
                OutData.Kugiri = "*";
            }
            else
            {
                OutData.Kugiri = string.Empty;
            }
        
            //決算処理フラグ
            ////チェックがあれば"1"、無ければstring.Empty
            //if (st[iX].Head.Kessan == "1")
            //{
            //    OutData.Kessan = "1";
            //}
            //else
            //{
            //    OutData.Kessan = string.Empty;
            //}

            //整理区分　2017/09/11
            //決算チェックありで勘定奉行の整理仕訳区分が"0"のとき：１、それ以外は０
            if (st[iX].Head.Kessan == global.FLGON && company.Arrange == global.FLGON)
            {
                OutData.Arrange = global.FLGON;
            }
            else
            {
                OutData.Arrange = global.FLGOFF;
            }

            ////年月日結合
            //OutData.Date = string.Format("{0:00}",int.Parse(st[iX].Head.Year)) + string.Format("{0:00}",int.Parse(st[iX].Head.Month)) + string.Format("{0:00}",int.Parse(st[iX].Head.Day));


            //日付 : 2017/09/12
            int sYear = int.Parse(st[iX].Head.Year);

            //西暦を求める
            if (global.pblReki == global.RWAREKI) //和暦のとき
            {
                sYear = sYear + int.Parse(company.Hosei);
            }
            else
            {
                sYear = sYear + 2000;
            }

            OutData.Date = sYear.ToString() + "/" + st[iX].Head.Month.PadLeft(2, '0') + "/" + st[iX].Head.Day.PadLeft(2, '0');

            //伝票番号
            OutData.DenNo = st[iX].Head.DenNo;
        
            //借方部門
            OutData.Kari.Bumon = st[iX].Gyou[i].Kari.Bumon;
    
            //借方科目
            OutData.Kari.Kamoku = st[iX].Gyou[i].Kari.Kamoku;
    
            //借方補助
            OutData.Kari.Hojo = st[iX].Gyou[i].Kari.Hojo;
    
            //借方金額
            OutData.Kari.Kin = st[iX].Gyou[i].Kari.Kin;
    
            //借方消費税額
            OutData.Kari.Tax = string.Empty;
    
            //借方消費税額計算区分
            if (st[iX].Gyou[i].Kari.TaxMas == string.Empty)
            {
                OutData.Kari.TaxMas = fncGetZeiFlag(OutData.Kari.Kamoku);
            }
            else if (st[iX].Gyou[i].Kari.TaxMas == "0")
            {
                OutData.Kari.TaxMas = string.Empty;
            }
            else
            {
                OutData.Kari.TaxMas = st[iX].Gyou[i].Kari.TaxMas;
            }
    
            //借方消費税区分
            if (st[iX].Gyou[i].Kari.TaxKbn == string.Empty)
            {
                OutData.Kari.TaxKbn = string.Empty;
            }
            else
            {
                OutData.Kari.TaxKbn = st[iX].Gyou[i].Kari.TaxKbn;
            }
    
            //借方事業区分
            OutData.Kari.JigyoKbn = string.Empty;
    
            //貸方部門
            OutData.Kashi.Bumon = st[iX].Gyou[i].Kashi.Bumon;

            //貸方科目
            OutData.Kashi.Kamoku = st[iX].Gyou[i].Kashi.Kamoku;

            //貸方補助
            OutData.Kashi.Hojo = st[iX].Gyou[i].Kashi.Hojo;
    
            //貸方金額
            OutData.Kashi.Kin = st[iX].Gyou[i].Kashi.Kin;
    
            //貸方消費税額
            OutData.Kashi.Tax = string.Empty;
    
            //貸方消費税額計算区分
            if (st[iX].Gyou[i].Kashi.TaxMas == string.Empty)
            {
                OutData.Kashi.TaxMas = fncGetZeiFlag(OutData.Kashi.Kamoku);
            }
            else if (st[iX].Gyou[i].Kashi.TaxMas == "0")
            {
                OutData.Kashi.TaxMas = string.Empty;
            }
            else
            {
                OutData.Kashi.TaxMas = st[iX].Gyou[i].Kashi.TaxMas;
            }
    
            //貸方消費税区分
            if (st[iX].Gyou[i].Kashi.TaxKbn == string.Empty)
            {
                OutData.Kashi.TaxKbn = string.Empty;
            }
            else
            {
                OutData.Kashi.TaxKbn = st[iX].Gyou[i].Kashi.TaxKbn;
            }
    
            //貸方事業区分
            OutData.Kashi.JigyoKbn = string.Empty;
        
            //摘要 
            OutData.Tekiyou = st[iX].Gyou[i].Tekiyou.TrimEnd();

            //出力文字列作成
            StringBuilder sb = new StringBuilder();
            sb.Append(OutData.Kugiri).Append(",");
            //sb.Append(OutData.Kessan).Append(",");
            sb.Append(OutData.Arrange).Append(",");
            sb.Append(OutData.Date).Append(",");
            sb.Append(OutData.DenNo).Append(",");
            sb.Append(OutData.Kari.Bumon).Append(",");
            sb.Append(OutData.Kari.Kamoku).Append(",");
            sb.Append(OutData.Kari.Hojo).Append(",");
            sb.Append(OutData.Kari.Kin).Append(",");
            sb.Append(OutData.Kari.Tax).Append(",");
            sb.Append(OutData.Kari.TaxMas).Append(",");
            sb.Append(OutData.Kari.TaxKbn).Append(",");
            sb.Append(OutData.Kari.JigyoKbn).Append(",");
            sb.Append(OutData.Kashi.Bumon).Append(",");
            sb.Append(OutData.Kashi.Kamoku).Append(",");
            sb.Append(OutData.Kashi.Hojo).Append(",");
            sb.Append(OutData.Kashi.Kin).Append(",");
            sb.Append(OutData.Kashi.Tax).Append(",");
            sb.Append(OutData.Kashi.TaxMas).Append(",");
            sb.Append(OutData.Kashi.TaxKbn).Append(",");
            sb.Append(OutData.Kashi.JigyoKbn).Append(",");
            sb.Append(OutData.Tekiyou);

            return sb.ToString();
        }

        ///---------------------------------------------------------
        /// <summary>
        ///     総勘定科目税処理フラグ取得 : 2017/06/06 </summary>
        /// <param name="kCode">
        ///     勘定科目コード</param>
        /// <returns>
        ///     税処理フラグ</returns>
        ///---------------------------------------------------------
        private string fncGetZeiFlag(string kCode)
        {
            string sRet = string.Empty;

            for (int i = 0; i < fgKamoku.Rows.Count; i++)
            {
                if (fgKamoku[0, i].Value.ToString() == kCode)
                {
                    sRet = fgKamoku[4, i].Value.ToString();
                    break;
                }
            }

            return sRet;

            //string sRet = string.Empty;

            //// 勘定奉行データベース接続文字列を取得する
            //string sc = SqlControl.obcConnectSting.get(global.pblDbName);

            //// 勘定奉行データベースへ接続する
            //SqlControl.DataControl dCon = new SqlControl.DataControl(sc);

            ////科目データ取得
            ////データリーダーを取得する
            //SqlDataReader dR;
            //string sqlSTRING = string.Empty;
            ////sqlSTRING += "SELECT sUcd,sNcd,sNm,tiIsTrk,tiIsZei FROM wkskm01 ";
            ////sqlSTRING += "WHERE tiIsTrk = 1 and sUcd = '" + string.Format("{0,6}",kCode) + "' ";

            //dR = dCon.free_dsReader(sqlSTRING);
            //while (dR.Read())
            //{
            //    if (dR["tiIsZei"].ToString() == "0" || dR["tiIsZei"].ToString() == "1")
            //    {
            //        sRet = string.Empty;
            //    }
            //    else if (dR["tiIsZei"].ToString() == "2")
            //    {
            //        sRet = "1";
            //    }
            //}

            //dR.Close();
            //dCon.Close();

            //return sRet;
        }

        private void Base_FormClosing(object sender, FormClosingEventArgs e)
        {
            int wrkImageHeight = 0;
            int wrkImageWidth = 0;
            int wrkImageX = 0;
            string mySql = string.Empty;

            try 
            {	            
                //×ボタン押下または終了ボタン
                if (e.CloseReason == CloseReason.UserClosing)
                {
                    if (this.Tag.ToString() != "MainEnd")
                    {
                        if (MessageBox.Show("処理中の伝票データをすべて削除して終了します。",Application.ProductName, MessageBoxButtons.YesNo,MessageBoxIcon.Question,MessageBoxDefaultButton.Button2) == DialogResult.Yes)
                        {
                            //Control.FreeSql dc = new Control.FreeSql(global.WorkDir + global.DIR_CONFIG, global.CONFIGFILE);

                            // ACCESSデータベースへ接続
                            SysControl.SetDBConnect Con = new SysControl.SetDBConnect();
                            OleDbCommand sCom = new OleDbCommand();
                            sCom.Connection = Con.cnOpen();

                            mySql = "update Config set ";
                            mySql += "ImgH = '" + wrkImageHeight.ToString() + "',";
                            mySql += "ImgW = '" + wrkImageWidth.ToString() + "',";
                            mySql += "ImgX = '" + wrkImageX + "',";
                            mySql += "sub2 = 0";

                            sCom.CommandText = mySql;
                            sCom.ExecuteNonQuery();

                            // データベース切断
                            sCom.Connection.Close();
    
                            MessageBox.Show("プログラムを終了します。",Application.ProductName,MessageBoxButtons.OK,MessageBoxIcon.Information);
                            utility.FileDelete(global.WorkDir + global.DIR_INCSV,"*");
                        }
                        else
                        {
                            e.Cancel = true;
                            this.Tag = string.Empty;
                            return;
                        }
                    }
                }

                this.Dispose();
	        }
	        catch (Exception ex)
	        {
                MessageBox.Show("画像表示データ書込み中" + Environment.NewLine + ex.Message,Application.ProductName,MessageBoxButtons.OK,MessageBoxIcon.Exclamation);
	        }
        }

        private void gcMultiRow1_CellValueChanged(object sender, GrapeCity.Win.MultiRow.CellEventArgs e)
        {
            string sWorkB;
            string sWorkA1;
            string sWorkA2;
            string CngVal;
    
            if (bCngFlag == true) return; //多重処理を避ける
            
            switch (e.CellName)
            {
                case "txtTekiyou":  //摘要
                    bCngFlag = true;
                    if (gcMultiRow1.GetValue(e.RowIndex, e.CellName) != null)
                    {
                        sWorkB = gcMultiRow1.GetValue(e.RowIndex, e.CellName).ToString();
                    }
                    else
                    {
                        sWorkB = string.Empty;
                    }

                    sWorkA1 = string.Empty;
                    sWorkA2 = string.Empty;

                    for (int i = 0; i < sWorkB.Length ; i++)
			        {
			            sWorkA1 += sWorkB.Substring(i, 1);

                        if (System.Text.Encoding.GetEncoding("Shift_JIS").GetByteCount(sWorkA1) <= 40)
                        {
                            sWorkA2 += sWorkB.Substring(i, 1);
                        }
			        }

                    gcMultiRow1.SetValue(e.RowIndex, "txtTekiyou",sWorkA2);
                    bCngFlag = false;
                    break;

                case "txtKin_K":    //--借方金額カンマ編集・変換データ適切チェック処理 (金額)
                    bCngFlag = true;
                    KinCellValueChange(e);
                    bCngFlag = false;
                    break;

                case "txtKin_S":    //--貸方金額カンマ編集・変換データ適切チェック処理 (金額)
                    bCngFlag = true;
                    KinCellValueChange(e);
                    bCngFlag = false;
                    break;

                case "txtKCode_K":  //借方勘定科目
                
                    bCngFlag = true;
                    if (gcMultiRow1.GetValue(e.RowIndex, e.CellName) != null)
                    {
                        CngVal = gcMultiRow1.GetValue(e.RowIndex, e.CellName).ToString().Trim();          //両端空白削除
                        CngVal = CngVal.Replace(" ", string.Empty);                                         //文中空白削除
                        gcMultiRow1.SetValue(e.RowIndex, e.CellName, CngVal.Replace("-", string.Empty));    //"-"ハイフン削除

                        if (utility.NumericCheck(gcMultiRow1.GetValue(e.RowIndex, e.CellName).ToString()))
                        {
                            gcMultiRow1.SetValue(e.RowIndex, e.CellName, string.Format("{0:###0}", int.Parse(gcMultiRow1.GetValue(e.RowIndex, e.CellName).ToString())));
                        }
                        //勘定科目名表示
                        if (gcMultiRow1.GetValue(e.RowIndex, e.CellName).ToString() != string.Empty)
                        {
                            GetKamokeName(gcMultiRow1, e.RowIndex, "txtKName_K", "txtKCode_K");
                        }
                    }
                    else
                    {
                        gcMultiRow1.SetValue(e.RowIndex, e.CellName, string.Empty);
                        gcMultiRow1.SetValue(e.RowIndex, "txtKName_K", string.Empty);
                    }

                    //MessageBox.Show(e.RowIndex.ToString() + " : " + gcMultiRow1.GetValue(e.RowIndex, e.CellName).ToString());

                    bCngFlag = false;
                
                    break;
                                    
                case "txtHojo_K":   //借方補助コード
                 
                    bCngFlag = true;
                    if (gcMultiRow1.GetValue(e.RowIndex, e.CellName) != null)
                    {
                        CngVal = gcMultiRow1.GetValue(e.RowIndex, e.CellName).ToString().Trim();                   //両端空白削除
                        CngVal = CngVal.Replace(" ", string.Empty);                                                 //文中空白削除
                        gcMultiRow1.SetValue(e.RowIndex, e.CellName, CngVal.Replace("-", string.Empty));   //"-"ハイフン削除

                        if (utility.NumericCheck(gcMultiRow1.GetValue(e.RowIndex, e.CellName).ToString()))
                        {
                            gcMultiRow1.SetValue(e.RowIndex, e.CellName, string.Format("{0:###0}", int.Parse(gcMultiRow1.GetValue(e.RowIndex, e.CellName).ToString())));
                        }

                        //補助科目名表示
                        GetHojoName(gcMultiRow1, e.RowIndex, "txtKName_K", "txtKCode_K", "txtHojoName_K", "txtHojo_K");
                    }
                    else
                    {
                        gcMultiRow1.SetValue(e.RowIndex, e.CellName, string.Empty);
                        gcMultiRow1.SetValue(e.RowIndex, "txtHojoName_K", string.Empty);
                    }
                    
                    bCngFlag = false;           
                    break;

                case "txtBCode_K":  //借方部門コード
                
                    bCngFlag = true;
                    if (gcMultiRow1.GetValue(e.RowIndex, e.CellName) != null)
                    {
                        CngVal = gcMultiRow1.GetValue(e.RowIndex, e.CellName).ToString().Trim();                      //両端空白削除
                        CngVal = CngVal.Replace(" ", string.Empty);  //文中空白削除
                        gcMultiRow1.SetValue(e.RowIndex, e.CellName, CngVal.Replace("-", string.Empty));  //"-"ハイフン削除

                        if (utility.NumericCheck(gcMultiRow1.GetValue(e.RowIndex, e.CellName).ToString()))
                        {
                            gcMultiRow1.SetValue(e.RowIndex, e.CellName, string.Format("{0:###0}", int.Parse(gcMultiRow1.GetValue(e.RowIndex, e.CellName).ToString())));
                        }

                        //部門名表示
                        if (gcMultiRow1.GetValue(e.RowIndex, e.CellName).ToString() != string.Empty)
                        {
                            GetBumonName(gcMultiRow1, e.RowIndex, "txtBName_K", "txtBCode_K");
                        }
                    }
                    else
                    {
                        gcMultiRow1.SetValue(e.RowIndex, e.CellName, string.Empty);
                        gcMultiRow1.SetValue(e.RowIndex, "txtBName_K", string.Empty);
                    }
                    
                    bCngFlag = false;  

                    break;
                    
                case "txtZeik_K":   //借方税区分コード
                
                    bCngFlag = true;
                    if (gcMultiRow1.GetValue(e.RowIndex, e.CellName) != null)
                    {
                        CngVal = gcMultiRow1.GetValue(e.RowIndex, e.CellName).ToString().Trim();        //両端空白削除
                        CngVal = CngVal.Replace(" ", string.Empty);                                     //文中空白削除
                        gcMultiRow1.SetValue(e.RowIndex, e.CellName, CngVal.Replace("-", string.Empty)); //"-"ハイフン削除

                        if (utility.NumericCheck(gcMultiRow1.GetValue(e.RowIndex, e.CellName).ToString()))
                        {
                            gcMultiRow1.SetValue(e.RowIndex, e.CellName, string.Format("{0:#0}",int.Parse(gcMultiRow1.GetValue(e.RowIndex, e.CellName).ToString())));
                        }
                    }
                    else
                    {
                        gcMultiRow1.SetValue(e.RowIndex, e.CellName, string.Empty);
                    }

                    bCngFlag = false;
                    break;

                case "txtZeis_K":   //税処理

                    bCngFlag = true;

                    if (gcMultiRow1.GetValue(e.RowIndex, e.CellName) == null)
                    {
                        gcMultiRow1.SetValue(e.RowIndex, e.CellName, string.Empty);
                    }
                    bCngFlag = false;
                    break;

                case "txtKCode_S":  //貸方勘定科目
                
                    bCngFlag = true;
                    if (gcMultiRow1.GetValue(e.RowIndex, e.CellName) != null)
                    {
                        CngVal = gcMultiRow1.GetValue(e.RowIndex, e.CellName).ToString().Trim();          //両端空白削除
                        CngVal = CngVal.Replace(" ", string.Empty);                                         //文中空白削除
                        gcMultiRow1.SetValue(e.RowIndex, e.CellName, CngVal.Replace("-", string.Empty));    //"-"ハイフン削除

                        if (utility.NumericCheck(gcMultiRow1.GetValue(e.RowIndex, e.CellName).ToString()))
                        {
                            gcMultiRow1.SetValue(e.RowIndex, e.CellName, string.Format("{0:###0}", int.Parse(gcMultiRow1.GetValue(e.RowIndex, e.CellName).ToString())));
                        }
                        //勘定科目名表示
                        if (gcMultiRow1.GetValue(e.RowIndex, e.CellName).ToString() != string.Empty)
                        {
                            GetKamokeName(gcMultiRow1, e.RowIndex, "txtKName_S", "txtKCode_S");
                        }
                    }
                    else
                    {
                        gcMultiRow1.SetValue(e.RowIndex, e.CellName, string.Empty);
                        gcMultiRow1.SetValue(e.RowIndex, "txtKName_S", string.Empty);
                    }
                    
                    bCngFlag = false;

                    break;
                                    
                case "txtHojo_S":   //貸方補助コード
                 
                    bCngFlag = true;
                    if (gcMultiRow1.GetValue(e.RowIndex, e.CellName) != null)
                    {
                        CngVal = gcMultiRow1.GetValue(e.RowIndex, e.CellName).ToString().Trim();                   //両端空白削除
                        CngVal = CngVal.Replace(" ", string.Empty);                                                 //文中空白削除
                        gcMultiRow1.SetValue(e.RowIndex, e.CellName, CngVal.Replace("-", string.Empty));   //"-"ハイフン削除

                        if (utility.NumericCheck(gcMultiRow1.GetValue(e.RowIndex, e.CellName).ToString()))
                        {
                            gcMultiRow1.SetValue(e.RowIndex, e.CellName, string.Format("{0:###0}", int.Parse(gcMultiRow1.GetValue(e.RowIndex, e.CellName).ToString())));
                        }

                        //補助科目名表示
                        GetHojoName(gcMultiRow1, e.RowIndex, "txtKName_S", "txtKCode_S", "txtHojoName_S", "txtHojo_S");
                    }
                    else
                    {
                        gcMultiRow1.SetValue(e.RowIndex, e.CellName, string.Empty);
                        gcMultiRow1.SetValue(e.RowIndex, "txtHojoName_S", string.Empty);
                    }
                    
                    bCngFlag = false;           
                    break;

                case "txtBCode_S":  //貸方部門コード
                
                    bCngFlag = true;
                    if (gcMultiRow1.GetValue(e.RowIndex, e.CellName) != null)
                    {
                        CngVal = gcMultiRow1.GetValue(e.RowIndex, e.CellName).ToString().Trim();                      //両端空白削除
                        CngVal = CngVal.Replace(" ", string.Empty);  //文中空白削除
                        gcMultiRow1.SetValue(e.RowIndex, e.CellName, CngVal.Replace("-", string.Empty));  //"-"ハイフン削除

                        if (utility.NumericCheck(gcMultiRow1.GetValue(e.RowIndex, e.CellName).ToString()))
                        {
                            gcMultiRow1.SetValue(e.RowIndex, e.CellName, string.Format("{0:###0}", int.Parse(gcMultiRow1.GetValue(e.RowIndex, e.CellName).ToString())));
                        }

                        //部門名表示
                        if (gcMultiRow1.GetValue(e.RowIndex, e.CellName).ToString() != string.Empty)
                        {
                            GetBumonName(gcMultiRow1, e.RowIndex, "txtBName_S", "txtBCode_S");
                        }
                    }
                    else
                    {
                        gcMultiRow1.SetValue(e.RowIndex, e.CellName, string.Empty);
                        gcMultiRow1.SetValue(e.RowIndex, "txtBName_S", string.Empty);
                    }
                    
                    bCngFlag = false;  

                    break;
                    
                case "txtZeik_S":   //貸方税区分コード
                
                    bCngFlag = true;
                    if (gcMultiRow1.GetValue(e.RowIndex, e.CellName) != null)
                    {
                        CngVal = gcMultiRow1.GetValue(e.RowIndex, e.CellName).ToString().Trim();        //両端空白削除
                        CngVal = CngVal.Replace(" ", string.Empty);                                     //文中空白削除
                        gcMultiRow1.SetValue(e.RowIndex, e.CellName, CngVal.Replace("-", string.Empty)); //"-"ハイフン削除

                        if (utility.NumericCheck(gcMultiRow1.GetValue(e.RowIndex, e.CellName).ToString()))
                        {
                            gcMultiRow1.SetValue(e.RowIndex, e.CellName, string.Format("{0:#0}",int.Parse(gcMultiRow1.GetValue(e.RowIndex, e.CellName).ToString())));
                        }
                    }
                    else
                    {
                        gcMultiRow1.SetValue(e.RowIndex, e.CellName, string.Empty);
                    }

                    bCngFlag = false;
                    break;

                case "txtZeis_S":   //貸方税処理

                    bCngFlag = true;

                    if (gcMultiRow1.GetValue(e.RowIndex, e.CellName) == null)
                    {
                        gcMultiRow1.SetValue(e.RowIndex, e.CellName, string.Empty);
                    }
                    bCngFlag = false;
                    break;

		        default:
                    break;
	        }

            //1行目の摘要複写チェック欄はチェック不可
            chkFu0.Enabled = false;
        }

        /// <summary>
        /// 金額セル値変更時の処理
        /// </summary>
        /// <param name="e">CellEventArgs</param>
        private void KinCellValueChange(GrapeCity.Win.MultiRow.CellEventArgs e)
        {
            if (gcMultiRow1.GetValue(e.RowIndex, e.CellName) != null)
            {
                if (errCheck.ChkKinIndi(gcMultiRow1.GetValue(e.RowIndex, e.CellName).ToString()))
                {
                    if (gcMultiRow1.GetValue(e.RowIndex, e.CellName).ToString() != string.Empty)
                    {
                        gcMultiRow1.SetValue(e.RowIndex, e.CellName, string.Format("{0:#,###}", int.Parse(gcMultiRow1.GetValue(e.RowIndex, e.CellName).ToString().Replace(",", string.Empty))));
                    }
                }
            }
            else
            {
                gcMultiRow1.SetValue(e.RowIndex, e.CellName, string.Empty);
            }
        }

        private void cmdExit_Click(object sender, EventArgs e)
        {
            //処理終了
            this.Tag = "cmdExit";
            this.Close();
        }

        private void btnFirst_Click(object sender, EventArgs e)
        {
            //Firstボタン「｜＜＜」クリック時
            DlgDataGet();
            DenIndex = 0;
            DataShow(DenIndex, DenData, eTbl);
        }

        private void btnBefore_Click(object sender, EventArgs e)
        {
            //前ボタン「＜」クリック時
            DlgDataGet();
            DenIndex --;
            DataShow(DenIndex, DenData, eTbl);
        }

        private void btnNext_Click(object sender, EventArgs e)
        {
            //次ボタン「＞」クリック時
            DlgDataGet();
            DenIndex ++;
            DataShow(DenIndex, DenData, eTbl);
        }

        private void btnEnd_Click(object sender, EventArgs e)
        {
            //ENDボタン「＜」クリック時
            DlgDataGet();
            DenIndex = global.pblDenNum - 1;
            DataShow(DenIndex, DenData, eTbl);
        }

        /// <summary>
        /// 表示中の伝票データを配列に取り込む
        /// </summary>
        private void DlgDataGet()
        {
            //伝票ヘッダ
            DenData[DenIndex].Head.Year = string.Format("{0:00}", txtYear.Text);
            DenData[DenIndex].Head.Month = string.Format("{0:00}", txtMonth.Text);
            DenData[DenIndex].Head.Day = string.Format("{0:00}", txtDay.Text);
        
            if (utility.NumericCheck(txtDenNo.Text))
            {
                DenData[DenIndex].Head.DenNo = txtDenNo.Text;
            }
            else
            {
                DenData[DenIndex].Head.DenNo = string.Empty;
            }
        
            //決算
            if (ChkKessan.Checked == true)
            {
                DenData[DenIndex].Head.Kessan = "1";
            }
            else
            {
                DenData[DenIndex].Head.Kessan = "0";
            }

            //複数枚チェック
            if (chkFukusuChk.Checked == true)
            {
                DenData[DenIndex].Head.FukusuChk = "1";
            }
            else
            {
                DenData[DenIndex].Head.FukusuChk = "0";
            }
    
            //行データ
            for (int i = 0; i < global.MAXGYOU; i++)
            {
                //摘要複写
                if (i == 0)
                {
                    if (chkFu0.Checked == true)
                    {
                        DenData[DenIndex].Gyou[i].CopyChk = "1";
                    }
                    else
                    {
                        DenData[DenIndex].Gyou[i].CopyChk = "0";
                    }
                }

                if (i == 1)
                {
                    if (chkFu1.Checked == true)
                    {
                        DenData[DenIndex].Gyou[i].CopyChk = "1";
                    }
                    else
                    {
                        DenData[DenIndex].Gyou[i].CopyChk = "0";
                    }
                }

                if (i == 2)
                {
                    if (chkFu2.Checked == true)
                    {
                        DenData[DenIndex].Gyou[i].CopyChk = "1";
                    }
                    else
                    {
                        DenData[DenIndex].Gyou[i].CopyChk = "0";
                    }
                }

                if (i == 3)
                {
                    if (chkFu3.Checked == true)
                    {
                        DenData[DenIndex].Gyou[i].CopyChk = "1";
                    }
                    else
                    {
                        DenData[DenIndex].Gyou[i].CopyChk = "0";
                    }
                }

                if (i == 4)
                {
                    if (chkFu4.Checked == true)
                    {
                        DenData[DenIndex].Gyou[i].CopyChk = "1";
                    }
                    else
                    {
                        DenData[DenIndex].Gyou[i].CopyChk = "0";
                    }
                }

                if (i == 5)
                {
                    if (chkFu5.Checked == true)
                    {
                        DenData[DenIndex].Gyou[i].CopyChk = "1";
                    }
                    else
                    {
                        DenData[DenIndex].Gyou[i].CopyChk = "0";
                    }
                }

                if (i == 6)
                {
                    if (chkFu6.Checked == true)
                    {
                        DenData[DenIndex].Gyou[i].CopyChk = "1";
                    }
                    else
                    {
                        DenData[DenIndex].Gyou[i].CopyChk = "0";
                    }
                }
                
                //取消チェック
                if (i == 0)
                {
                    if (chkDel0.Checked == true)
                    {
                        DenData[DenIndex].Gyou[i].Torikeshi = "1";
                    }
                    else
                    {
                        DenData[DenIndex].Gyou[i].Torikeshi = "0";
                    }
                }

                if (i == 1)
                {
                    if (chkDel1.Checked == true)
                    {
                        DenData[DenIndex].Gyou[i].Torikeshi = "1";
                    }
                    else
                    {
                        DenData[DenIndex].Gyou[i].Torikeshi = "0";
                    }
                }
                if (i == 2)
                {
                    if (chkDel2.Checked == true)
                    {
                        DenData[DenIndex].Gyou[i].Torikeshi = "1";
                    }
                    else
                    {
                        DenData[DenIndex].Gyou[i].Torikeshi = "0";
                    }
                }
                if (i == 3)
                {
                    if (chkDel3.Checked == true)
                    {
                        DenData[DenIndex].Gyou[i].Torikeshi = "1";
                    }
                    else
                    {
                        DenData[DenIndex].Gyou[i].Torikeshi = "0";
                    }
                }
                if (i == 4)
                {
                    if (chkDel4.Checked == true)
                    {
                        DenData[DenIndex].Gyou[i].Torikeshi = "1";
                    }
                    else
                    {
                        DenData[DenIndex].Gyou[i].Torikeshi = "0";
                    }
                }
                if (i == 5)
                {
                    if (chkDel5.Checked == true)
                    {
                        DenData[DenIndex].Gyou[i].Torikeshi = "1";
                    }
                    else
                    {
                        DenData[DenIndex].Gyou[i].Torikeshi = "0";
                    }
                }
                if (i == 6)
                {
                    if (chkDel6.Checked == true)
                    {
                        DenData[DenIndex].Gyou[i].Torikeshi = "1";
                    }
                    else
                    {
                        DenData[DenIndex].Gyou[i].Torikeshi = "0";
                    }
                }

                //借方
                DenData[DenIndex].Gyou[i].Kari.Bumon = gcMultiRow1.GetValue(i,MultiRow.DP_KARI_CODEB).ToString();
                DenData[DenIndex].Gyou[i].Kari.Kamoku = gcMultiRow1.GetValue(i,MultiRow.DP_KARI_CODE).ToString();
                DenData[DenIndex].Gyou[i].Kari.Hojo= gcMultiRow1.GetValue(i,MultiRow.DP_KARI_CODEH).ToString();
                DenData[DenIndex].Gyou[i].Kari.Kin = gcMultiRow1.GetValue(i,MultiRow.DP_KARI_KIN).ToString().Replace(",",string.Empty);
                DenData[DenIndex].Gyou[i].Kari.TaxMas = gcMultiRow1.GetValue(i,MultiRow.DP_KARI_ZEI_S).ToString();
                DenData[DenIndex].Gyou[i].Kari.TaxKbn = gcMultiRow1.GetValue(i,MultiRow.DP_KARI_ZEI).ToString();
                
                //貸方
                DenData[DenIndex].Gyou[i].Kashi.Bumon = gcMultiRow1.GetValue(i,MultiRow.DP_KASHI_CODEB).ToString();
                DenData[DenIndex].Gyou[i].Kashi.Kamoku = gcMultiRow1.GetValue(i,MultiRow.DP_KASHI_CODE).ToString();
                DenData[DenIndex].Gyou[i].Kashi.Hojo= gcMultiRow1.GetValue(i,MultiRow.DP_KASHI_CODEH).ToString();
                DenData[DenIndex].Gyou[i].Kashi.Kin = gcMultiRow1.GetValue(i,MultiRow.DP_KASHI_KIN).ToString().Replace(",",string.Empty);
                DenData[DenIndex].Gyou[i].Kashi.TaxMas = gcMultiRow1.GetValue(i,MultiRow.DP_KASHI_ZEI_S).ToString();
                DenData[DenIndex].Gyou[i].Kashi.TaxKbn = gcMultiRow1.GetValue(i,MultiRow.DP_KASHI_ZEI).ToString();
  
                DenData[DenIndex].Gyou[i].Tekiyou = gcMultiRow1.GetValue(i,MultiRow.DP_TEKIYOU).ToString();
            }
        }

        private void hScrollBar1_Scroll(object sender, ScrollEventArgs e)
        {
            //横スクロールバー操作時
    
            //同じ伝票のときは無視
            if (hScrollBar1.Value == DenIndex) return;

            //ダイアログ入力データ取得
            DlgDataGet();
            DenIndex = hScrollBar1.Value;
            DataShow(DenIndex, DenData, eTbl);
        }

        private void btnDltDen_Click(object sender, EventArgs e)
        {
            //伝票削除ボタンクリック時
            string DeleteDenNo;
            string DeleteDen;
            string CSV_Name;
            string ImgNAME;
            int sCnt;
    
            //確認
            //キャンセル
            if (MessageBox.Show("削除する伝票を印刷しますか？","削除",MessageBoxButtons.YesNo,MessageBoxIcon.Question) == DialogResult.No)
            {   
                return;
            }
            else 
            {
                //現在の伝票データを取得
                DlgDataGet();

                //一枚印刷
                cPrint Prn = new cPrint();
                Prn.Denpyo(DenData, DenIndex, global.PRINTMODEONE);
            }

            //伝票削除 
            //結合伝票のとき削除対象の先頭伝票をさがす
            if (DenIndex > 0)
            {
                sCnt = DenIndex;

                while (DenData[sCnt].Head.FukusuChk == "1")
	            {   
                    sCnt --;

                    //↓全ての伝票が結合されている場合、エラーとなるのを回避する
                    if (sCnt < 0)
                    {
                        //先頭行も伝票結合となっていた場合、１伝票目を先頭とする
                        sCnt = 0;
                        break;
                    }
                    //↑全ての伝票が結合されている場合、エラーとなるのを回避する
	            }

                DenIndex = sCnt;
            }
    
            //削除される伝票データ保持
            DeleteDenNo = DenData[DenIndex].Head.DenNo;
            DeleteDen = DenIndex.ToString();
    
            //ＣＳＶファイルを削除
            CSV_Name = DenData[DenIndex].Head.CsvFile;
            utility.FileDelete(global.WorkDir + global.DIR_INCSV, CSV_Name);
                 
            //画像ファイルを削除
            ImgNAME = DenData[DenIndex].Head.image;
            utility.FileDelete(global.WorkDir + global.DIR_INCSV, ImgNAME);
             
            //一枚目を削除
            DenDataShift(DenIndex);
    
            //結合伝票を削除
            while (true)
            {
                //最後の伝票が削除済みなら抜ける
                if (DenIndex > global.pblDenNum - 1) break;
    
                //複数枚チェックがなければ抜ける
                if (DenData[DenIndex].Head.FukusuChk == "0") break;
        
                //ＣＳＶファイルを削除
                CSV_Name = DenData[DenIndex].Head.CsvFile;
                utility.FileDelete(global.WorkDir + global.DIR_INCSV, CSV_Name);
        
                //画像ファイルを削除
                ImgNAME = DenData[DenIndex].Head.image;
                utility.FileDelete(global.WorkDir + global.DIR_INCSV, ImgNAME);
        
                //伝票配列をシフト
                DenDataShift(DenIndex);
            }
    
            //伝票がすべて削除された場合
            if (global.pblDenNum <= 0)
            {
                MessageBox.Show("伝票がすべて削除されました。", "終了", MessageBoxButtons.OK, MessageBoxIcon.Information);

                //終了処理
                MainEnd(DenData, false);
            }
            else
            {
                //現伝票番号が伝票数より大きくなった場合
                if (DenIndex > global.pblDenNum - 1)
                {
                    DenIndex = global.pblDenNum - 1;
                }

                MessageBox.Show("伝票を削除しました。", "削除", MessageBoxButtons.OK, MessageBoxIcon.Information);

                //エラーチェック処理 
                Boolean x1;
                eTbl = ChkMainNew(DenData, out x1);

                //エラーなしで処理を終了するとき：true、エラー有りまたはエラーなしで処理を終了しないとき：false
                if (x1 == true)
                {
                    MainEnd(DenData, true);      //汎用データを作成して処理を終了する
                }
            }
        }

        /// 伝票配列データを1件シフトしてデータを削除する
        /// </summary>
        /// <param name="iX">削除するデータの配列添え字</param>
        private void DenDataShift(int iX)
        {
            int sCnt;

            for (sCnt = iX; sCnt < DenData.Length - 1; sCnt++)
			{
                DenData[sCnt].Head.image = DenData[sCnt + 1].Head.image;            //画像ファイル名
                DenData[sCnt].Head.CsvFile = DenData[sCnt + 1].Head.CsvFile;        //CSVファイル名  2004/6/24
                DenData[sCnt].Head.Year = DenData[sCnt + 1].Head.Year;              //年
                DenData[sCnt].Head.Month = DenData[sCnt + 1].Head.Month;            //月
                DenData[sCnt].Head.Day = DenData[sCnt + 1].Head.Day;                //日
                DenData[sCnt].Head.Kessan = DenData[sCnt + 1].Head.Kessan;          //決算処理フラグ
                DenData[sCnt].Head.FukusuChk = DenData[sCnt + 1].Head.FukusuChk;    //複数毎チェック
                DenData[sCnt].Head.DenNo = DenData[sCnt + 1].Head.DenNo;            //伝票No.
                DenData[sCnt].Head.Kari_T = DenData[sCnt + 1].Head.Kari_T;          //借方伝票計
                DenData[sCnt].Head.Kashi_T = DenData[sCnt + 1].Head.Kashi_T;        //貸方伝票計
                DenData[sCnt].Head.FukuMai = DenData[sCnt + 1].Head.FukuMai;        //複数毎数
                
                for (int i = 0; i < global.MAXGYOU; i++)
			    {
                    DenData[sCnt].Gyou[i].GyouNum = DenData[sCnt + 1].Gyou[i].GyouNum;          //行番号

                    //借方データ
                    DenData[sCnt].Gyou[i].Kari.Bumon = DenData[sCnt + 1].Gyou[i].Kari.Bumon;    //部門コード
                    DenData[sCnt].Gyou[i].Kari.Kamoku = DenData[sCnt + 1].Gyou[i].Kari.Kamoku;  //科目コード
                    DenData[sCnt].Gyou[i].Kari.Hojo = DenData[sCnt + 1].Gyou[i].Kari.Hojo;      //補助コード
                    DenData[sCnt].Gyou[i].Kari.Kin = DenData[sCnt + 1].Gyou[i].Kari.Kin;        //金額
                    DenData[sCnt].Gyou[i].Kari.TaxMas = DenData[sCnt + 1].Gyou[i].Kari.TaxMas;  //消費税計算区分
                    DenData[sCnt].Gyou[i].Kari.TaxKbn = DenData[sCnt + 1].Gyou[i].Kari.TaxKbn;  //税区分
                    
                    //貸方データ
                    DenData[sCnt].Gyou[i].Kashi.Bumon = DenData[sCnt + 1].Gyou[i].Kashi.Bumon;    //部門コード
                    DenData[sCnt].Gyou[i].Kashi.Kamoku = DenData[sCnt + 1].Gyou[i].Kashi.Kamoku;  //科目コード
                    DenData[sCnt].Gyou[i].Kashi.Hojo = DenData[sCnt + 1].Gyou[i].Kashi.Hojo;      //補助コード
                    DenData[sCnt].Gyou[i].Kashi.Kin = DenData[sCnt + 1].Gyou[i].Kashi.Kin;        //金額
                    DenData[sCnt].Gyou[i].Kashi.TaxMas = DenData[sCnt + 1].Gyou[i].Kashi.TaxMas;  //消費税計算区分
                    DenData[sCnt].Gyou[i].Kashi.TaxKbn = DenData[sCnt + 1].Gyou[i].Kashi.TaxKbn;  //税区分

                    DenData[sCnt].Gyou[i].CopyChk = DenData[sCnt + 1].Gyou[i].CopyChk;     //摘要複写チェック
                    DenData[sCnt].Gyou[i].Tekiyou = DenData[sCnt + 1].Gyou[i].Tekiyou;     //摘要
                    DenData[sCnt].Gyou[i].Torikeshi = DenData[sCnt + 1].Gyou[i].Torikeshi;   //取消チェック
			    }
			}
            
            global.pblDenNum = global.pblDenNum - 1;    //伝票枚数の減算
            //DenData.CopyTo(DenData = new Entity.InputRecord[global.pblDenNum], 0);  //配列の要素数変更
        }

        private void cmdChudan_Click(object sender, EventArgs e)
        {
            string sFileName;

            if (MessageBox.Show("現在の伝票を退避して処理を終了します。よろしいですか？","処理中断確認",MessageBoxButtons.YesNo,MessageBoxIcon.Question,MessageBoxDefaultButton.Button2) == DialogResult.No)
                return;

            //会社別の中断伝票フォルダ作成
            if (Directory.Exists(global.WorkDir + global.DIR_BREAK + string.Format("{0:000}",global.pblComNo)) == false)
                Directory.CreateDirectory(global.WorkDir + global.DIR_BREAK + string.Format("{0:000}",global.pblComNo));
                
            //ダイアログ入力データ取得
            DlgDataGet();

            //出力文字列インスタンス作成
            StringBuilder sb = new StringBuilder();

            //ＣＳＶファイルに書き出し
            for (int iX = 0; iX < global.pblDenNum; iX++)
            {

                //ヘッダー情報
                sFileName = DenData[iX].Head.CsvFile;

                //出力ファイルインスタンス作成
                StreamWriter outFile = new StreamWriter(global.WorkDir + global.DIR_BREAK + string.Format("{0:000}",global.pblComNo) + @"\" + sFileName,
                                                        false, System.Text.Encoding.GetEncoding(932));
                
                //出力文字列作成
                sb.Clear();
                sb.Append("*").Append(",");
                sb.Append(DenData[iX].Head.image).Append(",");
                sb.Append(DenData[iX].Head.Year).Append(",");
                sb.Append(DenData[iX].Head.Month).Append(",");
                sb.Append(DenData[iX].Head.Day).Append(",");
                sb.Append(DenData[iX].Head.DenNo).Append(",");
                sb.Append(DenData[iX].Head.Kessan).Append(",");
                sb.Append(DenData[iX].Head.FukusuChk);
        
                outFile.WriteLine(sb.ToString());   //一行書き出し
        
                //明細情報
                for (int i = 0; i < global.MAXGYOU; i++)
                {
                    if (utility.NumericCheck(DenData[iX].Gyou[i].GyouNum))
                    {
                        sb.Clear();
                        sb.Append(DenData[iX].Gyou[i].Torikeshi).Append(",");
                        sb.Append(DenData[iX].Gyou[i].GyouNum).Append(",");
                        sb.Append(DenData[iX].Gyou[i].Kari.Bumon).Append(",");
                        sb.Append(DenData[iX].Gyou[i].Kari.Kamoku).Append(",");
                        sb.Append(DenData[iX].Gyou[i].Kari.Hojo).Append(",");
                        sb.Append(DenData[iX].Gyou[i].Kari.Kin).Append(",");
                        sb.Append(DenData[iX].Gyou[i].Kari.TaxMas).Append(",");
                        sb.Append(DenData[iX].Gyou[i].Kari.TaxKbn).Append(",");
                        sb.Append(DenData[iX].Gyou[i].Kashi.Bumon).Append(",");
                        sb.Append(DenData[iX].Gyou[i].Kashi.Kamoku).Append(",");
                        sb.Append(DenData[iX].Gyou[i].Kashi.Hojo).Append(",");
                        sb.Append(DenData[iX].Gyou[i].Kashi.Kin).Append(",");
                        sb.Append(DenData[iX].Gyou[i].Kashi.TaxMas).Append(",");
                        sb.Append(DenData[iX].Gyou[i].Kashi.TaxKbn).Append(",");
                        sb.Append(DenData[iX].Gyou[i].CopyChk).Append(",");
                        sb.Append(DenData[iX].Gyou[i].Tekiyou);

                        outFile.WriteLine(sb.ToString());   //一行書き出し
                    }
                }
        
                outFile.Close();
        
                ////中断フォルダへ移動
                //File.Move(global.WorkDir + global.DIR_INCSV + global.DIVFILE,
                //          global.WorkDir + global.DIR_BREAK + string.Format("{0:000}",global.pblComNo) + @"\" + sFileName);
            }
                
            //画像ファイルを退避
            foreach(string files in System.IO.Directory.GetFiles(global.WorkDir + global.DIR_INCSV ,"*.bmp"))
            {
                //パスを含まないファイル名を取得
                string bFileName = files.Replace(global.WorkDir + global.DIR_INCSV, string.Empty);
 
                // ファイルを移動する
                File.Move(files, global.WorkDir + global.DIR_BREAK + string.Format("{0:000}",global.pblComNo) + @"\" + bFileName);
            }

            //分割フォルダの全てのファイルを削除する
            utility.FileDelete(global.WorkDir + global.DIR_INCSV, "*");

            //中断処理終了
            global.pblDenNum = 0;
            MessageBox.Show("ファイルの退避処理を行いました。プログラムを終了します。", Application.ProductName,MessageBoxButtons.OK,MessageBoxIcon.Information);

            //処理終了
            MainEnd(DenData, false);
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            //編集モード終了
            this.gcMultiRow1.EndEdit();

            //ダイアログ入力データ取得
            btnOk.Focus();
            DlgDataGet();

            //エラーチェック 
            Boolean x1;
            eTbl = ChkMainNew(DenData, out x1);

            //エラーなしで処理を終了するとき：true、エラー有りまたはエラーなしで処理を終了しないとき：false
            if (x1 == true)
            {
                MainEnd(DenData, true);      //汎用データを作成して処理を終了する
            }
        }

        ///-----------------------------------------------------------
        /// <summary>
        ///     MultiRow勘定科目名表示 : 2017/06/06</summary>
        /// <param name="gmr">
        ///     MultiRowoオブジェクト名</param>
        /// <param name="i">
        ///     rowIndex</param>
        /// <param name="cName">
        ///     科目名セル名</param>
        /// <param name="cCode">
        ///     科目コードセル名</param>
        ///-----------------------------------------------------------
        private void GetKamokeName(GcMultiRow gmr, int i, string cName, string cCode)
        {
            string KanjoCode = string.Empty;
            
            // 接続文字列を取得する
            string sc = SqlControl.obcConnectSting.get(global.pblDbName);

            // データベース接続
            SqlControl.DataControl sCon = new SqlControl.DataControl(sc);

            string mySql = string.Empty;
            SqlDataReader dR;                                             //データリーダー

            //勘定科目取得
            if (utility.NumericCheck(gmr.GetValue(i, cCode).ToString().Trim()))
            {
                KanjoCode = string.Format("{0:D10}", int.Parse(gmr.GetValue(i, cCode).ToString().Trim()));
            }
            else
            {
                KanjoCode = gmr.GetValue(i, cCode).ToString().Trim();
            }

            //科目名表示
            //mySql += "select sUcd,sNm from wkskm01 ";
            //mySql += "where tiIsTrk = 1 ";
            //mySql += "and sUcd = '" + 
            //            string.Format("{0,6}", gmr.GetValue(i, cCode).ToString().Trim()) + "'";

            mySql += "SELECT AccountItemCode, AccountItemName FROM tbAccountItem ";
            mySql += "WHERE (tbAccountItem.IsUse = 1) and ";
            mySql += "(tbAccountItem.AccountingPeriodID = " + global.pblAccPID + ") and ";
            mySql += "(AccountItemCode = '" + KanjoCode + "')";
            
            //データリーダーを取得する
            dR = sCon.free_dsReader(mySql);

            if (dR.HasRows)
            {
                dR.Read();
                gmr.SetValue(i, cName, dR["AccountItemName"].ToString().Trim());
                gmr[i, cName].Style.ForeColor = Color.Blue;
            }
            else
            {
                gmr.SetValue(i, cName, "存在しない勘定科目コードです");
                gmr[i, cName].Style.ForeColor = Color.Red;
            }

            dR.Close();
            sCon.Close();
        }

        /// <summary>
        /// MultiRow補助科目名表示
        /// </summary>
        /// <param name="gmr">MultiRowoオブジェクト名</param>
        /// <param name="i">rowindex</param>
        /// <param name="cName">勘定科目名セル名</param>
        /// <param name="cCode">勘定科目コードセル名</param>
        /// <param name="hName">補助科目名セル名</param>
        /// <param name="hCode">補助科目コードセル名</param>
        private void GetHojoName(GcMultiRow gmr, int i, string cName, string cCode, string hName, string hCode)
        {
            string KanjoCode = string.Empty;
            string hojoCode = string.Empty;
            Boolean hCodestatus = false;                                    //補助コードの有無ステータス
            int hCodeCount = 0;                                             //補助コードの該当有無

            // 接続文字列取得
            string sc = SqlControl.obcConnectSting.get(global.pblDbName);

            // データベース接続
            SqlControl.DataControl sCon = new SqlControl.DataControl(sc);

            string mySql = string.Empty;
            SqlDataReader dR;                                             //データリーダー

            //勘定科目取得
            if (utility.NumericCheck(gmr.GetValue(i, cCode).ToString().Trim()))
            {
                KanjoCode = string.Format("{0:D10}", int.Parse(gmr.GetValue(i, cCode).ToString().Trim()));
            }
            else
            {
                KanjoCode = gmr.GetValue(i, cCode).ToString().Trim();
            }

            //補助科目取得
            if (utility.NumericCheck(gmr.GetValue(i, hCode).ToString().Trim()))
            {
                hojoCode = string.Format("{0:D10}", int.Parse(gmr.GetValue(i, hCode).ToString().Trim()));
            }
            else
            {
                hojoCode = gmr.GetValue(i, hCode).ToString().Trim();
            }

            //補助コードがあるか？
            //mySql += "select sNcd,sUcd,sHjoUcd,wkhjm01.sNm ";
            //mySql += "from wkskm01 inner join wkhjm01 ";
            //mySql += "on wkskm01.sNcd = wkhjm01.sSknNcd ";
            //mySql += "where sHjoUcd <> '000000' and sUcd = '" + string.Format("{0,6}", gmr.GetValue(i, cCode).ToString().Trim()) + "' ";
            //mySql += "order by sSknNcd,sHjoUcd";

            //補助コードがあるか？
            mySql += "select tbAccountItem.AccountItemID,tbAccountItem.AccountItemCode,";
            mySql += "tbAccountItem.AccountItemName,tbSubAccountItem.SubAccountItemCode,";
            mySql += "tbSubAccountItem.SubAccountItemName ";
            mySql += "from tbAccountItem inner join tbSubAccountItem ";
            mySql += "on tbAccountItem.AccountItemID = tbSubAccountItem.AccountItemID ";
            mySql += "where tbAccountItem.AccountingPeriodID = " + global.pblAccPID + " and ";
            mySql += "SubAccountItemCode <> '0000000000' and ";
            mySql += "tbAccountItem.AccountItemCode = '" + KanjoCode + "'";

            //データリーダーを取得し勘定科目に補助科目が設定されているか調べる
            dR = sCon.free_dsReader(mySql);
            if (dR.HasRows) hCodestatus = true;
            dR.Close();

            //勘定科目に補助コードが登録されているとき
            if (hCodestatus == true)
            {
                if (gmr.GetValue(i, hCode).ToString().Trim() == string.Empty)
                {
                    gmr.SetValue(i, hName, "補助コード未登録です");
                    gmr[i, hName].Style.ForeColor = Color.Red;
                }
                else
                {
                    //その他を含めた補助科目のデータリーダーを取得する
                    //mySql = string.Empty;
                    //mySql += "select sNcd,sUcd,sHjoUcd,wkhjm01.sNm ";
                    //mySql += "from wkskm01 inner join wkhjm01 ";
                    //mySql += "on wkskm01.sNcd = wkhjm01.sSknNcd ";
                    //mySql += "where sUcd = '" + string.Format("{0,6}", gmr.GetValue(i, cCode).ToString().Trim()) + "' ";
                    //mySql += "order by sSknNcd,sHjoUcd";

                    //その他を含めた補助科目のデータリーダーを取得する
                    mySql = string.Empty;
                    mySql += "select tbAccountItem.AccountItemID,tbAccountItem.AccountItemCode,";
                    mySql += "tbAccountItem.AccountItemName,tbSubAccountItem.SubAccountItemCode,";
                    mySql += "tbSubAccountItem.SubAccountItemName ";
                    mySql += "from tbAccountItem inner join tbSubAccountItem ";
                    mySql += "on tbAccountItem.AccountItemID = tbSubAccountItem.AccountItemID ";
                    mySql += "where tbAccountItem.AccountingPeriodID = " + global.pblAccPID + " and ";
                    mySql += "tbAccountItem.AccountItemCode = '" + KanjoCode + "'";

                    dR = sCon.free_dsReader(mySql);

                    while (dR.Read())
                    {
                        if (dR["SubAccountItemCode"].ToString().Trim() == hojoCode)
                        {
                            gmr.SetValue(i, hName, dR["SubAccountItemName"].ToString().Trim());
                            gmr[i, hName].Style.ForeColor = Color.Blue;
                            hCodeCount = 1;
                            break;
                        }
                    }

                    if (hCodeCount == 0)
                    {
                        gmr.SetValue(i, hName, "存在しないコードです");
                        gmr[i, hName].Style.ForeColor = Color.Red;
                    }

                    dR.Close();
                }
            }
            else if (gmr.GetValue(i, hCode).ToString().Trim() == string.Empty)
            {
                gmr.SetValue(i, hName, string.Empty);
                gmr[i, hName].Style.ForeColor = Color.Blue;
            }
            else
            {
                gmr.SetValue(i, hName, "存在しないコードです");
                gmr[i, hName].Style.ForeColor = Color.Red;
            }
            
            sCon.Close();
        }

        ///-------------------------------------------------------------
        /// <summary>
        ///     MultiRow部門名表示 </summary>
        /// <param name="gmr">
        ///     MultiRowoオブジェクト名</param>
        /// <param name="i">
        ///     rowindex</param>
        /// <param name="cName">
        ///     部門名セル名</param>
        /// <param name="cCode">
        ///     部門コードセル名</param>
        ///-------------------------------------------------------------
        private void GetBumonName(GcMultiRow gmr, int i, string cName, string cCode)
        {
            // 接続文字列取得
            string sc = SqlControl.obcConnectSting.get(global.pblDbName);

            // データベース接続
            SqlControl.DataControl sCon = new SqlControl.DataControl(sc);   //コントロールインスタンス生成
            
            string mySql = string.Empty;
            SqlDataReader dR;
            string CodeB;

            //部門コード編集
            if (utility.NumericCheck(gmr.GetValue(i, cCode).ToString().Trim()))
            {
                CodeB = string.Format("{0:D15}", int.Parse(gmr.GetValue(i, cCode).ToString().Trim()));
            }
            else
            {
                CodeB = gmr.GetValue(i, cCode).ToString().Trim();
            }
                       
                       
            //勘定奉行データベースへ接続する
            mySql = string.Empty;
            //mySql += "SELECT sUcd,sNm from wkbnm01 ";
            //mySql += "where sUcd = '" + CodeB + "'";

            mySql += "select DepartmentID,DepartmentCode,DepartmentName ";
            mySql += "from tbDepartment ";
            mySql += "where tbDepartment.DepartmentCode = '" + CodeB + "'";

            //データリーダーを取得する
            dR = sCon.free_dsReader(mySql);
            if (dR.HasRows)
            {
                dR.Read();
                gmr.SetValue(i, cName, dR["DepartmentName"].ToString().Trim());
                gmr[i, cName].Style.ForeColor = Color.Blue;
            }
            else
            {
                gmr.SetValue(i, cName, "存在しないコードです");
                gmr[i, cName].Style.ForeColor = Color.Red;
            }

            dR.Close();
            sCon.Close();
        }

        private void txt_KeyPress(object sender, KeyPressEventArgs e)
        {
            //数字またはバックスペースキーのみ許可する
            if (sender == txtYear || sender == txtMonth || sender == txtDay || sender == txtDenNo)
            {
                 if ((e.KeyChar < '0' || e.KeyChar > '9') && e.KeyChar != '\b')
                 {
                     e.Handled = true;
                 }
            }
        }

        /// <summary>
        /// cellによってマスター表示タブを切り替える
        /// </summary>
        /// <param name="e">セル関連イベント</param>
        private void Cell_EnterClick(CellEventArgs e)
        {
            if (e.RowIndex < 0) return;

            //取り消し欄がチェックされている行であればぬける
            if (e.RowIndex == 0 && chkDel0.Checked == true) return;
            if (e.RowIndex == 1 && chkDel1.Checked == true) return;
            if (e.RowIndex == 2 && chkDel2.Checked == true) return;
            if (e.RowIndex == 3 && chkDel3.Checked == true) return;
            if (e.RowIndex == 4 && chkDel4.Checked == true) return;
            if (e.RowIndex == 5 && chkDel5.Checked == true) return;
            if (e.RowIndex == 6 && chkDel6.Checked == true) return;

            //科目欄クリック時
            if (e.CellName == MultiRow.DP_KARI_CODE || e.CellName == MultiRow.DP_KARI_NAME ||
                e.CellName == MultiRow.DP_KASHI_CODE || e.CellName == MultiRow.DP_KASHI_NAME)
            {
                tabData.SelectedIndex = global.TAB_KAMOKU;
                return;
            }

            //借方補助欄クリック時
            if (e.CellName == MultiRow.DP_KARI_CODEH || e.CellName == MultiRow.DP_KARI_NAMEH)
            {
                tabData.SelectedIndex = global.TAB_KAMOKU;
                this.fgHojo.RowCount = 0;

                //選択された科目の補助設定がある場合、補助リストを表示
                GridViewShow_Hojo(this.fgHojo, gcMultiRow1.GetValue(e.RowIndex, MultiRow.DP_KARI_CODE).ToString());
                return;
            }

            //貸方補助欄クリック時
            if (e.CellName == MultiRow.DP_KASHI_CODEH || e.CellName == MultiRow.DP_KASHI_NAMEH)
            {
                tabData.SelectedIndex = global.TAB_KAMOKU;
                this.fgHojo.RowCount = 0;

                //選択された科目の補助設定がある場合、補助リストを表示
                GridViewShow_Hojo(this.fgHojo, gcMultiRow1.GetValue(e.RowIndex, MultiRow.DP_KASHI_CODE).ToString());
                return;
            }

            //税処理欄クリック時
            if (e.CellName == MultiRow.DP_KARI_ZEI_S || e.CellName == MultiRow.DP_KASHI_ZEI_S)
            {
                tabData.SelectedIndex = global.TAB_TAXBUMON;
                return;
            }

            //税区分欄クリック時
            if (e.CellName == MultiRow.DP_KARI_ZEI || e.CellName == MultiRow.DP_KASHI_ZEI)
            {
                tabData.SelectedIndex = global.TAB_TAXBUMON;
                return;
            }

            //部門欄クリック時
            if (e.CellName == MultiRow.DP_KARI_CODEB || e.CellName == MultiRow.DP_KARI_NAMEB ||
                e.CellName == MultiRow.DP_KASHI_CODEB || e.CellName == MultiRow.DP_KASHI_NAMEB)
            {
                tabData.SelectedIndex = global.TAB_TAXBUMON;
                return;
            }

            //摘要欄クリック時
            if (e.CellName == MultiRow.DP_TEKIYOU)
            {
                tabData.SelectedIndex = global.TAB_TEKIYOU;
                return;
            }
        }

        private void gcMultiRow1_CellClick(object sender, CellEventArgs e)
        {
            Cell_EnterClick(e);
            gcMultiRow1.BeginEdit(true);
        }

        private void ChkErrColor_Click(object sender, EventArgs e)
        {
        }

        private void fgBumon_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            //伝票データへコード、名称をセットする
            fgDataSet(fgBumon, MultiRow.DP_KARI_CODEB, MultiRow.DP_KARI_NAMEB, MultiRow.DP_KASHI_CODEB, MultiRow.DP_KASHI_NAMEB);
        }

        /// <summary>
        /// マスター表示グリッドから伝票へデータをセットする
        /// </summary>
        /// <param name="Dgv">マスター表示用 DataGridViewオブジェクト名</param>
        /// <param name="cuCellCode_Kari">借方コードセル名</param>
        /// <param name="cuCellName_Kari">借方名称セル名</param>
        /// <param name="cuCellCode_Kashi">貸方コードセル名</param>
        /// <param name="cuCellName_Kashi">貸方名称セル名</param>
        private void fgDataSet(DataGridView Dgv, string cuCellCode_Kari, string cuCellName_Kari, string cuCellCode_Kashi, string cuCellName_Kashi)
        {
            string sKmkCode;    //コード
            string sKmkName;    //名称

            if (Dgv.Rows.Count == 0) return;

            sKmkCode = Dgv.SelectedRows[0].Cells[0].Value.ToString().Trim();
            sKmkName = Dgv.SelectedRows[0].Cells[1].Value.ToString().Trim();
            
            if ((gcMultiRow1.CurrentCellPosition.RowIndex == 0 && chkDel0.Checked == false) || 
                (gcMultiRow1.CurrentCellPosition.RowIndex == 1 && chkDel1.Checked == false) || 
                (gcMultiRow1.CurrentCellPosition.RowIndex == 2 && chkDel2.Checked == false) || 
                (gcMultiRow1.CurrentCellPosition.RowIndex == 3 && chkDel3.Checked == false) || 
                (gcMultiRow1.CurrentCellPosition.RowIndex == 4 && chkDel4.Checked == false) || 
                (gcMultiRow1.CurrentCellPosition.RowIndex == 5 && chkDel5.Checked == false) || 
                (gcMultiRow1.CurrentCellPosition.RowIndex == 6 && chkDel6.Checked == false))
            {
                //借方科目にフォーカスがある時
                if (gcMultiRow1.CurrentCellPosition.CellName == cuCellCode_Kari ||
                    gcMultiRow1.CurrentCellPosition.CellName == cuCellName_Kari)
                {
                    gcMultiRow1.SetValue(gcMultiRow1.CurrentCellPosition.RowIndex, cuCellCode_Kari, sKmkCode);
                    gcMultiRow1.SetValue(gcMultiRow1.CurrentCellPosition.RowIndex, cuCellName_Kari, sKmkName);

                    //テキストカラーを戻す
                    gcMultiRow1[gcMultiRow1.CurrentCellPosition.RowIndex, gcMultiRow1.CurrentCellPosition.CellName].Style.ForeColor = Color.Blue;
                    gcMultiRow1.Focus();
                }

                //貸方科目にフォーカスがある時
                if (gcMultiRow1.CurrentCellPosition.CellName == cuCellCode_Kashi ||
                    gcMultiRow1.CurrentCellPosition.CellName == cuCellName_Kashi)
                {
                    gcMultiRow1.SetValue(gcMultiRow1.CurrentCellPosition.RowIndex, cuCellCode_Kashi, sKmkCode);
                    gcMultiRow1.SetValue(gcMultiRow1.CurrentCellPosition.RowIndex, cuCellName_Kashi, sKmkName);

                    //テキストカラーを戻す
                    gcMultiRow1[gcMultiRow1.CurrentCellPosition.RowIndex, gcMultiRow1.CurrentCellPosition.CellName].Style.ForeColor = Color.Blue;
                    gcMultiRow1.Focus();
                }
            }
        }

        private void fgKamoku_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            //伝票データへコード、名称をセットする
            fgDataSet(fgKamoku, MultiRow.DP_KARI_CODE, MultiRow.DP_KARI_NAME, MultiRow.DP_KASHI_CODE, MultiRow.DP_KASHI_NAME);
        }

        private void fgHojo_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            //伝票データへコード、名称をセットする
            fgDataSet(fgHojo, MultiRow.DP_KARI_CODEH, MultiRow.DP_KARI_NAMEH, MultiRow.DP_KASHI_CODEH, MultiRow.DP_KASHI_NAMEH);
        }

        /// <summary>
        /// 税処理、税区分表示グリッドから伝票へデータをセットする
        /// </summary>
        /// <param name="Dgv">マスター表示用 DataGridViewオブジェクト名</param>
        /// <param name="cuCellCode_Kari">借方コードセル名</param>
        /// <param name="cuCellCode_Kashi">貸方コードセル名</param>
        private void fgTaxDataSet(DataGridView Dgv, string cuCellCode_Kari,string cuCellCode_Kashi)
        {
            string sKmkCode;    //コード

            if (Dgv.Rows.Count == 0) return;

            sKmkCode = Dgv.SelectedRows[0].Cells[0].Value.ToString().Trim();

            //if (gcMultiRow1.Rows[gcMultiRow1.CurrentCellPosition.RowIndex].Cells[MultiRow.DP_DELCHK].Value.ToString() == "No")
            
            if ((gcMultiRow1.CurrentCellPosition.RowIndex == 0 && chkDel0.Checked == false) ||
                (gcMultiRow1.CurrentCellPosition.RowIndex == 1 && chkDel1.Checked == false) ||
                (gcMultiRow1.CurrentCellPosition.RowIndex == 2 && chkDel2.Checked == false) ||
                (gcMultiRow1.CurrentCellPosition.RowIndex == 3 && chkDel3.Checked == false) ||
                (gcMultiRow1.CurrentCellPosition.RowIndex == 4 && chkDel4.Checked == false) ||
                (gcMultiRow1.CurrentCellPosition.RowIndex == 5 && chkDel5.Checked == false) ||
                (gcMultiRow1.CurrentCellPosition.RowIndex == 6 && chkDel6.Checked == false))
            {
                //借方科目にフォーカスがある時
                if (gcMultiRow1.CurrentCellPosition.CellName == cuCellCode_Kari)
                {
                    gcMultiRow1.SetValue(gcMultiRow1.CurrentCellPosition.RowIndex, cuCellCode_Kari, sKmkCode);

                    //テキストカラーを戻す
                    gcMultiRow1[gcMultiRow1.CurrentCellPosition.RowIndex, gcMultiRow1.CurrentCellPosition.CellName].Style.ForeColor = Color.Blue;
                    gcMultiRow1.Focus();
                }

                //貸方科目にフォーカスがある時
                if (gcMultiRow1.CurrentCellPosition.CellName == cuCellCode_Kashi)
                {
                    gcMultiRow1.SetValue(gcMultiRow1.CurrentCellPosition.RowIndex, cuCellCode_Kashi, sKmkCode);

                    //テキストカラーを戻す
                    gcMultiRow1[gcMultiRow1.CurrentCellPosition.RowIndex, gcMultiRow1.CurrentCellPosition.CellName].Style.ForeColor = Color.Blue;
                    gcMultiRow1.Focus();
                }
            }
        }

        /// <summary>
        /// 摘要表示グリッドから伝票へデータをセットする
        /// </summary>
        /// <param name="Dgv">マスター表示用 DataGridViewオブジェクト名</param>
        /// <param name="cuCellName">セル名</param>
        private void fgTekiyoDataSet(DataGridView Dgv, string cuCellName)
        {
            string sKmkName;    //コード

            if (Dgv.Rows.Count == 0) return;

            sKmkName = Dgv.SelectedRows[0].Cells[0].Value.ToString().Trim();

            //if (gcMultiRow1.Rows[gcMultiRow1.CurrentCellPosition.RowIndex].Cells[MultiRow.DP_DELCHK].Value.ToString() == "No")
           
            if ((gcMultiRow1.CurrentCellPosition.RowIndex == 0 && chkDel0.Checked == false) ||
               (gcMultiRow1.CurrentCellPosition.RowIndex == 1 && chkDel1.Checked == false) ||
               (gcMultiRow1.CurrentCellPosition.RowIndex == 2 && chkDel2.Checked == false) ||
               (gcMultiRow1.CurrentCellPosition.RowIndex == 3 && chkDel3.Checked == false) ||
               (gcMultiRow1.CurrentCellPosition.RowIndex == 4 && chkDel4.Checked == false) ||
               (gcMultiRow1.CurrentCellPosition.RowIndex == 5 && chkDel5.Checked == false) ||
               (gcMultiRow1.CurrentCellPosition.RowIndex == 6 && chkDel6.Checked == false))
            {
                if (gcMultiRow1.CurrentCellPosition.CellName == cuCellName)
                {
                    gcMultiRow1.SetValue(gcMultiRow1.CurrentCellPosition.RowIndex, cuCellName, sKmkName);

                    //テキストカラーを戻す
                    gcMultiRow1[gcMultiRow1.CurrentCellPosition.RowIndex, gcMultiRow1.CurrentCellPosition.CellName].Style.ForeColor = Color.Blue;
                    gcMultiRow1.Focus();
                }
            }
        }

        private void fgTax_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            //伝票データへコードをセットする
            fgTaxDataSet(fgTax, MultiRow.DP_KARI_ZEI, MultiRow.DP_KASHI_ZEI);
        }

        private void fgTaxMas_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            //伝票データへコードをセットする
            fgTaxDataSet(fgTaxMas, MultiRow.DP_KARI_ZEI_S, MultiRow.DP_KASHI_ZEI_S);
        }

        private void fgTekiyo_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            //伝票データへ摘要をセットする
            fgTekiyoDataSet(fgTekiyo, MultiRow.DP_TEKIYOU);
        }

        private void fgNg_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            //エラー箇所をフォーカス
        
            if (fgNg.RowCount == 0) return;
    
            //ページ番号取得
            int pNo = int.Parse(fgNg.SelectedRows[0].Cells[0].Value.ToString().Trim()) - 1;
    
            //行番号取得
            int GyoNo = int.Parse(fgNg.SelectedRows[0].Cells[1].Value.ToString().Trim()) - 1;
     
            //エラーセル名取得
            string CellName = fgNg.SelectedRows[0].Cells[5].Value.ToString().Trim();
   
            //ダイアログ入力データ取得
            DlgDataGet();
            DenIndex = pNo;
            DataShow(DenIndex,DenData,eTbl);    //データ表示
    
            switch (CellName)
	        {
                case "kessan":
                    ChkKessan.Focus();
                    break;

                case "fukusu":
                    chkFukusuChk.Focus();
                    break;

                case "txtYear":
                    txtYear.Focus();
                    break;

                case "txtDenNo":
                    txtDenNo.Focus();
                    break;

                case "txtSagaku_T":
                    break;

		        default:
                    gcMultiRow1.Focus();
                    gcMultiRow1.CurrentCellPosition = new CellPosition(GyoNo, CellName);
                    gcMultiRow1.BeginEdit(true);
                    break;
	        }
        }

        private void fgKamoku_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            //選択された勘定科目の補助科目を表示する
            GridViewShow_Hojo(fgHojo, fgKamoku.SelectedRows[0].Cells[0].Value.ToString());
        }

        private void fgKamoku_SelectionChanged(object sender, EventArgs e)
        {
            if (global.MASTERLOAD_STATUS == 1) return;
 
            //選択された勘定科目の補助科目を表示する
            GridViewShow_Hojo(fgHojo, fgKamoku.SelectedRows[0].Cells[0].Value.ToString());
        }

        private void cmdPlus_Click(object sender, EventArgs e)
        {
            if (leadImg.ScaleFactor < global.ZOOM_MAX)
            {
                leadImg.ScaleFactor += global.ZOOM_STEP;
            }
            global.miMdlZoomRate = (float)leadImg.ScaleFactor;
        }

        private void cmdMinus_Click(object sender, EventArgs e)
        {
            if (leadImg.ScaleFactor > global.ZOOM_MIN)
            {
                leadImg.ScaleFactor -= global.ZOOM_STEP;
            }
            global.miMdlZoomRate = (float)leadImg.ScaleFactor;
        }

        private void button1_Click(object sender, EventArgs e)
        {

            fgKamoku.CurrentCell = null;
        }

        private void ChkErrColor_CheckedChanged(object sender, EventArgs e)
        {
            ErrColorChange();
        }

        private void ErrColorChange()
        {
            Show_NGColor(DenIndex, eTbl);

            //ＮＧ項目カラー表示
            if (ChkErrColor.Checked == false)
            {
                //Alternating表示
                gcMultiRow1.AlternatingRowsDefaultCellStyle.BackColor =  Color.FromArgb(200, 249, 196);

                //決算バックカラーの取得、戻し
                global.pblKessanColor = this.ChkKessan.BackColor;
                this.ChkKessan.BackColor = System.Drawing.SystemColors.Control;

                //複数枚バックカラーの取得、戻し
                global.pblFukuColor = this.chkFukusuChk.BackColor;
                this.chkFukusuChk.BackColor = System.Drawing.SystemColors.Control;

                //差額バックカラーの取得、戻し
                global.pblSagakuColor = gcMultiRow1.ColumnFooters[0].Cells[MultiRow.DP_SAGAKU_T].Style.BackColor;
                gcMultiRow1.ColumnFooters[0].Cells[MultiRow.DP_SAGAKU_T].Style.BackColor = Color.White;
            }
            else
            {
                //Alternating表示
                gcMultiRow1.AlternatingRowsDefaultCellStyle.BackColor = Color.White;

                //決算バックカラーを戻す
                this.ChkKessan.BackColor = global.pblKessanColor;

                //複数枚バックカラーを戻す
                this.chkFukusuChk.BackColor = global.pblFukuColor;

                //差額バックカラーを戻す
                gcMultiRow1.ColumnFooters[0].Cells[MultiRow.DP_SAGAKU_T].Style.BackColor = global.pblSagakuColor;
            }
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            // グレースケールに変換
            GrayscaleCommand grayScaleCommand = new GrayscaleCommand();
            grayScaleCommand.BitsPerPixel = 8;
            grayScaleCommand.Run(leadImg.Image);
            leadImg.Refresh();

            MessageBox.Show(leadImg.Image.GrayscaleMode.ToString());

            //// ネガポジ反転します。
            //InvertCommand invertCommand = new InvertCommand();
            //invertCommand.Run(_viewer.Image);
            //leadImg.Refresh();
        }

        private void cmdImgPrn_Click(object sender, EventArgs e)
        { 
            //印刷確認
            if (global.pblImageFile == string.Empty) return;
            if (MessageBox.Show("この伝票画像を印刷してよろしいですか？", "印刷確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No) return;

            //画像印刷
            cPrint prn = new cPrint();
            prn.Image(leadImg);
        }

        private void btnPrint_Click(object sender, EventArgs e)
        {
            //確認
            if (MessageBox.Show("表示中の伝票を印刷しますか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question,MessageBoxDefaultButton.Button1) == DialogResult.No) return;

            DlgDataGet();

            cPrint Prn = new cPrint();
            Prn.Denpyo(DenData, DenIndex, global.PRINTMODEALL);
        }

        //private void TestPrint(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        //{
        //    for (float x  = 1; x < 120; x++)
        //    {
        //        for (float y = 1; y < 50; y++)
        //        {
        //            SetXY(x, y, e);
        //            e.Graphics.DrawString(x.ToString().Substring(0,1), new Font("ＭＳ ゴシック", PRINTFONTSIZE), Brushes.Black, PrnX, PrnY);
        //        }
        //    }
        //}

        private void chkFu1_CheckedChanged(object sender, EventArgs e)
        {
            ShowTekiyou1gyo(1, 1,DenData);
        }

        private void chkFu2_CheckedChanged(object sender, EventArgs e)
        {
            ShowTekiyou1gyo(2, 1, DenData);
        }

        private void chkFu3_CheckedChanged(object sender, EventArgs e)
        {
            ShowTekiyou1gyo(3, 1, DenData);
        }

        private void chkFu4_CheckedChanged(object sender, EventArgs e)
        {
            ShowTekiyou1gyo(4, 1, DenData);
        }

        private void chkFu5_CheckedChanged(object sender, EventArgs e)
        {
            ShowTekiyou1gyo(5, 1, DenData);
        }

        private void chkFu6_CheckedChanged(object sender, EventArgs e)
        {
            ShowTekiyou1gyo(6, 1, DenData);
        }

        private void chkDel0_CheckedChanged(object sender, EventArgs e)
        {
            //取消欄がチェックされているなら行を無効とする
            if (chkDel0.Checked == true) 
            {
                for (int idx = 0; idx <= MultiRow.DP_INDEX; idx++)
                {
                    gcMultiRow1.Rows[0].Cells[idx].Enabled = false;
                }

                chkFu0.Enabled = false;
            }
            else
            {
                //取消欄がチェックされていないなら行を有効とする
                for (int idx = 0; idx <= MultiRow.DP_INDEX; idx++)
                {
                    gcMultiRow1.Rows[0].Cells[idx].Enabled = true;
                }
            }
        }

        private void chkDel1_CheckedChanged(object sender, EventArgs e)
        {
            //取消欄がチェックされているなら行を無効とする
            if (chkDel1.Checked == true)
            {
                for (int idx = 0; idx <= MultiRow.DP_INDEX; idx++)
                {
                    gcMultiRow1.Rows[1].Cells[idx].Enabled = false;
                }
                chkFu1.Enabled = false;
            }
            else
            {
                //取消欄がチェックされていないなら行を有効とする
                for (int idx = 0; idx <= MultiRow.DP_INDEX; idx++)
                {
                    gcMultiRow1.Rows[1].Cells[idx].Enabled = true;
                }
                chkFu1.Enabled = true;
            }
        }

        private void chkDel2_CheckedChanged(object sender, EventArgs e)
        {
            //取消欄がチェックされているなら行を無効とする
            if (chkDel2.Checked == true)
            {
                for (int idx = 0; idx <= MultiRow.DP_INDEX; idx++)
                {
                    gcMultiRow1.Rows[2].Cells[idx].Enabled = false;
                }
                chkFu2.Enabled = false;
            }
            else
            {
                //取消欄がチェックされていないなら行を有効とする
                for (int idx = 0; idx <= MultiRow.DP_INDEX; idx++)
                {
                    gcMultiRow1.Rows[2].Cells[idx].Enabled = true;
                }
                chkFu2.Enabled = true;
            }
        }

        private void chkDel3_CheckedChanged(object sender, EventArgs e)
        {
            //取消欄がチェックされているなら行を無効とする
            if (chkDel3.Checked == true)
            {
                for (int idx = 0; idx <= MultiRow.DP_INDEX; idx++)
                {
                    gcMultiRow1.Rows[3].Cells[idx].Enabled = false;
                }
                chkFu3.Enabled = false;
            }
            else
            {
                //取消欄がチェックされていないなら行を有効とする
                for (int idx = 0; idx <= MultiRow.DP_INDEX; idx++)
                {
                    gcMultiRow1.Rows[3].Cells[idx].Enabled = true;
                }
                chkFu3.Enabled = true;
            }
        }

        private void chkDel4_CheckedChanged(object sender, EventArgs e)
        {
            //取消欄がチェックされているなら行を無効とする
            if (chkDel4.Checked == true)
            {
                for (int idx = 0; idx <= MultiRow.DP_INDEX; idx++)
                {
                    gcMultiRow1.Rows[4].Cells[idx].Enabled = false;
                }
                chkFu4.Enabled = false;
            }
            else
            {
                //取消欄がチェックされていないなら行を有効とする
                for (int idx = 0; idx <= MultiRow.DP_INDEX; idx++)
                {
                    gcMultiRow1.Rows[4].Cells[idx].Enabled = true;
                }
                chkFu4.Enabled = true;
            }
        }

        private void chkDel5_CheckedChanged(object sender, EventArgs e)
        {
            //取消欄がチェックされているなら行を無効とする
            if (chkDel5.Checked == true)
            {
                for (int idx = 0; idx <= MultiRow.DP_INDEX; idx++)
                {
                    gcMultiRow1.Rows[5].Cells[idx].Enabled = false;
                }
                chkFu5.Enabled = false;
            }
            else
            {
                //取消欄がチェックされていないなら行を有効とする
                for (int idx = 0; idx <= MultiRow.DP_INDEX; idx++)
                {
                    gcMultiRow1.Rows[5].Cells[idx].Enabled = true;
                }
                chkFu5.Enabled = true;
            }
        }

        private void chkDel6_CheckedChanged(object sender, EventArgs e)
        {
            //取消欄がチェックされているなら行を無効とする
            if (chkDel6.Checked == true)
            {
                for (int idx = 0; idx <= MultiRow.DP_INDEX; idx++)
                {
                    gcMultiRow1.Rows[6].Cells[idx].Enabled = false;
                }
                chkFu6.Enabled = false;
            }
            else
            {
                //取消欄がチェックされていないなら行を有効とする
                for (int idx = 0; idx <= MultiRow.DP_INDEX; idx++)
                {
                    gcMultiRow1.Rows[6].Cells[idx].Enabled = true;
                }
                chkFu6.Enabled = true;
            }
        }

        private void txtYear_Enter(object sender, EventArgs e)
        {
            TextBox  Obj = new TextBox();

            if (sender == txtYear) Obj = txtYear;
            if (sender == txtMonth) Obj = txtMonth;
            if (sender == txtDay) Obj = txtDay;
            if (sender == txtDenNo) Obj = txtDenNo;

            Obj.BackColor = Color.LightGray;
            Obj.SelectAll();
        }

        private void txtYear_Leave(object sender, EventArgs e)
        {
            TextBox Obj = new TextBox();

            if (sender == txtYear) Obj = txtYear;
            if (sender == txtMonth) Obj = txtMonth;
            if (sender == txtDay) Obj = txtDay;
            if (sender == txtDenNo) Obj = txtDenNo;
            
            Obj.BackColor = Color.White;

            if (utility.NumericCheck(Obj.Text))
            {
                Obj.Text = string.Format("{0:00}", int.Parse(Obj.Text));
            }
            else
            {
                Obj.Text = "00";
            }
        }

        private void Base_KeyPress(object sender, KeyPressEventArgs e)
        {
        }

        private void gcMultiRow1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
            {
                //ダイアログ入力データ取得
                btnOk.Focus();
                DlgDataGet();

                //エラーチェック 
                Boolean x1;
                eTbl = ChkMainNew(DenData, out x1);

                //エラーなしで処理を終了するとき：true、エラー有りまたはエラーなしで処理を終了しないとき：false
                if (x1 == true)
                {
                    MainEnd(DenData, true);      //汎用データを作成して処理を終了する
                }
            }
        }

        private void gcMultiRow1_CellContentClick(object sender, CellEventArgs e)
        {

        }

        private void gcMultiRow1_CellEnter(object sender, CellEventArgs e)
        {
            if (gcMultiRow1.EditMode == EditMode.EditProgrammatically)
            {
                Cell_EnterClick(e);
                gcMultiRow1.BeginEdit(true);
            }
        }

        private void gcMultiRow1_Enter(object sender, EventArgs e)
        {
        }
    }
}

