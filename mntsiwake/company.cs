using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Windows.Forms;

namespace mntsiwake
{
    class company
    {
        //会社コードデータ
        public static string Name;          //会社名
        public static DateTime fromDate;    //会計期間期首年月日
        public static string FromYear;      //会計期間期首年
        public static string FromMonth;     //会計期間期首月
        public static string FromDay;       //会計期間期首日

        public static DateTime ToDate;      //会計期間期末年月日
        public static string ToYear;        //会計期間期末年
        public static string ToMonth;       //会計期間期末月
        public static string ToDay;         //会計期間期末日

        public static string Kaisi;         //入力開始月
        public static string TaxMas;        //消費税計算区分
        public static string Gengou;        //元号
        public static string Hosei;         //年号補正値
        public static string Middle;        //中間期決算フラグ
        public static string Reki;          //西暦年または元号
        public static string gsVersion;     //バージョン情報
        public static string Arrange;       //整理仕訳区分    2011/06/07 

        //日付の入力範囲データ（マスター内の指定期間）
        public static string LmFromYear;    //入力開始年
        public static string LmFromMonth;   //入力開始月
        public static string LmFromDay;     //入力開始日
        public static string LmStSoeji;     //入力期間開始添え字
        public static string LmToYear;      //入力期限年
        public static string LmToMonth;     //入力期限月
        public static string LmToDay;       //入力期限日
        public static string LmEdSoeji;     //入力期間終了添え字
        public static string LmLock;        //制限の種類
        public static Boolean LmFlag;       //入力可能フラグ

        /// <summary>
        /// 勘定奉行データベースより会社情報を取得する
        /// </summary>
        public void CompDataLoad()
        {
            //勘定奉行データベース接続文字列を取得する
            string sc = SqlControl.obcConnectSting.get(global.pblDbName);

            //勘定奉行データベースへ接続する
            SqlControl.DataControl dCon = new SqlControl.DataControl(sc);

            //会社データ取得
            //データリーダーを取得する
            SqlDataReader dR;

            string sqlSTRING = "SELECT sDateKisyu,sDateKimatu,tiKaisi,sCorpNm,sGngo,sHosei,tiIsMiddle,tiIsVersion FROM wdhead";

            dR = dCon.free_dsReader(sqlSTRING);

            try
            {
                while (dR.Read())
                {
                    company.Name = dR["sCorpNm"].ToString().Trim();
                    company.FromYear = dR["sDateKisyu"].ToString().Trim().Substring(0, 4);
                    company.FromMonth = dR["sDateKisyu"].ToString().Trim().Substring(4, 2);
                    company.FromDay = dR["sDateKisyu"].ToString().Trim().Substring(6, 2);
                    company.ToYear = dR["sDateKimatu"].ToString().Trim().Substring(0, 4);
                    company.ToMonth = dR["sDateKimatu"].ToString().Trim().Substring(4, 2);
                    company.ToDay = dR["sDateKimatu"].ToString().Trim().Substring(6, 2);
                    company.Kaisi = dR["tiKaisi"].ToString().Trim();
                    company.Gengou = dR["sGngo"].ToString().Trim();
                    company.Hosei = dR["sHosei"].ToString().Trim();
                    company.Middle = dR["tiIsMiddle"].ToString().Trim();
                    company.gsVersion = dR["tiIsVersion"].ToString().Trim();
                }

                //西暦のとき
                if (Hosei == "0")
                {
                    Reki = "20";
                }
                else
                {
                    //和暦のとき
                    Reki = Gengou;
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "エラー", MessageBoxButtons.OK);
            }
            finally
            {
                dR.Close();
                dCon.Close();
            }
        }

        /// <summary>
        /// 勘定奉行データベースより伝票入力指定期間情報を取得する
        /// </summary>
        public void LimitDataLoad()
        {        
            //勘定奉行データベース接続文字列を取得する
            string sc = SqlControl.obcConnectSting.get(global.pblDbName);

            //勘定奉行データベースへ接続する
            SqlControl.DataControl dCon = new SqlControl.DataControl(sc);

            //伝票入力指定期間を取得
            //データリーダーを取得する
            SqlDataReader dR;

            string sqlSTRING = "SELECT tiStSoeji,sDnStDate,tiEdSoeji,sDnEdDate,tiIsLock FROM wjdnpyo2";

            dR = dCon.free_dsReader(sqlSTRING);

            try 
            {
                while (dR.Read())
                {
                    company.LmStSoeji = dR["tiStSoeji"].ToString().Trim();
                    company.LmFromYear = dR["sDnStDate"].ToString().Trim().Substring(0, 4);
                    company.LmFromMonth = dR["sDnStDate"].ToString().Trim().Substring(4, 2);
                    company.LmFromDay = dR["sDnStDate"].ToString().Trim().Substring(6, 2);

                    company.LmEdSoeji = dR["tiEdSoeji"].ToString().Trim();
                    company.LmToYear = dR["sDnEdDate"].ToString().Trim().Substring(0, 4);
                    company.LmToMonth = dR["sDnEdDate"].ToString().Trim().Substring(4, 2);
                    company.LmToDay = dR["sDnEdDate"].ToString().Trim().Substring(6, 2);

                    company.LmLock = dR["tiIsLock"].ToString().Trim();
                }
	        }
	        catch (Exception e)
	        {
                MessageBox.Show(e.Message, "エラー", MessageBoxButtons.OK);
	        }
            finally     
            {
                dR.Close();
                dCon.Close();
            }
        }
    }

}
