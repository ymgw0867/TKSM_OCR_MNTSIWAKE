using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.OleDb;
using System.Data.SqlClient;

namespace mntsiwake
{
    class Control
    {
        /// <summary>
        /// DataControlクラスの基本クラス
        /// </summary>
        public class BaseControl
        {
            private utility.DBConnect dbConnect;

            /// <summary>
            /// BaseControlのコンストラクタ。DBConnectクラスのインスタンスを作成します。
            /// </summary>
            public BaseControl(string sDBPath, string sDBName)
            {
                dbConnect = new utility.DBConnect(sDBPath, sDBName);
            }

            ///-----------------------------------------------
            /// <summary>
            ///     データベース接続メソッド : 2017/06/06</summary>
            /// <returns>
            ///     データベース接続情報を取得します</returns>
            ///-----------------------------------------------
            public SqlConnection GetConnection()
            {
                return dbConnect.Cn;
            }

        }

        /// <summary>
        /// データコントロールクラス BaseControlを継承する
        /// </summary>
        public class DataControl : BaseControl
        {

            private Access.DataAccess dAccess;
            //public OleDbConnection Cn = new OleDbConnection();
            public SqlConnection Cn = new SqlConnection();

            /// <summary>
            /// DataControlクラスのコンストラクタ。データアクセスクラスのインスタンスを作成します。
            /// </summary>
            public DataControl(string sDBPath,string sDBName) : base(sDBPath,sDBName)
            {
                // データアクセスクラスのインスタンスを作成する
                dAccess = new Access.DataAccess();
            }

            /// <summary>
            /// データベースの接続を解除します
            /// </summary>
            public void Close()
            {
                if (Cn.State == System.Data.ConnectionState.Open)
                {
                    Cn.Close();
                }
            }

            /// <summary>
            /// 条件付きデータリーダー取得インターフェイスを引数としたメソッド
            /// </summary>
            /// <param name="IDSR">データリーダーを取得するインターフェイス</param>
            /// <param name="tempString">SQL文のwhere以下の条件を記述した文字列</param>
            /// <returns>条件式に一致する引数で指定されたマスターのデータリーダー</returns>
            //public OleDbDataReader FillByAccess(Access.DataAccess.IFillBy IDSR, string tempString)
            //{
            //    // データベース接続情報を取得する
            //    Cn = this.GetConnection();

            //    return IDSR.GetdsReader(Cn, tempString);

            //}

            public SqlDataReader FillByAccess(Access.DataAccess.IFillBy IDSR, string tempString)
            {
                // データベース接続情報を取得する
                Cn = this.GetConnection();

                return IDSR.GetdsReader(Cn, tempString);

            }

            ///------------------------------------------------------------------
            /// <summary>
            ///     条件付きデータリーダを取得します : 2017/06/06 </summary>
            /// <param name="tempString">
            ///     SQL文を記述した文字列</param>
            /// <returns>
            ///     データリーダー</returns>
            ///------------------------------------------------------------------
            public SqlDataReader free_dsReader(string tempString)
            {

                try
                {
                    return FillByAccess(new Access.DataAccess.free_dsReader(), tempString);
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }

            public SqlDataReader free_dsReaderOLE(string tempString)
            {

                try
                {
                    return FillByAccess(new Access.DataAccess.free_dsReader(), tempString);
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }
        }

        /// <summary>
        /// フリーＳＱＬ実行
        /// </summary>
        public class FreeSql : DataControl
        {
            //private static OleDbCommand SCom = new OleDbCommand();
            private static SqlCommand SCom = new SqlCommand();

            public FreeSql(string sDBPath, string sDBName) : base(sDBPath, sDBName)
            {
                // データアクセスクラスのインスタンスを作成する
                //dAccess = new Access.DataAccess();
            }

            public Boolean Execute(string tempSql)
            {

                try
                {
                    //データベース接続情報の取得
                    Cn = this.GetConnection();

                    SCom.CommandText = tempSql;
                    SCom.Connection = Cn;

                    // SQLの実行
                    SCom.ExecuteNonQuery();
                    return true;
                }

                catch
                {
                    return false;
                }

            }
        }
    }
}
