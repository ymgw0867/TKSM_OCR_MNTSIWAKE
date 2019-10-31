using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.OleDb;    // 2017/06/5
using System.Data.SqlClient;

namespace mntsiwake
{
    class SqlControl
    {
        /// <summary>
        /// DataControlクラスの基本クラス
        /// </summary>
        public class BaseControl
        {
            private utility.SQLDBConnect dbConnect;

            ///--------------------------------------------------------
            /// <summary>
            ///     BaseControlのコンストラクタ。
            ///     DBConnectクラスのインスタンスを作成します。</summary>
            ///--------------------------------------------------------
            public BaseControl(string sConnect)
            {
                dbConnect = new utility.SQLDBConnect(sConnect);
            }

            ///-------------------------------------------------------
            /// <summary>
            ///     データベース接続メソッド </summary>
            /// <returns>
            ///     データベース接続情報を取得します</returns>
            ///-------------------------------------------------------
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

            ///---------------------------------------------------------------
            /// <summary>
            ///     DataControlクラスのコンストラクタ。
            ///     データアクセスクラスのインスタンスを作成します。</summary>
            ///---------------------------------------------------------------
            public DataControl(string sConnect)
                : base(sConnect)
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
            ///     条件付きデータリーダー取得インターフェイスを引数としたメソッド </summary>
            /// <param name="IDSR">
            ///     データリーダーを取得するインターフェイス</param>
            /// <param name="tempString">
            ///     SQL文のwhere以下の条件を記述した文字列</param>
            /// <returns>
            ///     条件式に一致する引数で指定されたマスターのデータリーダー</returns>
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

            ///----------------------------------------------------------
            /// <summary>
            ///     条件付きデータリーダを取得します </summary>
            /// <param name="tempString">
            ///     SQL文を記述した文字列</param>
            /// <returns>
            ///     データリーダー</returns>
            ///----------------------------------------------------------
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
        }

        ///---------------------------------------------------------------
        /// <summary>
        ///     勘定奉行データベースへの接続文字列を取得する </summary>
        ///---------------------------------------------------------------
        public class obcConnectSting
        {
            ///---------------------------------------------------------------------
            /// <summary>
            ///     勘定奉行データベースへの接続文字列を取得する </summary>
            /// <param name="dbName">
            ///     接続データベース名</param>
            /// <returns>
            ///     接続文字列</returns>
            ///---------------------------------------------------------------------
            public static string get(string dbName)
            {
                SqlConnectionStringBuilder cb = new SqlConnectionStringBuilder();
                cb.DataSource = Properties.Settings.Default.sqlServerName;
                cb.InitialCatalog = dbName;
                cb.IntegratedSecurity = false;
                cb.UserID = Properties.Settings.Default.sqlUID;
                cb.Password = Properties.Settings.Default.sqlPWD;

                return cb.ToString();
            }
        }

    }
}
