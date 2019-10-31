using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.OleDb;
using System.Data.SqlClient;

namespace mntsiwake
{
    class Access
    {
        public class DataAccess
        {
            public DataAccess()
            {
            }

            // 条件付きデータリーダー取得インターフェイス
            public interface IFillBy
            {
                // 抽象メソッド：2017/06/05 OleDbConnection → SqlConnection
                SqlDataReader GetdsReader(SqlConnection tempConnection, string tempString);
            }

            // データリーダー取得クラス
            public class free_dsReader : IFillBy
            {
                //private OleDbCommand SCom = new OleDbCommand();
                private SqlCommand SCom = new SqlCommand();
                private String mySql;
                //private OleDbDataReader dR;
                private SqlDataReader dR;

                ///--------------------------------------------------------------
                /// <summary>
                ///     データリーダー取得 : 2017/06/05</summary>
                /// <param name="tempConnection">
                ///     データベース接続情報</param>
                /// <param name="tempString">
                ///     SQL文</param>
                /// <returns>
                ///     データリーダー</returns>
                ///--------------------------------------------------------------
                public SqlDataReader GetdsReader(SqlConnection tempConnection, string tempString)
                {
                    //throw new Exception("The method or operation is not implemented.");

                    mySql = "";
                    mySql += tempString;
                    SCom.CommandText = mySql;
                    SCom.Connection = tempConnection;
                    dR = SCom.ExecuteReader();
                    return dR;
                }
            }

            // データリーダー取得クラス : 2017/06/05
            public class free_dsReaderOLE
            {
                private OleDbCommand SCom = new OleDbCommand();
                private String mySql;
                private OleDbDataReader dR;

                ///------------------------------------------------------
                /// <summary>
                ///     データリーダー取得 </summary>
                /// <param name="tempConnection">
                ///     データベース接続情報</param>
                /// <param name="tempString">
                ///     SQL文</param>
                /// <returns>
                ///     データリーダー</returns>
                ///------------------------------------------------------
                public OleDbDataReader GetdsReaderOLE(OleDbConnection tempConnection, string tempString)
                {
                    //throw new Exception("The method or operation is not implemented.");

                    mySql = "";
                    mySql += tempString;
                    SCom.CommandText = mySql;
                    SCom.Connection = tempConnection;
                    dR = SCom.ExecuteReader();
                    return dR;
                }
            }
        }
    }
}
