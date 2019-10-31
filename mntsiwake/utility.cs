using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.OleDb;
using System.Windows.Forms;
using System.Drawing;
using System.IO;
using System.Data.SqlClient;

namespace mntsiwake
{
    class utility
    {
        /// <summary>
        /// レジストリからインストールディレクトリを取得する
        /// </summary>
        /// <returns>ディレクトリ名</returns>
        public static string GetPath()
        {
            // 操作するレジストリ・キーの名前
            string rKeyName = @"SOFTWARE\FKDL";

            // 取得処理を行う対象となるレジストリの値の名前
            string rGetValueName = "InstDir";

            // レジストリの取得
            // レジストリ・キーのパスを指定してレジストリを開く
            Microsoft.Win32.RegistryKey rKey = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(rKeyName);

            // レジストリの値を取得
            string location = (string)rKey.GetValue(rGetValueName);

            // 開いたレジストリ・キーを閉じる
            rKey.Close();

            // コンソールに取得したレジストリの値を表示
            Console.WriteLine(location);
            return location;

        }

        // ローカルデータベース接続
        public class DBConnect
        {
            //OleDbConnection cn = new OleDbConnection();
            SqlConnection cn = new SqlConnection();     // 2017/06/05

            public SqlConnection Cn
            {
                get
                {
                    return cn;
                }
            }
            
            /// <summary>
            /// ローカルデータベースへの接続
            /// </summary>
            /// <param name="sDBPath">パス</param>
            /// <param name="sDBName">データベース名</param>
            public DBConnect(string sDBPath,string sDBName)
            {
                try
                {
                    // データベース接続文字列
                    cn.ConnectionString = "";
                    cn.ConnectionString += "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=";
                    cn.ConnectionString += sDBPath;
                    cn.ConnectionString += @"\";
                    cn.ConnectionString += sDBName;

                    //cn.ConnectionString += ";Jet OLEDB:Database Password=";
                    //cn.ConnectionString += sDBPword;

                    cn.Open();
                }

                catch (Exception e)
                {
                    throw e;
                }
            }
        }

        // 勘定奉行データベース接続
        public class SQLDBConnect
        {
            SqlConnection cn = new SqlConnection();

            public SqlConnection Cn
            {
                get
                {
                    return cn;
                }
            }

            ///----------------------------------------------------
            /// <summary>
            ///     SQLServerへ接続 2017/06/05 </summary>
            /// <param name="sConnect">
            ///     接続文字列</param>
            ///----------------------------------------------------
            public SQLDBConnect(string sConnect)
            {
                try
                {
                    // データベース接続文字列
                    cn.ConnectionString = sConnect;
                    cn.Open();
                }

                catch (Exception e)
                {
                    throw e;
                }
            }
        }

        /// <summary>
        /// 任意のディレクトリのファイルを削除する
        /// </summary>
        /// <param name="sPath">指定するディレクトリ</param>
        /// <param name="sFileType">ファイル名及び形式</param>
        public static void FileDelete(string sPath,string sFileType)
        {
            //sFileTypeワイルドカード"*"は、すべてのファイルを意味する
            foreach(string files in System.IO.Directory.GetFiles(sPath,sFileType))
            {
                // ファイルを削除する
                System.IO.File.Delete(files);
            }
        }

        /// <summary>
        /// 文字列の値が数字かチェックする
        /// </summary>
        /// <param name="tempStr">検証する文字列</param>
        /// <returns>数字:true,数字でない:false</returns>
        public static bool NumericCheck(string tempStr)
        {
            double d;

            if (tempStr == null) return false;

            if (double.TryParse(tempStr, System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out d) == false)
                return false;

            return true;
        }

        /// <summary>
        /// Nullをstring.Empty("")に置き換える
        /// </summary>
        /// <param name="tempStr">stringオブジェクト</param>
        /// <returns>nullのときstring.Empty、not nullのときそのまま値を返す</returns>
        public static string NulltoStr(string tempStr)
        {
            if (tempStr == null)
            {
                return string.Empty;
            }
            else
            {
                return tempStr;
            }
        }

        /// <summary>
        /// ウィンドウ最小サイズの設定
        /// </summary>
        /// <param name="tempFrm">対象とするウィンドウオブジェクト</param>
        /// <param name="wSize">width</param>
        /// <param name="hSize">Height</param>
        public static void WindowsMinSize(Form tempFrm, int wSize, int hSize)
        {
            tempFrm.MinimumSize = new Size(wSize, hSize);
        }

        /// <summary>
        /// ウィンドウ最小サイズの設定
        /// </summary>
        /// <param name="tempFrm">対象とするウィンドウオブジェクト</param>
        /// <param name="wSize">width</param>
        /// <param name="hSize">height</param>
        public static void WindowsMaxSize(Form tempFrm, int wSize, int hSize)
        {
            tempFrm.MaximumSize = new Size(wSize, hSize);
        }

        /// <summary>
        /// DSNファイルを開き接続文字列を作成する
        /// </summary>
        /// <param name="sDsnPath">DSNファイルパス名</param>
        /// <returns>接続文字列</returns>
        public static string GetConnect(String sDsnPath)
        {
            //DSNファイルを開く
            // StreamReader の新しいインスタンスを生成する
            StreamReader cReader = (new StreamReader(sDsnPath,Encoding.Default));

            // 読み込んだ結果をすべて格納するための変数を宣言する
            string stResult = string.Empty;

            // 1行とばす
            string stBuffer = cReader.ReadLine();

            // 読み込みできる文字がなくなるまで繰り返す
            while (cReader.Peek() >= 0) 
            {
                // ファイルを 1 行ずつ読み込む
                stBuffer = cReader.ReadLine();
                // 読み込んだものを追加で格納する
                stResult += stBuffer + ";";
            }

            // cReader を閉じる
            cReader.Close();

            stResult = "Provider=SQLOLEDB;" + stResult;

            //パスワードが設定されている場合のみ、パスワードを追加
            if (global.pblDsnPassWord.Trim() != string.Empty)
            {
                stResult += "PWD=" + global.pblDsnPassWord.Trim() + ";";
            }

            return stResult; 

        }

        /// <summary>
        ///  DSNファイルを開き接続文字列を作成する（データベース指定）
        /// </summary>
        /// <param name="sDBName">接続するデータベース名</param>
        /// <returns>接続文字列</returns>
        public static string GetDBConnect(string sDBName)
        {
            return ConvDsn(global.pblDsnPath) + "DATABASE=" + sDBName + ";";
        }

        /// <summary>
        /// DSNファイルを開き接続文字列を作成する
        /// </summary>
        /// <param name="sDsnPath">DSNファイルパス名</param>
        /// <returns>接続文字列</returns>
        public static string ConvDsn(String sDsnPath)
        {
            //DSNファイルを開く
            // StreamReader の新しいインスタンスを生成する
            StreamReader cReader = (new StreamReader(sDsnPath, Encoding.Default));

            // 読み込んだ結果をすべて格納するための変数を宣言する
            string stResult = string.Empty;

            // 1行とばす
            string stBuffer = cReader.ReadLine();

            // 読み込みできる文字がなくなるまで繰り返す
            while (cReader.Peek() >= 0)
            {
                // ファイルを 1 行ずつ読み込む
                stBuffer = cReader.ReadLine();
                // "DATABESEの指定以外の文字列を追加で格納する
                if (stBuffer.Contains("DATABASE") == false) stResult += stBuffer + ";";
            }

            // cReader を閉じる
            cReader.Close();

            stResult = "Provider=SQLOLEDB;" + stResult;

            //パスワードが設定されている場合のみ、パスワードを追加
            if (global.pblDsnPassWord.Trim() != string.Empty)
            {
                stResult += "PWD=" + global.pblDsnPassWord.Trim() + ";";
            }

            return stResult;

        }

    }
}
