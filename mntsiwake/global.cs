using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;

namespace mntsiwake
{
    static class global
    {
        //フォルダ・ファイル
        public static string WorkDir;           //インストールディレクトリ
        public static string DIR_OK = @"Ok\";
        public static string DIR_INCSV = @"分割\";
        public static string DIR_BREAK = @"中断\";
        public static string DIR_CONFIG = @"henkan\";
        public static string DIR_HENKAN = @"henkan\";
        public static string INFILE = "kanjo2kocr.csv";             //入力ファイル
        public static string OUTFILE = "ＯＣＲ勘定奉行変換.csv";     //出力ファイル
        public static string DIVFILE = "div.csv";                   //分割ファイル
        public static string TMPREAD = "kanjo2ktmpread.dat";        //入力ファイルのコピー
        public static string tmpFile = "kanjo2ktmpfile.dat";        //出力ファイルのコピー
        public static string LOGFILE = "kanjo2kerrlog.csv";         //エラーログファイル

        //画像表示
        public static int pblImageHeight;           //画像フォームの高さ
        public static int pblImageWidth;            //画像フォームの幅
        public static int pblImageX;                //画像の表示倍率（％）
        public static string pblImageFile;          //画像ファイル名

        //画像ファイル名
        public static string msMdlImgPath = string.Empty;          //表示中の画像のフルパス

        //表示関係
        public static float miMdlZoomRate = 0;      //現在の表示倍率
        
        //表示倍率（%）
        public static float ZOOM_RATE = 0.19f;      //標準倍率
        public static float ZOOM_MAX  = 2.00f;      //最大倍率
        public static float ZOOM_MIN  = 0.05f;      //最小倍率
        public static float ZOOM_STEP = 0.02f;      //ステップ倍率
        
        //データベース
        public static string pblComNo;              //会社№
        public static string pblDbName;             //選択された会社のデータベース名
        public static string pblBfDbName;           //前回選択したデータベース名
        public static string pblDsnPath;            //データソースのパス
        public static string pblDsnFlg;             //LAN使用のフラグ
        public static string pblDsnPassWord;        //パスワード
        public static string pblKanjoVer;
        public static string CONFIGFILE = "Kanjo2kconfig.mdb";
        public static string gsTaxMas;              //税処理
        public static string pblAccPID;             //AccountPeriodID

        //固定部門、勘定科目,補助科目
        public static string pblHeadBumon;          //固定部門
        public static string pblHeadKamoku;         //勘定科目
        public static string pblHeadHojo;           //補助科目

        //画面デザイン：配色
        public static int BACK_COLOR = 16777215;
        public static int NON_COLOR = -2147483633;

        public static Color pblBackColor;
        public static Color pblForeColor;
        public static Color pblNonColor;
        public static Color pblErrBackColor;
        public static Color pblErrForeColor;
        public static Color pblKinBackColor;
        public static Color pblKessanColor;
        public static Color pblFukuColor;
        public static Color pblSagakuColor;

        //その他
        public static int pblCombineMax;        //伝票結合最大枚数
        public static int pblMaisu;             //伝票結合枚数(枚数チェック用）
        public static int pblItem = 0;          //伝票結合行数(チェック用）
        public static int pblSelFILE;
        public static string VER_21 = "1";
        public static int MAXDEN = 36;
        public static int MAXGYOU = 7;          //最大行数
        public static int MAX2000 = 35;         //2000バージョン 最大処理行数
        public static int MAX21   = 250;        //21バージョン 最大処理行数
        public static string FLGON_2 = "2";
        public static string FLGON   = "1";
        public static string FLGOFF = "0";
        public static int pblDenNum;            //データ数
        public static Boolean pblBumonFlg;      //部門フラグの有無
        
        public static string TEKIYO_SPACE_ZEN  = "　";      //全角スペース
        public static string TEKIYO_SPACE_HAN = " ";        //半角スペース
        
        //結合伝票の合計金額
        public static decimal pblKari_T;        //借方合計
        public static decimal pblKashi_T;       //貸方合計
        public static decimal pblFukumai;       //伝票結合枚数

        //相手科目ステータス
        public static Boolean pblFlgKariKamoku;     //借方合計
        public static Boolean pblFlgKashiKamoku;    //貸方合計

        //タブ表示
        public static int TAB_ERR       = 0;    //エラー表示
        public static int TAB_KAMOKU    = 1;    //勘定科目
        public static int TAB_TAXBUMON  = 2;    //税処理、税区分・部門コード
        public static int TAB_TEKIYOU   = 3;    //摘要
        public static int TAB_COM       = 4;    //会社情報

        //印刷関係
        public static int  PRINTMODEONE = 0;    
        public static int  PRINTMODEALL = 1;
        
        public static string pblReki;           //西暦、和暦区分取得値
        public static string RWAREKI = "1";     //和暦の区分"1"をあらわす
        public static string RSEIREKI = "0";    //西暦の区分"0"をあらわす

        public static int LEN_KAMOKU = 4;       //勘定科目
        public static int LEN_BUMON = 4;        //部門
        public static int LEN_HOJO = 4;         //補助科目
        public static int LEN_PROJECT = 11;     //プロジェクトコード
        public static int LEN_SUBPROJECT = 2;   //サブプロジェクトコード

        //マスターロードステータス
        public static int MASTERLOAD_STATUS = 0;    
    }
}
