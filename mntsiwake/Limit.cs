using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace mntsiwake
{
    class Limit
    {
        /// <summary>
        /// 日付入力期間（実際の処理に使用する値）
        /// </summary>
        public class LimitKikan
        {
            public static string FromYear;    //入力開始年
            public static string FromMonth;   //入力開始月
            public static string FromDay;     //入力開始日
            public static string StSoeji;     //入力期間開始添え字
            public static string ToYear;      //入力期限年
            public static string ToMonth;     //入力期限月
            public static string ToDay;       //入力期限日
            public static string EdSoeji;     //入力期間終了添え字
            public static string Lock;        //制限の種類
            public static Boolean Flag;       //入力可能フラグ

            /// <summary>
            /// 通常仕訳の入力期間：とりあえずマスターの指定期間を入れておく
            /// </summary>
            public void initialDataSet()
            {
                FromYear = company.LmFromYear;
                FromMonth = company.LmFromMonth;
                FromDay = company.LmFromDay;
                StSoeji = company.LmStSoeji;
                ToYear = company.LmToYear;
                ToMonth = company.LmToMonth;
                ToDay = company.LmToDay;
                EdSoeji = company.LmEdSoeji;
                Lock = company.LmLock;
                Flag = company.LmFlag;
            }
        }

        /// <summary>
        /// 四半期決算期間1
        /// </summary>
        public class BfQuaKessanDate1
        {
            public static string FromYear;                  //入力開始年
            public static string FromMonth;                 //入力開始月
            public static string FromDay;                   //入力開始日
            public static string StSoeji = string.Empty;    //入力期間開始添え字
            public static string ToYear;                    //入力期限年
            public static string ToMonth;                   //入力期限月
            public static string ToDay;                     //入力期限日
            public static string EdSoeji = string.Empty;    //入力期間終了添え字
            public static string Lock = string.Empty;       //制限の種類
            public static Boolean Flag = false;             //入力可能フラグ

            private string LimitDate;
            private string FromLimitDate;
            private string ToLimitDate;

            /// <summary>
            /// 最初の四半期決算期間を取得する
            /// </summary>
            public void GetKessanDate()
            {
                //開始日    
                LimitDate = company.FromYear + "/" + company.FromMonth + "/" + company.FromDay;
    
                //期首日に2ヶ月足す
                FromLimitDate = (DateTime.Parse(LimitDate).AddMonths(2)).ToString("yyyyMMdd");

                //最初の四半期決算開始日を代入
                FromYear = FromLimitDate.Substring(0,4);
                FromMonth = FromLimitDate.Substring(4,2);
                FromDay = FromLimitDate.Substring(6,2);
    
                //終了日
                LimitDate = company.ToYear + "/" + company.ToMonth + "/" + company.ToDay;
        
                //期末日から9ヶ月引く
                ToLimitDate = (DateTime.Parse(LimitDate).AddMonths(-9)).ToString("yyyyMMdd");
    
                //最初の四半期決算終了日を代入
                ToYear = ToLimitDate.Substring(0,4);
                ToMonth = ToLimitDate.Substring(4,2);
                ToDay = ToLimitDate.Substring(6, 2);

            }
        }

        /// <summary>
        /// 四半期決算期間2
        /// </summary>
        public class BfQuaKessanDate2
        {
            public static string FromYear;                  //入力開始年
            public static string FromMonth;                 //入力開始月
            public static string FromDay;                   //入力開始日
            public static string StSoeji = string.Empty;    //入力期間開始添え字
            public static string ToYear;                    //入力期限年
            public static string ToMonth = string.Empty;    //入力期限月
            public static string ToDay;                     //入力期限日
            public static string EdSoeji = string.Empty;    //入力期間終了添え字
            public static string Lock = string.Empty;       //制限の種類
            public static Boolean Flag = false;             //入力可能フラグ

            private string LimitDate;
            private string FromLimitDate;
            private string ToLimitDate;

            /// <summary>
            /// 2度目の四半期決算期間を取得する
            /// </summary>
            public void GetKessanDate()
            {
                //開始日    
                LimitDate = company.FromYear + "/" + company.FromMonth + "/" + company.FromDay;

                //期首日に5ヶ月足す
                FromLimitDate = (DateTime.Parse(LimitDate).AddMonths(5)).ToString("yyyyMMdd");

                //四半期決算開始日を代入
                FromYear = FromLimitDate.Substring(0, 4);
                FromMonth = FromLimitDate.Substring(4, 2);
                FromDay = FromLimitDate.Substring(6, 2);

                //終了日
                LimitDate = company.ToYear + "/" + company.ToMonth + "/" + company.ToDay;

                //期末日から6ヶ月引く
                ToLimitDate = (DateTime.Parse(LimitDate).AddMonths(-6)).ToString("yyyyMMdd");

                //四半期決算終了日を代入
                ToYear = ToLimitDate.Substring(0, 4);
                ToMonth = ToLimitDate.Substring(4, 2);
                ToDay = ToLimitDate.Substring(6, 2);

            }
        }

        /// <summary>
        /// 四半期決算期間3
        /// </summary>
        public class BfQuaKessanDate3
        {
            public static string FromYear;                  //入力開始年
            public static string FromMonth;                 //入力開始月
            public static string FromDay;                   //入力開始日
            public static string StSoeji = string.Empty;    //入力期間開始添え字
            public static string ToYear;                    //入力期限年
            public static string ToMonth;                   //入力期限月
            public static string ToDay;                     //入力期限日
            public static string EdSoeji = string.Empty;    //入力期間終了添え字
            public static string Lock = string.Empty;       //制限の種類
            public static Boolean Flag = false;             //入力可能フラグ

            private string LimitDate;
            private string FromLimitDate;
            private string ToLimitDate;

            /// <summary>
            /// 3度目の四半期決算期間を取得する
            /// </summary>
            public void GetKessanDate()
            {
                //開始日    
                LimitDate = company.FromYear + "/" + company.FromMonth + "/" + company.FromDay;

                //期首日に8ヶ月足す
                FromLimitDate = (DateTime.Parse(LimitDate).AddMonths(8)).ToString("yyyyMMdd");

                //四半期決算開始日を代入
                FromYear = FromLimitDate.Substring(0, 4);
                FromMonth = FromLimitDate.Substring(4, 2);
                FromDay = FromLimitDate.Substring(6, 2);

                //終了日
                LimitDate = company.ToYear + "/" + company.ToMonth + "/" + company.ToDay;

                //期末日から3ヶ月引く
                ToLimitDate = (DateTime.Parse(LimitDate).AddMonths(-3)).ToString("yyyyMMdd");

                //四半期決算終了日を代入
                ToYear = ToLimitDate.Substring(0, 4);
                ToMonth = ToLimitDate.Substring(4, 2);
                ToDay = ToLimitDate.Substring(6, 2);

            }
        }

        /// <summary>
        /// 中間期決算期間
        /// </summary>
        public class MidKessanDate
        {
            public static string FromYear;                  //入力開始年
            public static string FromMonth;                 //入力開始月
            public static string FromDay;                   //入力開始日
            public static string StSoeji = string.Empty;    //入力期間開始添え字
            public static string ToYear;                    //入力期限年
            public static string ToMonth;                   //入力期限月
            public static string ToDay;                     //入力期限日
            public static string EdSoeji = string.Empty;    //入力期間終了添え字
            public static string Lock = string.Empty;       //制限の種類
            public static Boolean Flag;                     //入力可能フラグ

            private string LimitDate;
            private string FromLimitDate;
            private string ToLimitDate;

            /// <summary>
            /// 中間期決算期間を取得する
            /// </summary>
            public void GetKessanDate()
            {
                //開始日    
                LimitDate = company.FromYear + "/" + company.FromMonth + "/" + company.FromDay;

                //期首日に5ヶ月足す
                FromLimitDate = (DateTime.Parse(LimitDate).AddMonths(5)).ToString("yyyyMMdd");

                //中間期決算開始日を代入
                FromYear = FromLimitDate.Substring(0, 4);
                FromMonth = FromLimitDate.Substring(4, 2);
                FromDay = FromLimitDate.Substring(6, 2);

                //終了日
                LimitDate = company.ToYear + "/" + company.ToMonth + "/" + company.ToDay;

                //期末日から6ヶ月引く
                ToLimitDate = (DateTime.Parse(LimitDate).AddMonths(-6)).ToString("yyyyMMdd");

                //中間期決算終了日を代入
                ToYear = ToLimitDate.Substring(0, 4);
                ToMonth = ToLimitDate.Substring(4, 2);
                ToDay = ToLimitDate.Substring(6, 2);

            }
        }

        /// <summary>
        /// 元の中間期決算期間
        /// </summary>
        public class BfMidKessan
        {
            public static string FromYear;                  //入力開始年
            public static string FromMonth;                 //入力開始月
            public static string FromDay;                   //入力開始日
            public static string StSoeji = string.Empty;    //入力期間開始添え字
            public static string ToYear;                    //入力期限年
            public static string ToMonth;                   //入力期限月
            public static string ToDay;                     //入力期限日
            public static string EdSoeji = string.Empty;    //入力期間終了添え字
            public static string Lock = string.Empty;       //制限の種類
            public static Boolean Flag = false;             //入力可能フラグ

            private string LimitDate;
            private string FromLimitDate;
            private string ToLimitDate;

            /// <summary>
            /// 元の中間期決算期間の取得
            /// </summary>
            public void GetKessanDate()
            {
                //開始日    
                LimitDate = company.FromYear + "/" + company.FromMonth + "/" + company.FromDay;

                //期首日に5ヶ月足す
                FromLimitDate = (DateTime.Parse(LimitDate).AddMonths(5)).ToString("yyyyMMdd");

                //中間期決算開始日を代入
                FromYear = FromLimitDate.Substring(0, 4);
                FromMonth = FromLimitDate.Substring(4, 2);
                FromDay = FromLimitDate.Substring(6, 2);

                //終了日
                LimitDate = company.ToYear + "/" + company.ToMonth + "/" + company.ToDay;

                //期末日から6ヶ月引く
                ToLimitDate = (DateTime.Parse(LimitDate).AddMonths(-6)).ToString("yyyyMMdd");

                //中間期決算終了日を代入
                ToYear = ToLimitDate.Substring(0, 4);
                ToMonth = ToLimitDate.Substring(4, 2);
                ToDay = ToLimitDate.Substring(6, 2);
            }
        }

        /// <summary>
        /// 決算期間
        /// </summary>
        public class KessanDate
        {
            public static string FromYear;    //入力開始年
            public static string FromMonth;   //入力開始月
            public static string FromDay;     //入力開始日
            public static string StSoeji;     //入力期間開始添え字
            public static string ToYear;      //入力期限年
            public static string ToMonth;     //入力期限月
            public static string ToDay;       //入力期限日
            public static string EdSoeji;     //入力期間終了添え字
            public static string Lock;        //制限の種類
            public static Boolean Flag;       //入力可能フラグ

            private string LimitDate;
            private string FromLimitDate;
            private string ToLimitDate = string.Empty;

            /// <summary>
            /// 決算期間の取得
            /// </summary>
            public void GetKessanDate()
            {
                //開始日    
                LimitDate = company.FromYear + "/" + company.FromMonth + "/" + company.FromDay;

                //期首日に11ヶ月足す
                FromLimitDate = (DateTime.Parse(LimitDate).AddMonths(11)).ToString("yyyyMMdd");

                //決算開始日を代入
                FromYear = FromLimitDate.Substring(0, 4);
                FromMonth = FromLimitDate.Substring(4, 2);
                FromDay = FromLimitDate.Substring(6, 2);

                //決算終了日を代入
                ToYear = company.ToYear;
                ToMonth = company.ToMonth;
                ToDay = company.ToDay;
            }
        }

        /// <summary>
        /// 元の決算期間
        /// </summary>
        public class BfKessan
        {
            public static string FromYear;                  //入力開始年
            public static string FromMonth;                 //入力開始月
            public static string FromDay;                   //入力開始日
            public static string StSoeji = string.Empty;    //入力期間開始添え字
            public static string ToYear;                    //入力期限年
            public static string ToMonth;                   //入力期限月
            public static string ToDay;                     //入力期限日
            public static string EdSoeji = string.Empty;    //入力期間終了添え字
            public static string Lock = string.Empty;       //制限の種類
            public static Boolean Flag = false;             //入力可能フラグ

            private string LimitDate;
            private string FromLimitDate;
            private string ToLimitDate = string.Empty;

            /// <summary>
            /// 元の決算期間の取得
            /// </summary>
            public void GetKessanDate()
            {
                //開始日    
                LimitDate = company.FromYear + "/" + company.FromMonth + "/" + company.FromDay;

                //期首日に11ヶ月足す
                FromLimitDate = (DateTime.Parse(LimitDate).AddMonths(11)).ToString("yyyyMMdd");

                //決算開始日を代入
                FromYear = FromLimitDate.Substring(0, 4);
                FromMonth = FromLimitDate.Substring(4, 2);
                FromDay = FromLimitDate.Substring(6, 2);

                //決算終了日を代入
                ToYear = company.ToYear;
                ToMonth = company.ToMonth;
                ToDay = company.ToDay;
            }
        }

        /// <summary>
        /// パラメータ日付の翌日日付を求める
        /// </summary>
        public class GetNextDay
        {
            private DateTime bDate;

            /// <summary>
            /// パラメータ日付の翌日日付を求める
            /// </summary>
            /// <param name="tempDate">日付</param>
            public GetNextDay(DateTime tempDate)
            {
                //パラメータ日付に1日足す
                bDate = tempDate.AddDays(1);
            }

            /// <summary>
            /// 翌日の年を取得する
            /// </summary>
            /// <returns>年：文字型</returns>
            public string GetYear()
            {
                return bDate.ToString().Substring(0, 4);
            }

            /// <summary>
            /// 翌日の月を取得する
            /// </summary>
            /// <returns>月：文字型</returns>
            public string GetMonth()
            {
                return bDate.ToString().Substring(4, 2);
            }
            
            /// <summary>
            /// 翌日の日を取得する
            /// </summary>
            /// <returns>日：文字型</returns>
            public string GetDay()
            {
                return bDate.ToString().Substring(6, 2);
            }

            /// <summary>
            /// 翌日の日付を取得する
            /// </summary>
            /// <returns>翌日日付：日付型</returns>
            public DateTime GetDate()
            {
                return bDate;
            }

        }
    }
}
