using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Drawing;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Drawing.Printing;
using Leadtools.WinForms;

namespace mntsiwake
{
    class cPrint
    {
        //プリント制御
        private int PRINT_Den;                          //印刷伝票
        private int PRINTMAXGYOU = 5;                   //最大印刷伝票数
        private int PRINTFONTSIZE = 8;                  //印刷フォントサイズ
        private int PrintMode;                          //全部印刷、一枚印刷の区分
        private int PrintPage = 1;                      //頁カウント
        private int Loopcnt = 0;                        //印刷データ数カウント
        private int wrkFirstDen = 0;
        private decimal KariSum = 0;
        private decimal KashiSum = 0;

        private float PrnX;                             //印刷位置X
        private float PrnY;                             //印刷位置Y
        private RasterImageViewer prnImage;             //印刷するLeadTools画像

        Entity.InputRecord[] DenData;                   //伝票データ配列

        /// <summary>
        /// 伝票画像を印刷する
        /// </summary>
        /// <param name="Img">LeadTools画像</param>
        public void Image(RasterImageViewer Img)
        {
            prnImage = Img;
            PrintDocument PrnImg = new PrintDocument();
            PrnImg.PrinterSettings = new PrinterSettings();

            //用紙方向：縦
            PrnImg.DefaultPageSettings.Landscape = false;

            //用紙サイズ：A4
            foreach (System.Drawing.Printing.PaperSize ps in PrnImg.PrinterSettings.PaperSizes)
            {
                if (ps.Kind == System.Drawing.Printing.PaperKind.A4)
                {
                    PrnImg.DefaultPageSettings.PaperSize = ps;
                    break;
                }
            }

            //印刷実行
            PrnImg.PrintPage += new PrintPageEventHandler(Image_PrintPage);
            PrnImg.Print();
        }

        /// <summary>
        /// 伝票画像印刷イベントハンドラ
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Image_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            //画像印刷
            int savePage = prnImage.Image.Page;

            try
            {
                using (Image img = prnImage.Image.ConvertToGdiPlusImage())
                {
                    e.Graphics.DrawImage(img, 0, 0);
                }
            }
            catch (Exception eX)
            {
                MessageBox.Show("伝票画像印刷中に不具合が発生したため印刷を中断します" + Environment.NewLine + eX.Message, "エラーメッセージ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                e.HasMorePages = false;
            }
            return;
        }

        public void Denpyo(Entity.InputRecord[] st, int dIndex, int pMode)
        {
            //伝票データ
            DenData = st;

            //現在の伝票
            PRINT_Den = dIndex;     

            //プリントモード
            PrintMode = pMode;

            //印刷設定
            System.Drawing.Printing.PrintDocument PrnDen = new System.Drawing.Printing.PrintDocument();
            PrnDen.PrinterSettings = new System.Drawing.Printing.PrinterSettings();

            //用紙方向：ヨコ
            PrnDen.DefaultPageSettings.Landscape = true;　

            //用紙サイズ：A4
            foreach (System.Drawing.Printing.PaperSize ps in PrnDen.PrinterSettings.PaperSizes)
            {
                if (ps.Kind == System.Drawing.Printing.PaperKind.A4)
                {
                    PrnDen.DefaultPageSettings.PaperSize = ps;
                    break;
                }
            }

            //印刷実行
            PrnDen.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(Denpyo_Print);
            PrnDen.Print();
        }

        /// <summary>
        /// 伝票認識内容印刷イベントハンドラ
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Denpyo_Print(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            string wrkWord;
            decimal KariMinSum;
            decimal KashiMinSum;
            Boolean EndFlg;
            int wrkPrintDen;

            Pen pen = new Pen(Color.Black, 1);  // Create pen
            Point[] points = new Point[2];      // ポイント構造体の配列       

            int wrkXBase = 4;
            int wrkYBase = 1;

            int pX = wrkXBase;
            int pY = wrkYBase;

            //複数チェック付きの伝票を印刷可能とする
            //複写チェックがなくなるか、先頭の伝票まで前に戻る
            if (PrintPage == 1)
            {
                for (int Cnt = PRINT_Den; Cnt >= 0; Cnt--)
                {
                    if (DenData[Cnt].Head.FukusuChk == "0")
                    {
                        wrkFirstDen = Cnt;
                        break;
                    }
                }
            }

            //伝票ヘッダ出力
            WritePrintHead("伝票認識内容", pX, pY, e);

            //ページ番号
            pX = wrkXBase + 100;
            SetXY(pX, pY, e);
            e.Graphics.DrawString("Page " + PrintPage.ToString(), new Font("ＭＳ ゴシック", PRINTFONTSIZE), Brushes.Black, PrnX, PrnY);

            pY++;

            //ライン
            pY++;

            //伝票行ヘッダ
            pY++;
            pX = wrkXBase;
            WriteGyouHead(pX, pY, e);

            //ライン
            pY = pY + 3;
            SetXY(wrkXBase, pY, e);
            points[0] = new Point((int)PrnX, (int)PrnY);
            points[1] = new Point((int)1100, (int)PrnY);
            e.Graphics.DrawLines(pen, points);

            while (true)
            {
                //初回起動時：１枚印刷なら、引数の伝票のみ印刷
                if (PrintPage == 1)
                {
                    if (PrintMode == global.PRINTMODEONE)
                    {
                        wrkPrintDen = PRINT_Den;
                    }
                    else
                    {
                        wrkPrintDen = wrkFirstDen + Loopcnt;
                    }
                }
                //2ページ以降
                else
                {
                    wrkPrintDen = wrkFirstDen + Loopcnt;
                }

                //伝票ヘッダデータ書込み
                pY = pY + 2;
                pX = wrkXBase;
                WriteDenHead(wrkPrintDen, pX, pY, e);

                //小計金額のクリア
                KariMinSum = 0;
                KashiMinSum = 0;


                //伝票行データ
                for (int i = 0; i < global.MAXGYOU; i++)
                {
                    pX = wrkXBase;
                    pY++;

                    if (DenData[wrkPrintDen].Gyou[i].Torikeshi == "0")
                    {
                        WriteGyouData(wrkPrintDen, i, pX, pY, e);

                        if (utility.NumericCheck(DenData[wrkPrintDen].Gyou[i].Kari.Kin))
                        {
                            KariSum += int.Parse(DenData[wrkPrintDen].Gyou[i].Kari.Kin);
                            KariMinSum += int.Parse(DenData[wrkPrintDen].Gyou[i].Kari.Kin);
                        }

                        if (utility.NumericCheck(DenData[wrkPrintDen].Gyou[i].Kashi.Kin))
                        {
                            KashiSum += int.Parse(DenData[wrkPrintDen].Gyou[i].Kashi.Kin);
                            KashiMinSum += int.Parse(DenData[wrkPrintDen].Gyou[i].Kashi.Kin);
                        }
                    }
                }

                //ライン
                pY++;
                SetXY(wrkXBase, pY, e);

                // ポイント構造体の配列
                points[0] = new Point((int)PrnX, (int)PrnY);
                points[1] = new Point((int)1100, (int)PrnY);

                //直線を引く
                e.Graphics.DrawLines(pen, points);

                pX = wrkXBase;
                //小計金額印字
                WriteMinSum(KariMinSum, KashiMinSum, pX, pY, e);

                //貸借ライン
                SetXY(44, pY - 7, e);

                // ポイント構造体の配列
                points[0] = new Point((int)PrnX, (int)PrnY);
                points[1] = new Point((int)PrnX, (int)PrnY + 88);
                e.Graphics.DrawLines(pen, points);

                //摘要ライン
                SetXY(83, pY - 7, e);
                points[0] = new Point((int)PrnX, (int)PrnY);
                points[1] = new Point((int)PrnX, (int)PrnY + 88);
                e.Graphics.DrawLines(pen, points);

                //データ数カウント
                Loopcnt++;

                //印刷終了判定
                EndFlg = false;

                //１枚印刷モードのとき
                if (PrintMode == global.PRINTMODEONE)
                {
                    EndFlg = true;
                    break;
                }
                //最終伝票に達したとき
                else if (wrkFirstDen + Loopcnt >= global.pblDenNum)
                {
                    EndFlg = true;
                    break;
                }
                //複数チェックがなくなったとき
                else if (DenData[wrkFirstDen + Loopcnt].Head.FukusuChk == "0")
                {
                    EndFlg = true;
                    break;
                }
                //頁あたりの印刷データ数完了（印刷データがまだあり）
                else if (Loopcnt == PRINTMAXGYOU)
                {
                    break;
                }
            }

            if (EndFlg == true)
            {
                //合計金額
                pY += 2;
                pX = wrkXBase;

                pX += 2;
                SetXY(pX, pY, e);
                e.Graphics.DrawString("金額合計：", new Font("ＭＳ ゴシック", PRINTFONTSIZE), Brushes.Black, PrnX, PrnY);

                pX += 24;
                SetXY(pX, pY, e);
                wrkWord = string.Format("{0:#,##0}", KariSum);
                wrkWord = string.Format("{0, 14}", wrkWord.Trim());
                e.Graphics.DrawString(wrkWord, new Font("ＭＳ ゴシック", PRINTFONTSIZE), Brushes.Black, PrnX, PrnY);

                pX += 39;
                SetXY(pX, pY, e);
                wrkWord = string.Format("{0:#,##0}", KashiSum);
                wrkWord = string.Format("{0, 14}", wrkWord.Trim());
                e.Graphics.DrawString(wrkWord, new Font("ＭＳ ゴシック", PRINTFONTSIZE), Brushes.Black, PrnX, PrnY);

                pX += 13;
                SetXY(pX, pY, e);
                e.Graphics.DrawString("貸借差額：" + System.Math.Abs(KariSum - KashiSum).ToString(), new Font("ＭＳ ゴシック", PRINTFONTSIZE), Brushes.Black, PrnX, PrnY);

                e.HasMorePages = false;
            }
            else
            {
                e.HasMorePages = true;
                PrintPage++;
            }

                ////5枚に1回改ページ
                //if ((Loopcnt % PRINTMAXGYOU) == 0)
                //{
                //    e.HasMorePages = true;
                //    pY = wrkYBase;
                //}
        }

        /// <summary>
        /// 印刷X座標Y座標設定
        /// </summary>
        /// <param name="X">X座標ピッチ</param>
        /// <param name="Y">Y座標ピッチ</param>
        /// <param name="eX"></param>
        private void SetXY(float X, float Y, System.Drawing.Printing.PrintPageEventArgs eX)
        {
            Font stringfont = new System.Drawing.Font("ＭＳ ゴシック", 8);
            PrnX = eX.Graphics.MeasureString("-", stringfont).Width * X;
            PrnY = eX.Graphics.MeasureString("-", stringfont).Height * Y;
        }

        /// <summary>
        /// 印刷ヘッダ出力
        /// </summary>
        /// <param name="sTitle">タイトル</param>
        /// <param name="Den">現在の伝票インデックス</param>
        /// <param name="X">X方向ピッチ</param>
        /// <param name="Y">Y方向ピッチ</param>
        /// <param name="e"></param>
        private void WritePrintHead(string sTitle, int X, int Y, System.Drawing.Printing.PrintPageEventArgs e)
        {
            int wrkNowX;
            int wrkNowY;
            wrkNowX = X + 40;
            wrkNowY = Y;

            SetXY(wrkNowX, wrkNowY, e);
            e.Graphics.DrawString(sTitle, new Font("ＭＳ ゴシック", PRINTFONTSIZE + 3, FontStyle.Bold), Brushes.Black, PrnX, PrnY);
        }

        /// <summary>
        /// 行ヘッダ部出力
        /// </summary>
        /// <param name="X">X座標ピッチ</param>
        /// <param name="Y">Y座標ピッチ</param>
        /// <param name="eX"></param>
        private void WriteGyouHead(int X, int Y, System.Drawing.Printing.PrintPageEventArgs eX)
        {
            int wrkNowX;
            int wrkNowY;

            wrkNowY = Y;
            wrkNowX = X;

            SetXY(wrkNowX, wrkNowY, eX);
            eX.Graphics.DrawString("［借　　方］", new Font("ＭＳ ゴシック", PRINTFONTSIZE), Brushes.Black, PrnX, PrnY);

            wrkNowX = wrkNowX + 39;
            SetXY(wrkNowX, wrkNowY, eX);
            eX.Graphics.DrawString("［貸　　方］", new Font("ＭＳ ゴシック", PRINTFONTSIZE), Brushes.Black, PrnX, PrnY);

            wrkNowY = Y + 2;
            wrkNowX = X;

            wrkNowX = wrkNowX + 1;
            SetXY(wrkNowX, wrkNowY, eX);
            eX.Graphics.DrawString("部門", new Font("ＭＳ ゴシック", PRINTFONTSIZE), Brushes.Black, PrnX, PrnY);

            wrkNowX = wrkNowX + 3;
            SetXY(wrkNowX, wrkNowY, eX);
            eX.Graphics.DrawString("科目", new Font("ＭＳ ゴシック", PRINTFONTSIZE), Brushes.Black, PrnX, PrnY);

            wrkNowX = wrkNowX + 11;
            SetXY(wrkNowX, wrkNowY, eX);
            eX.Graphics.DrawString("補助", new Font("ＭＳ ゴシック", PRINTFONTSIZE), Brushes.Black, PrnX, PrnY);

            wrkNowX = wrkNowX + 17;
            SetXY(wrkNowX, wrkNowY, eX);
            eX.Graphics.DrawString("金額", new Font("ＭＳ ゴシック", PRINTFONTSIZE), Brushes.Black, PrnX, PrnY);

            wrkNowX = wrkNowX + 3;
            SetXY(wrkNowX, wrkNowY, eX);
            eX.Graphics.DrawString("処 区", new Font("ＭＳ ゴシック", PRINTFONTSIZE), Brushes.Black, PrnX, PrnY);

            //wrkNowX = wrkNowX + 2;
            //SetXY(wrkNowX, wrkNowY, eX);
            //eX.Graphics.DrawString("区",new Font("ＭＳ ゴシック",PRINTFONTSIZE), Brushes.Black, PrnX, PrnY);

            wrkNowX = wrkNowX + 5;
            SetXY(wrkNowX, wrkNowY, eX);
            eX.Graphics.DrawString("部門", new Font("ＭＳ ゴシック", PRINTFONTSIZE), Brushes.Black, PrnX, PrnY);

            wrkNowX = wrkNowX + 3;
            SetXY(wrkNowX, wrkNowY, eX);
            eX.Graphics.DrawString("科目", new Font("ＭＳ ゴシック", PRINTFONTSIZE), Brushes.Black, PrnX, PrnY);

            wrkNowX = wrkNowX + 11;
            SetXY(wrkNowX, wrkNowY, eX);
            eX.Graphics.DrawString("補助", new Font("ＭＳ ゴシック", PRINTFONTSIZE), Brushes.Black, PrnX, PrnY);

            wrkNowX = wrkNowX + 17;
            SetXY(wrkNowX, wrkNowY, eX);
            eX.Graphics.DrawString("金額", new Font("ＭＳ ゴシック", PRINTFONTSIZE), Brushes.Black, PrnX, PrnY);

            wrkNowX = wrkNowX + 3;
            SetXY(wrkNowX, wrkNowY, eX);
            eX.Graphics.DrawString("処 区", new Font("ＭＳ ゴシック", PRINTFONTSIZE), Brushes.Black, PrnX, PrnY);

            //wrkNowX = wrkNowX + 2;
            //SetXY(wrkNowX, wrkNowY, eX);
            //eX.Graphics.DrawString("区",new Font("ＭＳ ゴシック",PRINTFONTSIZE), Brushes.Black, PrnX, PrnY);

            wrkNowX = wrkNowX + 5;
            SetXY(wrkNowX, wrkNowY, eX);
            eX.Graphics.DrawString("複写", new Font("ＭＳ ゴシック", PRINTFONTSIZE), Brushes.Black, PrnX, PrnY);

            wrkNowX = wrkNowX + 3;
            SetXY(wrkNowX, wrkNowY, eX);
            eX.Graphics.DrawString("摘要", new Font("ＭＳ ゴシック", PRINTFONTSIZE), Brushes.Black, PrnX, PrnY);
        }

        /// <summary>
        /// 伝票ヘッダデータ出力
        /// </summary>
        /// <param name="Den">現在の伝票インデックス</param>
        /// <param name="X">X座標ピッチ</param>
        /// <param name="Y">Y座標ピッチ</param>
        /// <param name="eX"></param>
        private void WriteDenHead(int Den, float X, float Y, System.Drawing.Printing.PrintPageEventArgs eX)
        {
            float wrkNowX;
            float wrkNowY;
            float wrkXBase;
            string wrkKessan;
            string wrkFukusu;

            wrkNowX = X;
            wrkNowY = Y;
            wrkXBase = 4;

            if (DenData[Den].Head.Kessan == "1")
            {
                wrkKessan = "●";
            }
            else
            {
                wrkKessan = "　";
            }

            if (DenData[Den].Head.FukusuChk == "1")
            {
                wrkFukusu = "●";
            }
            else
            {
                wrkFukusu = "　";
            }

            SetXY(wrkNowX, wrkNowY, eX);

            string pStr = string.Empty;
            pStr += "日付：";
            pStr += string.Format("{0,2}", DenData[Den].Head.Year) + "年 ";
            pStr += string.Format("{0,2}", DenData[Den].Head.Month) + "月 ";
            pStr += string.Format("{0,2}", DenData[Den].Head.Day) + "日　　";
            pStr += "伝票No.：" + string.Format("{0,6}", DenData[Den].Head.DenNo) + "　　";
            pStr += "決算：" + wrkKessan + "　";
            pStr += "伝票結合：" + wrkFukusu;

            eX.Graphics.DrawString(pStr, new Font("ＭＳ ゴシック", PRINTFONTSIZE), Brushes.Black, PrnX, PrnY);

            //ライン
            wrkNowY++;
            SetXY(wrkXBase, wrkNowY, eX);
            Point[] Points = { new Point((int)PrnX, (int)PrnY), new Point(1100, (int)PrnY) };
            eX.Graphics.DrawLines(new Pen(Color.Black, 1), Points);
        }

        /// <summary>
        /// 行データ出力
        /// </summary>
        /// <param name="Den">現在の伝票インデックス</param>
        /// <param name="Gyou">行インデックス</param>
        /// <param name="X">X座標ピッチ</param>
        /// <param name="Y">Y座標ピッチ</param>
        /// <param name="eX"></param>
        private void WriteGyouData(int Den, int Gyou, int X, int Y, System.Drawing.Printing.PrintPageEventArgs eX)
        {
            string wrkWord;
            string wrkName;

            int wrkNowX = X;
            int wrkNowY = Y;

            //借方部門
            wrkNowX++;
            SetXY(wrkNowX, wrkNowY, eX);
            if (DenData[Den].Gyou[Gyou].Kari.Bumon != string.Empty)
            {
                wrkWord = string.Format("{0,4}", DenData[Den].Gyou[Gyou].Kari.Bumon);
                eX.Graphics.DrawString(wrkWord, new Font("ＭＳ ゴシック", PRINTFONTSIZE), Brushes.Black, PrnX, PrnY);
            }

            //借方科目
            wrkNowX += 3;
            SetXY(wrkNowX, wrkNowY, eX);
            if (DenData[Den].Gyou[Gyou].Kari.Kamoku != string.Empty)
            {
                wrkWord = string.Format("{0,4}", DenData[Den].Gyou[Gyou].Kari.Kamoku);
                eX.Graphics.DrawString(wrkWord, new Font("ＭＳ ゴシック", PRINTFONTSIZE), Brushes.Black, PrnX, PrnY);
            }

            //借方科目名
            wrkNowX += 3;
            SetXY(wrkNowX, wrkNowY, eX);

            if (DenData[Den].Gyou[Gyou].Kari.Kamoku != string.Empty)
            {
                wrkName = KamokuCodeToName(DenData[Den].Gyou[Gyou].Kari.Kamoku);
                if (wrkName.Length > 7) wrkName = wrkName.Substring(0, 7);
                eX.Graphics.DrawString(wrkName, new Font("ＭＳ ゴシック", PRINTFONTSIZE), Brushes.Black, PrnX, PrnY);
            }

            //借方補助
            wrkNowX += 8;
            SetXY(wrkNowX, wrkNowY, eX);

            if (DenData[Den].Gyou[Gyou].Kari.Hojo != string.Empty)
            {
                wrkWord = string.Format("{0, 4}", DenData[Den].Gyou[Gyou].Kari.Hojo);
                eX.Graphics.DrawString(wrkWord, new Font("ＭＳ ゴシック", PRINTFONTSIZE), Brushes.Black, PrnX, PrnY);
            }

            //借方補助名
            wrkNowX += 3;
            SetXY(wrkNowX, wrkNowY, eX);
            if (DenData[Den].Gyou[Gyou].Kari.Hojo != string.Empty)
            {
                wrkName =
                wrkName = HojoCodeToName(DenData[Den].Gyou[Gyou].Kari.Kamoku, DenData[Den].Gyou[Gyou].Kari.Hojo);
                if (wrkName.Length > 7) wrkName = wrkName.Substring(0, 7);
                eX.Graphics.DrawString(wrkName, new Font("ＭＳ ゴシック", PRINTFONTSIZE), Brushes.Black, PrnX, PrnY);
            }

            //借方金額
            wrkNowX += 8;
            SetXY(wrkNowX, wrkNowY, eX);
            wrkWord = DenData[Den].Gyou[Gyou].Kari.Kin;
            if (utility.NumericCheck(wrkWord))
            {
                wrkWord = string.Format("{0:#,##0}", int.Parse(wrkWord));
                wrkWord = string.Format("{0, 14}", wrkWord);
                eX.Graphics.DrawString(wrkWord, new Font("ＭＳ ゴシック", PRINTFONTSIZE), Brushes.Black, PrnX, PrnY);
            }

            //借方消費税計算区分
            wrkNowX += 10;
            SetXY(wrkNowX, wrkNowY, eX);
            wrkWord = DenData[Den].Gyou[Gyou].Kari.TaxMas;
            if (wrkWord != String.Empty)
            {
                eX.Graphics.DrawString(wrkWord, new Font("ＭＳ ゴシック", PRINTFONTSIZE), Brushes.Black, PrnX, PrnY);
            }

            //借方消費税区分
            wrkNowX += 1;
            SetXY(wrkNowX, wrkNowY, eX);
            wrkWord = DenData[Den].Gyou[Gyou].Kari.TaxKbn;
            if (wrkWord != String.Empty)
            {
                eX.Graphics.DrawString(wrkWord, new Font("ＭＳ ゴシック", PRINTFONTSIZE), Brushes.Black, PrnX, PrnY);
            }

            //貸方部門
            wrkNowX += 3;
            SetXY(wrkNowX, wrkNowY, eX);
            if (DenData[Den].Gyou[Gyou].Kashi.Bumon != string.Empty)
            {
                wrkWord = string.Format("{0,4}", DenData[Den].Gyou[Gyou].Kashi.Bumon);
                eX.Graphics.DrawString(wrkWord, new Font("ＭＳ ゴシック", PRINTFONTSIZE), Brushes.Black, PrnX, PrnY);
            }

            //貸方科目
            wrkNowX += 3;
            SetXY(wrkNowX, wrkNowY, eX);
            if (DenData[Den].Gyou[Gyou].Kashi.Kamoku != string.Empty)
            {
                wrkWord = string.Format("{0,4}", DenData[Den].Gyou[Gyou].Kashi.Kamoku);
                eX.Graphics.DrawString(wrkWord, new Font("ＭＳ ゴシック", PRINTFONTSIZE), Brushes.Black, PrnX, PrnY);
            }

            //貸方科目名
            wrkNowX += 3;
            SetXY(wrkNowX, wrkNowY, eX);

            if (DenData[Den].Gyou[Gyou].Kashi.Kamoku != string.Empty)
            {
                wrkName = KamokuCodeToName(DenData[Den].Gyou[Gyou].Kashi.Kamoku);
                if (wrkName.Length > 7) wrkName = wrkName.Substring(0, 7);
                eX.Graphics.DrawString(wrkName, new Font("ＭＳ ゴシック", PRINTFONTSIZE), Brushes.Black, PrnX, PrnY);
            }

            //貸方補助
            wrkNowX += 8;
            SetXY(wrkNowX, wrkNowY, eX);

            if (DenData[Den].Gyou[Gyou].Kashi.Hojo != string.Empty)
            {
                wrkWord = string.Format("{0, 4}", DenData[Den].Gyou[Gyou].Kashi.Hojo);
                eX.Graphics.DrawString(wrkWord, new Font("ＭＳ ゴシック", PRINTFONTSIZE), Brushes.Black, PrnX, PrnY);
            }

            //貸方補助名
            wrkNowX += 3;
            SetXY(wrkNowX, wrkNowY, eX);
            if (DenData[Den].Gyou[Gyou].Kashi.Hojo != string.Empty)
            {
                wrkName = HojoCodeToName(DenData[Den].Gyou[Gyou].Kashi.Kamoku, DenData[Den].Gyou[Gyou].Kashi.Hojo);
                if (wrkName.Length > 7) wrkName = wrkName.Substring(0, 7);
                eX.Graphics.DrawString(wrkName, new Font("ＭＳ ゴシック", PRINTFONTSIZE), Brushes.Black, PrnX, PrnY);
            }

            //貸方金額
            wrkNowX += 8;
            SetXY(wrkNowX, wrkNowY, eX);
            wrkWord = DenData[Den].Gyou[Gyou].Kashi.Kin;
            if (utility.NumericCheck(wrkWord))
            {
                wrkWord = string.Format("{0:#,##0}", int.Parse(wrkWord));
                wrkWord = string.Format("{0, 14}", wrkWord);
                eX.Graphics.DrawString(wrkWord, new Font("ＭＳ ゴシック", PRINTFONTSIZE), Brushes.Black, PrnX, PrnY);
            }

            //貸方消費税計算区分
            wrkNowX += 10;
            SetXY(wrkNowX, wrkNowY, eX);
            wrkWord = DenData[Den].Gyou[Gyou].Kashi.TaxMas;
            if (wrkWord != String.Empty)
            {
                eX.Graphics.DrawString(wrkWord, new Font("ＭＳ ゴシック", PRINTFONTSIZE), Brushes.Black, PrnX, PrnY);
            }

            //貸方消費税区分
            wrkNowX += 1;
            SetXY(wrkNowX, wrkNowY, eX);
            wrkWord = DenData[Den].Gyou[Gyou].Kashi.TaxKbn;
            if (wrkWord != String.Empty)
            {
                eX.Graphics.DrawString(wrkWord, new Font("ＭＳ ゴシック", PRINTFONTSIZE), Brushes.Black, PrnX, PrnY);
            }

            //複写
            wrkNowX += 4;
            SetXY(wrkNowX, wrkNowY, eX);
            if (DenData[Den].Gyou[Gyou].CopyChk != "0")
            {
                eX.Graphics.DrawString("●", new Font("ＭＳ ゴシック", PRINTFONTSIZE), Brushes.Black, PrnX, PrnY);
            }

            //摘要
            wrkNowX += 2;
            SetXY(wrkNowX, wrkNowY, eX);
            wrkWord = DenData[Den].Gyou[Gyou].Tekiyou;
            if (wrkWord != String.Empty)
            {
                eX.Graphics.DrawString(wrkWord, new Font("ＭＳ ゴシック", PRINTFONTSIZE), Brushes.Black, PrnX, PrnY);
            }
        }

        ///----------------------------------------------------------
        /// <summary>
        ///     勘定科目名取得 : 2017/06/06 </summary>
        /// <param name="cCode">
        ///     勘定科目コード</param>
        /// <returns>
        ///     勘定科目名</returns>
        ///----------------------------------------------------------
        private string KamokuCodeToName(string cCode)
        {
            string KanjoCode = string.Empty;
            string GetName = string.Empty;

            // 接続文字列取得 2017/9/13
            string sc = SqlControl.obcConnectSting.get(global.pblDbName);

            // コントロールインスタンス生成
            SqlControl.DataControl sCon = new SqlControl.DataControl(sc);   
            string mySql = string.Empty;
            SqlDataReader dR;               //データリーダー

            //勘定科目取得
            if (utility.NumericCheck(cCode.Trim()))
            {
                KanjoCode = string.Format("{0:D10}", int.Parse(cCode.Trim()));
            }
            else
            {
                KanjoCode = cCode.Trim();
            }

            mySql += "SELECT AccountItemCode, AccountItemName FROM tbAccountItem ";
            mySql += "WHERE (tbAccountItem.IsUse = 1) and ";
            mySql += "(tbAccountItem.AccountingPeriodID = " + global.pblAccPID + ") and ";
            mySql += "(AccountItemCode = '" + KanjoCode + "')";

            //データリーダーを取得する
            dR = sCon.free_dsReader(mySql);

            if (dR.HasRows)
            {
                dR.Read();
                GetName = dR["AccountItemName"].ToString().Trim();
            }
            else
            {
                GetName = string.Empty;
            }

            dR.Close();
            sCon.Close();

            return GetName;
        }

        ///------------------------------------------------------
        /// <summary>
        ///     補助科目名取得 : 2017/06/06 </summary>
        /// <param name="cCode">
        ///     勘定科目コード</param>
        /// <param name="hCode">
        ///     補助科目コード</param>
        /// <returns>
        ///     補助科目名</returns>
        ///------------------------------------------------------
        private string HojoCodeToName(string cCode, string hCode)
        {
            string GetName = string.Empty;
            string KanjoCode = string.Empty;
            string hojoCode = string.Empty;
            Boolean hCodestatus = false;    //補助コードの有無ステータス

            // 補助科目無記入のとき
            if (hCode.Trim() == string.Empty)
            {
                return string.Empty;
            }

            // 接続文字列取得 2017/06/06
            string sc = SqlControl.obcConnectSting.get(global.pblDbName);

            //string sc = utility.GetDBConnect(global.pblDbName);             //SQLServer接続文字列取得
            SqlControl.DataControl sCon = new SqlControl.DataControl(sc);   //コントロールインスタンス生成
            string mySql = string.Empty;
            SqlDataReader dR;                                             //データリーダー

            //勘定科目取得
            if (utility.NumericCheck(cCode.Trim()))
            {
                KanjoCode = string.Format("{0:D10}", int.Parse(cCode.Trim()));
            }
            else
            {
                KanjoCode = cCode.Trim();
            }

            //補助科目取得
            if (utility.NumericCheck(hCode.Trim()))
            {
                hojoCode = string.Format("{0:D10}", int.Parse(hCode.Trim()));
            }
            else
            {
                hojoCode = hCode.Trim();
            }

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
            hCodestatus = dR.HasRows;
            dR.Close();

            //勘定科目に補助コードが登録されているとき
            if (hCodestatus)
            {
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
                        GetName = dR["SubAccountItemName"].ToString().Trim();
                        break;
                    }
                }

                dR.Close();
            }

            sCon.Close();

            return GetName;
        }

        /// <summary>
        /// 小計行出力
        /// </summary>
        /// <param name="Kari">借方合計金額</param>
        /// <param name="Kashi">貸方合計金額</param>
        /// <param name="X">X座標ピッチ</param>
        /// <param name="Y">Y座標ピッチ</param>
        /// <param name="eX"></param>
        private void WriteMinSum(decimal Kari, decimal Kashi, int X, int Y, System.Drawing.Printing.PrintPageEventArgs eX)
        {
            string wrkWord;

            int wrkNowX = X;
            int wrkNowY = Y;

            string wrkKari = Kari.ToString();
            string wrkKashi = Kashi.ToString();

            wrkNowX += 2;
            SetXY(wrkNowX, wrkNowY, eX);
            eX.Graphics.DrawString("　小　計：", new Font("ＭＳ ゴシック", PRINTFONTSIZE), Brushes.Black, PrnX, PrnY);

            wrkNowX += 23;
            SetXY(wrkNowX, wrkNowY, eX);
            wrkWord = string.Format("{0:#,##0}", Kari);
            wrkWord = "(" + string.Format("{0,14}", wrkWord) + ")";
            eX.Graphics.DrawString(wrkWord, new Font("ＭＳ ゴシック", PRINTFONTSIZE), Brushes.Black, PrnX, PrnY);

            wrkNowX += 39;
            SetXY(wrkNowX, wrkNowY, eX);
            wrkWord = string.Format("{0:#,##0}", Kashi);
            wrkWord = "(" + string.Format("{0,14}", wrkWord) + ")";
            eX.Graphics.DrawString(wrkWord, new Font("ＭＳ ゴシック", PRINTFONTSIZE), Brushes.Black, PrnX, PrnY);

        }
    }
}
