using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Drawing.Text;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace BlackcatShippingOrderPrintHelper
{
    public static class Agent
    {

        private static List<OrderInfo> _OrderPool { get; set; }
        private static string _SamplePath { get; set; }
       
        private static Graphics _Graphics;

        public static void CearAllOrderInfo()
        {
            _OrderPool.Clear();
        }

        public static void AddOrderInfo(OrderInfo o)
        {
            if (_OrderPool.Count >= 3)
            {
                throw new Exception("超過三個囉");
            }

            _OrderPool.Add(o);
        }

        public static void LoadSampleImage(string filepath)
        {
            _SamplePath = filepath;

        }

        private static Bitmap GetCode39Bitmap(string strSource)
        {
            int x = 5; 
            int y = 0; 
            int wideLength = 16; 
            int narrowLength = 8; 
            int BarCodeHeight = 192; 
            int intSourceLength = strSource.Length;
            string strEncode = "010010100"; //編碼字串 初值為 起始符號 *

            string AlphaBet = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ-. $/+%*"; //Code39的字母

            string[] Code39 = //Code39的各字母對應碼
            {
                 /**//* 0 */ "000110100",  
                 /**//* 1 */ "100100001",  
                 /**//* 2 */ "001100001",  
                 /**//* 3 */ "101100000",
                 /**//* 4 */ "000110001",  
                 /**//* 5 */ "100110000",  
                 /**//* 6 */ "001110000",  
                 /**//* 7 */ "000100101",
                 /**//* 8 */ "100100100",  
                 /**//* 9 */ "001100100",  
                 /**//* A */ "100001001",  
                 /**//* B */ "001001001",
                 /**//* C */ "101001000",  
                 /**//* D */ "000011001",  
                 /**//* E */ "100011000",  
                 /**//* F */ "001011000",
                 /**//* G */ "000001101",  
                 /**//* H */ "100001100",  
                 /**//* I */ "001001100",  
                 /**//* J */ "000011100",
                 /**//* K */ "100000011",  
                 /**//* L */ "001000011",  
                 /**//* M */ "101000010",  
                 /**//* N */ "000010011",
                 /**//* O */ "100010010",  
                 /**//* P */ "001010010",  
                 /**//* Q */ "000000111",  
                 /**//* R */ "100000110",
                 /**//* S */ "001000110",  
                 /**//* T */ "000010110",  
                 /**//* U */ "110000001",  
                 /**//* V */ "011000001",
                 /**//* W */ "111000000",  
                 /**//* X */ "010010001",  
                 /**//* Y */ "110010000",  
                 /**//* Z */ "011010000",
                 /**//* - */ "010000101",  
                 /**//* . */ "110000100",  
                 /**//*' '*/ "011000100",
                 /**//* $ */ "010101000",
                 /**//* / */ "010100010",  
                 /**//* + */ "010001010",  
                 /**//* % */ "000101010",  
                 /**//* * */ "010010100"
            };
            strSource = strSource.ToUpper();
            //實作圖片
            Bitmap objBitmap = new Bitmap(
              ((wideLength * 3 + narrowLength * 7) * (intSourceLength + 2)) + (x * 2),
              BarCodeHeight + (y * 2));
            Graphics objGraphics = Graphics.FromImage(objBitmap); //宣告GDI+繪圖介面
                                                                  //填上底色
            objGraphics.FillRectangle(Brushes.White, 0, 0, objBitmap.Width, objBitmap.Height);

            for (int i = 0; i < intSourceLength; i++)
            {
                //檢查是否有非法字元
                if (AlphaBet.IndexOf(strSource[i]) == -1 || strSource[i] == '*')
                {
                    objGraphics.DrawString("含有非法字元",
                      SystemFonts.DefaultFont, Brushes.Red, x, y);
                    return objBitmap;
                }
                //查表編碼
                strEncode = string.Format("{0}0{1}", strEncode,
                 Code39[AlphaBet.IndexOf(strSource[i])]);
            }

            strEncode = string.Format("{0}0010010100", strEncode); //補上結束符號 *

            int intEncodeLength = strEncode.Length; //編碼後長度
            int intBarWidth;

            for (int i = 0; i < intEncodeLength; i++) //依碼畫出Code39 BarCode
            {
                intBarWidth = strEncode[i] == '1' ? wideLength : narrowLength;
                objGraphics.FillRectangle(i % 2 == 0 ? Brushes.Black : Brushes.White,
                 x, y, intBarWidth, BarCodeHeight);
                x += intBarWidth;
            }
            return objBitmap;
        }


        static Agent()
        {
            _OrderPool = new List<OrderInfo>();
        }

        public static void Dispose()
        {
            _OrderPool = new List<OrderInfo>();
        }

        public static void ExportToJPG(string targetFilePath)
        {
            var rBitmap = (Bitmap)Image.FromFile(_SamplePath);
            _Graphics = Graphics.FromImage(rBitmap);

            #region 第一個
            if (_OrderPool.Count >= 1)
            {
                Brush br = new SolidBrush(Color.Black);

                _Graphics.DrawString(_OrderPool[0].包裹查詢號碼.轉運站代號, new Font(new FontFamily("Arial"), 12), br, new PointF(366, 125));
                _Graphics.DrawString(_OrderPool[0].包裹查詢號碼.轉運單號, new Font(new FontFamily("Arial"), 18), br, new PointF(546, 115));

                //配送資訊
                _Graphics.DrawString(_OrderPool[0].配送資訊.收貨日.ToString("yyyy-MM-dd"), new Font(new FontFamily("Arial"), 10), br, new PointF(95, 280));
                _Graphics.DrawString(_OrderPool[0].配送資訊.預定配達日.ToString("yyyy-MM-dd"), new Font(new FontFamily("Arial"), 10), br, new PointF(355, 280));
                _Graphics.DrawString(_OrderPool[0].配送資訊.指定時段, new Font(new FontFamily("Arial"), 9), br, new PointF(650, 280));

                if (_OrderPool[0].配送資訊.發貨所.Length <= 5)
                {
                    var tmpPlace = _OrderPool[0].配送資訊.發貨所;
                    for (int i = 0; i < 5 - _OrderPool[0].配送資訊.發貨所.Length; i++)
                    {
                        tmpPlace = "  " + tmpPlace;
                    }
                    _Graphics.DrawString(tmpPlace, new Font(new FontFamily("Arial"), 9), br, new PointF(885, 280));

                }

                //收件人
                _Graphics.DrawString(_OrderPool[0].收件人資訊.郵遞區號, new Font(new FontFamily("Arial"), 9), br, new PointF(150, 340));
                _Graphics.DrawString(_OrderPool[0].收件人資訊.地址, new Font(new FontFamily("Arial"), 10), br, new PointF(315, 340));
                _Graphics.DrawString(_OrderPool[0].收件人資訊.姓名, new Font(new FontFamily("Arial"), 10), br, new PointF(150, 410));
                _Graphics.DrawString(_OrderPool[0].收件人資訊.電話1, new Font(new FontFamily("Arial"), 10), br, new PointF(150, 480));
                _Graphics.DrawString(_OrderPool[0].收件人資訊.電話2, new Font(new FontFamily("Arial"), 10), br, new PointF(600, 480));


                //寄件人
                _Graphics.DrawString(_OrderPool[0].寄件人資訊.地址, new Font(new FontFamily("Arial"), 9), br, new PointF(150, 565));
                _Graphics.DrawString(_OrderPool[0].寄件人資訊.姓名, new Font(new FontFamily("Arial"), 8), br, new PointF(150, 620));
                _Graphics.DrawString(_OrderPool[0].寄件人資訊.電話1, new Font(new FontFamily("Arial"), 9), br, new PointF(150, 670));
                _Graphics.DrawString(_OrderPool[0].寄件人資訊.電話2, new Font(new FontFamily("Arial"), 9), br, new PointF(600, 670));

                //備註
                _Graphics.DrawString(_OrderPool[0].備註, new Font(new FontFamily("Arial"), 9), br, new PointF(180, 740));

                //品名
                _Graphics.DrawString(_OrderPool[0].品名, new Font(new FontFamily("Arial"), 9), br, new PointF(180, 810));


                //訂單編號
                _Graphics.DrawString(_OrderPool[0].訂單編號, new Font(new FontFamily("Arial"), 10), br, new PointF(260, 890));


                //客戶代號
                //BarCode
                var bmp1 = GetCode39Bitmap(_OrderPool[0].客戶代號);
                _Graphics.DrawImage(bmp1, new Rectangle(205, 990, 520, 50));
               

                //代號
                _Graphics.DrawString(_OrderPool[0].客戶代號, new Font(new FontFamily("Arial"), 10), br, new PointF(350, 1050));

                //金額
                _Graphics.DrawString(_OrderPool[0].代收金額.ToString("#"), new Font(new FontFamily("Arial"), 12), br, new PointF(920, 1010));

                //郵號條碼
                //Barcode
                var bmp2 = GetCode39Bitmap("+" + _OrderPool[0].郵號條碼.Replace("-", ""));
                _Graphics.DrawImage(bmp2, new Rectangle(1195, 60 , 510, 150));

                bmp2.Dispose();
             
                //Char
                _Graphics.DrawString(_OrderPool[0].郵號條碼, new Font(new FontFamily("Arial"), 28), br, new PointF(1695, 75));


                //收件人- 右
                _Graphics.DrawString(_OrderPool[0].收件人資訊.地址, new Font(new FontFamily("Arial"), 9), br, new PointF(1260, 260));
                _Graphics.DrawString(_OrderPool[0].收件人資訊.姓名, new Font(new FontFamily("Arial"), 10), br, new PointF(1260, 350));
                _Graphics.DrawString(_OrderPool[0].收件人資訊.電話1, new Font(new FontFamily("Arial"), 10), br, new PointF(1260, 410));
                _Graphics.DrawString(_OrderPool[0].收件人資訊.電話2, new Font(new FontFamily("Arial"), 10), br, new PointF(1660, 410));

                //寄件人- 右
                _Graphics.DrawString(_OrderPool[0].寄件人資訊.地址, new Font(new FontFamily("Arial"), 9), br, new PointF(1260, 490));
                _Graphics.DrawString(_OrderPool[0].寄件人資訊.姓名, new Font(new FontFamily("Arial"), 9), br, new PointF(1260, 540));
                _Graphics.DrawString(_OrderPool[0].寄件人資訊.電話1, new Font(new FontFamily("Arial"), 10), br, new PointF(1260, 600));
                _Graphics.DrawString(_OrderPool[0].寄件人資訊.電話2, new Font(new FontFamily("Arial"), 10), br, new PointF(1660, 600));

                //備註- 右
                _Graphics.DrawString(_OrderPool[0].備註, new Font(new FontFamily("Arial"), 9), br, new PointF(1315, 660));

                //品名- 右
                _Graphics.DrawString(_OrderPool[0].品名, new Font(new FontFamily("Arial"), 8), br, new PointF(1315, 715));


                //訂單編號- 右
                _Graphics.DrawString(_OrderPool[0].訂單編號, new Font(new FontFamily("Arial"), 8), br, new PointF(1420, 780));


                //收貨日- 右
                _Graphics.DrawString(_OrderPool[0].配送資訊.收貨日.ToString("yyyy-MM-dd"), new Font(new FontFamily("Arial"), 10), br, new PointF(1320, 835));

                //金額- 右
                _Graphics.DrawString(_OrderPool[0].代收金額.ToString("#"), new Font(new FontFamily("Arial"), 12), br, new PointF(1230, 1005));


                //客戶代號 -右
                _Graphics.DrawString(_OrderPool[0].客戶代號, new Font(new FontFamily("Arial"), 10), br, new PointF(1765, 835));


                //轉運單號 - 右
                //Barcode
                var bmp3 = GetCode39Bitmap(_OrderPool[0].包裹查詢號碼.轉運單號);
                _Graphics.DrawImage(bmp3, new Rectangle(1725, 910, 465, 170));
                bmp3.Dispose();
              

                //文字
                _Graphics.DrawString(_OrderPool[0].包裹查詢號碼.轉運單號, new Font(new FontFamily("Arial"), 10), br, new PointF(1849, 1080));

                //預定配達日- 右
                _Graphics.DrawString(_OrderPool[0].配送資訊.預定配達日.ToString("MM"), new Font(new FontFamily("Arial"), 10), br, new PointF(2120, 315));
                _Graphics.DrawString(_OrderPool[0].配送資訊.預定配達日.ToString("dd"), new Font(new FontFamily("Arial"), 10), br, new PointF(2160, 395));


                //指定送達時段- 右
                _Graphics.DrawString(_OrderPool[0].配送資訊.指定時段, new Font(new FontFamily("Arial"), 8), br, new PointF(2115, 540));

                //尺寸
                _Graphics.DrawString(_OrderPool[0].尺寸, new Font(new FontFamily("Arial"), 9), br, new PointF(2110, 730));


                _Graphics.Save();
            }
            #endregion

            #region 第二個

            var diff1 = 1170;
            if (_OrderPool.Count >= 2)
            {
                Brush br = new SolidBrush(Color.Black);

                _Graphics.DrawString(_OrderPool[1].包裹查詢號碼.轉運站代號, new Font(new FontFamily("Arial"), 12), br, new PointF(366, 125 + diff1));
                _Graphics.DrawString(_OrderPool[1].包裹查詢號碼.轉運單號, new Font(new FontFamily("Arial"), 18), br, new PointF(546, 115 + diff1));

                //配送資訊
                _Graphics.DrawString(_OrderPool[1].配送資訊.收貨日.ToString("yyyy-MM-dd"), new Font(new FontFamily("Arial"), 10), br, new PointF(95, 280 + diff1));
                _Graphics.DrawString(_OrderPool[1].配送資訊.預定配達日.ToString("yyyy-MM-dd"), new Font(new FontFamily("Arial"), 10), br, new PointF(355, 280 + diff1));
                _Graphics.DrawString(_OrderPool[1].配送資訊.指定時段, new Font(new FontFamily("Arial"), 9), br, new PointF(650, 280 + diff1));

                if (_OrderPool[1].配送資訊.發貨所.Length <= 5)
                {
                    var tmpPlace = _OrderPool[1].配送資訊.發貨所;
                    for (int i = 0; i < 5 - _OrderPool[1].配送資訊.發貨所.Length; i++)
                    {
                        tmpPlace = "  " + tmpPlace;
                    }
                    _Graphics.DrawString(tmpPlace, new Font(new FontFamily("Arial"), 9), br, new PointF(885, 280 + diff1));

                }

                //收件人
                _Graphics.DrawString(_OrderPool[1].收件人資訊.郵遞區號, new Font(new FontFamily("Arial"), 9), br, new PointF(150, 340 + diff1));
                _Graphics.DrawString(_OrderPool[1].收件人資訊.地址, new Font(new FontFamily("Arial"), 10), br, new PointF(315, 340 + diff1));
                _Graphics.DrawString(_OrderPool[1].收件人資訊.姓名, new Font(new FontFamily("Arial"), 10), br, new PointF(150, 410 + diff1));
                _Graphics.DrawString(_OrderPool[1].收件人資訊.電話1, new Font(new FontFamily("Arial"), 10), br, new PointF(150, 480 + diff1));
                _Graphics.DrawString(_OrderPool[1].收件人資訊.電話2, new Font(new FontFamily("Arial"), 10), br, new PointF(600, 480 + diff1));


                //寄件人
                _Graphics.DrawString(_OrderPool[1].寄件人資訊.地址, new Font(new FontFamily("Arial"), 9), br, new PointF(150, 565 + diff1));
                _Graphics.DrawString(_OrderPool[1].寄件人資訊.姓名, new Font(new FontFamily("Arial"), 8), br, new PointF(150, 620 + diff1));
                _Graphics.DrawString(_OrderPool[1].寄件人資訊.電話1, new Font(new FontFamily("Arial"), 9), br, new PointF(150, 670 + diff1));
                _Graphics.DrawString(_OrderPool[1].寄件人資訊.電話2, new Font(new FontFamily("Arial"), 9), br, new PointF(600, 670 + diff1));

                //備註
                _Graphics.DrawString(_OrderPool[1].備註, new Font(new FontFamily("Arial"), 9), br, new PointF(180, 740 + diff1));

                //品名
                _Graphics.DrawString(_OrderPool[1].品名, new Font(new FontFamily("Arial"), 9), br, new PointF(180, 810 + diff1));


                //訂單編號
                _Graphics.DrawString(_OrderPool[1].訂單編號, new Font(new FontFamily("Arial"), 10), br, new PointF(260, 890 + diff1));


                //客戶代號
                //BarCode
                var bmp1 = GetCode39Bitmap(_OrderPool[1].客戶代號);
                _Graphics.DrawImage(bmp1, new Rectangle(205, 990 + diff1, 520, 50));

                //代號
                _Graphics.DrawString(_OrderPool[1].客戶代號, new Font(new FontFamily("Arial"), 10), br, new PointF(350, 1050 + diff1));

                //金額
                _Graphics.DrawString(_OrderPool[1].代收金額.ToString("#"), new Font(new FontFamily("Arial"), 12), br, new PointF(920, 1010 + diff1));

                //郵號條碼
                //Barcode
                var bmp2 = GetCode39Bitmap("+" + _OrderPool[1].郵號條碼.Replace("-", ""));
                _Graphics.DrawImage(bmp2, new Rectangle(1195, 60+diff1, 510, 150));

                //Char
                _Graphics.DrawString(_OrderPool[1].郵號條碼, new Font(new FontFamily("Arial"), 28), br, new PointF(1695, 75+diff1));


                //收件人- 右
                _Graphics.DrawString(_OrderPool[1].收件人資訊.地址, new Font(new FontFamily("Arial"), 9), br, new PointF(1260, 260 + diff1));
                _Graphics.DrawString(_OrderPool[1].收件人資訊.姓名, new Font(new FontFamily("Arial"), 10), br, new PointF(1260, 350 + diff1));
                _Graphics.DrawString(_OrderPool[1].收件人資訊.電話1, new Font(new FontFamily("Arial"), 10), br, new PointF(1260, 410 + diff1));
                _Graphics.DrawString(_OrderPool[1].收件人資訊.電話2, new Font(new FontFamily("Arial"), 10), br, new PointF(1660, 410 + diff1));

                //寄件人- 右
                _Graphics.DrawString(_OrderPool[1].寄件人資訊.地址, new Font(new FontFamily("Arial"), 9), br, new PointF(1260, 490 + diff1));
                _Graphics.DrawString(_OrderPool[1].寄件人資訊.姓名, new Font(new FontFamily("Arial"), 9), br, new PointF(1260, 540 + diff1));
                _Graphics.DrawString(_OrderPool[1].寄件人資訊.電話1, new Font(new FontFamily("Arial"), 10), br, new PointF(1260, 600 + diff1));
                _Graphics.DrawString(_OrderPool[1].寄件人資訊.電話2, new Font(new FontFamily("Arial"), 10), br, new PointF(1660, 600 + diff1));

                //備註- 右
                _Graphics.DrawString(_OrderPool[1].備註, new Font(new FontFamily("Arial"), 9), br, new PointF(1315, 660 + diff1));

                //品名- 右
                _Graphics.DrawString(_OrderPool[1].品名, new Font(new FontFamily("Arial"), 8), br, new PointF(1315, 715 + diff1));


                //訂單編號- 右
                _Graphics.DrawString(_OrderPool[1].訂單編號, new Font(new FontFamily("Arial"), 8), br, new PointF(1420, 780 + diff1));


                //收貨日- 右
                _Graphics.DrawString(_OrderPool[1].配送資訊.收貨日.ToString("yyyy-MM-dd"), new Font(new FontFamily("Arial"), 10), br, new PointF(1320, 835 + diff1));

                //金額- 右
                _Graphics.DrawString(_OrderPool[1].代收金額.ToString("#"), new Font(new FontFamily("Arial"), 12), br, new PointF(1230, 1005 + diff1));


                //客戶代號 -右
                _Graphics.DrawString(_OrderPool[1].客戶代號, new Font(new FontFamily("Arial"), 10), br, new PointF(1765, 835 + diff1));


                //轉運單號 - 右
                //Barcode
                //轉運單號 - 右
                //Barcode
                var bmp3 = GetCode39Bitmap(_OrderPool[1].包裹查詢號碼.轉運單號);
                _Graphics.DrawImage(bmp3, new Rectangle(1725, 910+diff1, 465, 170));
                bmp3.Dispose();

                //文字
                _Graphics.DrawString(_OrderPool[1].包裹查詢號碼.轉運單號, new Font(new FontFamily("Arial"), 10), br, new PointF(1849, 1080 + diff1));

                //預定配達日- 右
                _Graphics.DrawString(_OrderPool[1].配送資訊.預定配達日.ToString("MM"), new Font(new FontFamily("Arial"), 10), br, new PointF(2120, 315 + diff1));
                _Graphics.DrawString(_OrderPool[1].配送資訊.預定配達日.ToString("dd"), new Font(new FontFamily("Arial"), 10), br, new PointF(2160, 395 + diff1));


                //指定送達時段- 右
                _Graphics.DrawString(_OrderPool[1].配送資訊.指定時段, new Font(new FontFamily("Arial"), 8), br, new PointF(2115, 540 + diff1));

                //尺寸
                _Graphics.DrawString(_OrderPool[1].尺寸, new Font(new FontFamily("Arial"), 9), br, new PointF(2110, 730 + diff1));


                _Graphics.Save();
            }
            #endregion

            #region 第三個

            var diff2 = 2340;
            if (_OrderPool.Count >= 3)
            {
                Brush br = new SolidBrush(Color.Black);

                _Graphics.DrawString(_OrderPool[2].包裹查詢號碼.轉運站代號, new Font(new FontFamily("Arial"), 12), br, new PointF(366, 125 + diff2));
                _Graphics.DrawString(_OrderPool[2].包裹查詢號碼.轉運單號, new Font(new FontFamily("Arial"), 18), br, new PointF(546, 115 + diff2));

                //配送資訊
                _Graphics.DrawString(_OrderPool[2].配送資訊.收貨日.ToString("yyyy-MM-dd"), new Font(new FontFamily("Arial"), 10), br, new PointF(95, 280 + diff2));
                _Graphics.DrawString(_OrderPool[2].配送資訊.預定配達日.ToString("yyyy-MM-dd"), new Font(new FontFamily("Arial"), 10), br, new PointF(355, 280 + diff2));
                _Graphics.DrawString(_OrderPool[2].配送資訊.指定時段, new Font(new FontFamily("Arial"), 9), br, new PointF(650, 280 + diff2));

                if (_OrderPool[2].配送資訊.發貨所.Length <= 5)
                {
                    var tmpPlace = _OrderPool[2].配送資訊.發貨所;
                    for (int i = 0; i < 5 - _OrderPool[2].配送資訊.發貨所.Length; i++)
                    {
                        tmpPlace = "  " + tmpPlace;
                    }
                    _Graphics.DrawString(tmpPlace, new Font(new FontFamily("Arial"), 9), br, new PointF(885, 280 + diff2));

                }

                //收件人
                _Graphics.DrawString(_OrderPool[2].收件人資訊.郵遞區號, new Font(new FontFamily("Arial"), 9), br, new PointF(150, 340 + diff2));
                _Graphics.DrawString(_OrderPool[2].收件人資訊.地址, new Font(new FontFamily("Arial"), 10), br, new PointF(315, 340 + diff2));
                _Graphics.DrawString(_OrderPool[2].收件人資訊.姓名, new Font(new FontFamily("Arial"), 10), br, new PointF(150, 410 + diff2));
                _Graphics.DrawString(_OrderPool[2].收件人資訊.電話1, new Font(new FontFamily("Arial"), 10), br, new PointF(150, 480 + diff2));
                _Graphics.DrawString(_OrderPool[2].收件人資訊.電話2, new Font(new FontFamily("Arial"), 10), br, new PointF(600, 480 + diff2));


                //寄件人
                _Graphics.DrawString(_OrderPool[2].寄件人資訊.地址, new Font(new FontFamily("Arial"), 9), br, new PointF(150, 565 + diff2));
                _Graphics.DrawString(_OrderPool[2].寄件人資訊.姓名, new Font(new FontFamily("Arial"), 8), br, new PointF(150, 620 + diff2));
                _Graphics.DrawString(_OrderPool[2].寄件人資訊.電話1, new Font(new FontFamily("Arial"), 9), br, new PointF(150, 670 + diff2));
                _Graphics.DrawString(_OrderPool[2].寄件人資訊.電話2, new Font(new FontFamily("Arial"), 9), br, new PointF(600, 670 + diff2));

                //備註
                _Graphics.DrawString(_OrderPool[2].備註, new Font(new FontFamily("Arial"), 9), br, new PointF(180, 740 + diff2));

                //品名
                _Graphics.DrawString(_OrderPool[2].品名, new Font(new FontFamily("Arial"), 9), br, new PointF(180, 810 + diff2));


                //訂單編號
                _Graphics.DrawString(_OrderPool[2].訂單編號, new Font(new FontFamily("Arial"), 10), br, new PointF(260, 890 + diff2));


                //客戶代號
                //BarCode
                var bmp1 = GetCode39Bitmap(_OrderPool[2].客戶代號);
              
                _Graphics.DrawImage(bmp1, new Rectangle(205, 990 + diff2, 520, 50));
                

                //代號
                _Graphics.DrawString(_OrderPool[2].客戶代號, new Font(new FontFamily("Arial"), 10), br, new PointF(350, 1050 + diff2));

                //金額
                _Graphics.DrawString(_OrderPool[2].代收金額.ToString("#"), new Font(new FontFamily("Arial"), 12), br, new PointF(920, 1010 + diff2));

                //郵號條碼
                //Barcode
                var bmp2 = GetCode39Bitmap("+" + _OrderPool[2].郵號條碼.Replace("-", ""));
                
                _Graphics.DrawImage(bmp2, new Rectangle(1195, 60 + diff2, 510, 150));

                //Char
                _Graphics.DrawString(_OrderPool[2].郵號條碼, new Font(new FontFamily("Arial"), 28), br, new PointF(1695, 75 + diff2));


                //收件人- 右
                _Graphics.DrawString(_OrderPool[2].收件人資訊.地址, new Font(new FontFamily("Arial"), 9), br, new PointF(1260, 260 + diff2));
                _Graphics.DrawString(_OrderPool[2].收件人資訊.姓名, new Font(new FontFamily("Arial"), 10), br, new PointF(1260, 350 + diff2));
                _Graphics.DrawString(_OrderPool[2].收件人資訊.電話1, new Font(new FontFamily("Arial"), 10), br, new PointF(1260, 410 + diff2));
                _Graphics.DrawString(_OrderPool[2].收件人資訊.電話2, new Font(new FontFamily("Arial"), 10), br, new PointF(1660, 410 + diff2));

                //寄件人- 右
                _Graphics.DrawString(_OrderPool[2].寄件人資訊.地址, new Font(new FontFamily("Arial"), 9), br, new PointF(1260, 490 + diff2));
                _Graphics.DrawString(_OrderPool[2].寄件人資訊.姓名, new Font(new FontFamily("Arial"), 9), br, new PointF(1260, 540 + diff2));
                _Graphics.DrawString(_OrderPool[2].寄件人資訊.電話1, new Font(new FontFamily("Arial"), 10), br, new PointF(1260, 600 + diff2));
                _Graphics.DrawString(_OrderPool[2].寄件人資訊.電話2, new Font(new FontFamily("Arial"), 10), br, new PointF(1660, 600 + diff2));

                //備註- 右
                _Graphics.DrawString(_OrderPool[2].備註, new Font(new FontFamily("Arial"), 9), br, new PointF(1315, 660 + diff2));

                //品名- 右
                _Graphics.DrawString(_OrderPool[2].品名, new Font(new FontFamily("Arial"), 8), br, new PointF(1315, 715 + diff2));


                //訂單編號- 右
                _Graphics.DrawString(_OrderPool[2].訂單編號, new Font(new FontFamily("Arial"), 8), br, new PointF(1420, 780 + diff2));


                //收貨日- 右
                _Graphics.DrawString(_OrderPool[2].配送資訊.收貨日.ToString("yyyy-MM-dd"), new Font(new FontFamily("Arial"), 10), br, new PointF(1320, 835 + diff2));

                //金額- 右
                _Graphics.DrawString(_OrderPool[2].代收金額.ToString("#"), new Font(new FontFamily("Arial"), 12), br, new PointF(1230, 1005 + diff2));


                //客戶代號 -右
                _Graphics.DrawString(_OrderPool[2].客戶代號, new Font(new FontFamily("Arial"), 10), br, new PointF(1765, 835 + diff2));


                //轉運單號 - 右
                //Barcode
                //轉運單號 - 右
                //Barcode
                var bmp3 = GetCode39Bitmap(_OrderPool[2].包裹查詢號碼.轉運單號);
                _Graphics.DrawImage(bmp3, new Rectangle(1725, 910+diff2, 465, 170));
                bmp3.Dispose();

                //文字
                _Graphics.DrawString(_OrderPool[2].包裹查詢號碼.轉運單號, new Font(new FontFamily("Arial"), 10), br, new PointF(1849, 1080 + diff2));

                //預定配達日- 右
                _Graphics.DrawString(_OrderPool[2].配送資訊.預定配達日.ToString("MM"), new Font(new FontFamily("Arial"), 10), br, new PointF(2120, 315 + diff2));
                _Graphics.DrawString(_OrderPool[2].配送資訊.預定配達日.ToString("dd"), new Font(new FontFamily("Arial"), 10), br, new PointF(2160, 395 + diff2));


                //指定送達時段- 右
                _Graphics.DrawString(_OrderPool[2].配送資訊.指定時段, new Font(new FontFamily("Arial"), 8), br, new PointF(2115, 540 + diff2));

                //尺寸
                _Graphics.DrawString(_OrderPool[2].尺寸, new Font(new FontFamily("Arial"), 9), br, new PointF(2110, 730 + diff2));


                _Graphics.Save();
            }
            #endregion

            ExportToFile(targetFilePath, rBitmap);

        }


      
        private static void ExportToFile(string path, Bitmap bitmap)
        {
            using (MemoryStream oMemoryStream = new MemoryStream())
            {

                //儲存圖片到 MemoryStream 物件，並且指定儲存影像之格式
                //Save image to MemoryStream  and set it to jpeg format.
                bitmap.Save(oMemoryStream, ImageFormat.Jpeg);
                //設定資料流位置
                //Set stream position start from zero
                oMemoryStream.Position = 0;
                //設定 buffer 長度
                //Set buffer length
                var data = new byte[oMemoryStream.Length];
                //將資料寫入 buffer
                //Wrire data to buffer
                oMemoryStream.Read(data, 0, Convert.ToInt32(oMemoryStream.Length));
                //將所有緩衝區的資料寫入資料流
                //Flush memory.
                oMemoryStream.Flush();
                File.WriteAllBytes(path, data);
            }

        }
    }
}
