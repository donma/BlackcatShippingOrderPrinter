using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace BlackcatShippingOrderPrintHelper.Test
{
    public partial class index : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
       
        }

     
     

        protected void Button1_Click(object sender, EventArgs e)
        {
            BlackcatShippingOrderPrintHelper.Agent.CearAllOrderInfo();
            //載入表格圖片
            BlackcatShippingOrderPrintHelper.Agent.LoadSampleImage(AppDomain.CurrentDomain.BaseDirectory + "sample.jpg");


            //第一聯的資料
            BlackcatShippingOrderPrintHelper.OrderInfo o1 = new OrderInfo();
            o1.包裹查詢號碼.轉運站代號 = "32";
            o1.包裹查詢號碼.轉運單號 = "7770000011";

            o1.配送資訊.指定時段 = "17-20時";
            o1.配送資訊.收貨日 = new DateTime(2016, 12, 21);
            o1.配送資訊.預定配達日 = new DateTime(2016, 12, 31);
            o1.配送資訊.發貨所 = "南港所";

            o1.收件人資訊.郵遞區號 = "12345";
            o1.收件人資訊.地址 = "台北市松江路yyy號yyy樓";
            o1.收件人資訊.姓名 = "許當麻1";
            o1.收件人資訊.電話1 = "02-25711956";
            o1.收件人資訊.電話2 = "0975543220";


            o1.寄件人資訊.地址 = "屏東縣恆春鎮海角六號";
            o1.寄件人資訊.姓名 = "潘蜥";
            o1.寄件人資訊.電話1 = "02-25711956";
            o1.寄件人資訊.電話2 = "0975543220";

            o1.備註 = "不可以自己寄給自己";
            o1.品名 = "北海道豬肉螃蟹鍋 x (2) 件";
            o1.訂單編號 = "T1213121543220";

            o1.客戶代號 = "273637101";
            o1.代收金額 = 199;
            o1.郵號條碼 = "11-115-34";
            o1.尺寸 = "150 cm";
            BlackcatShippingOrderPrintHelper.Agent.AddOrderInfo(o1);



            //第二聯的資料
            BlackcatShippingOrderPrintHelper.OrderInfo o2 = new OrderInfo();
            o2.包裹查詢號碼.轉運站代號 = "32";
            o2.包裹查詢號碼.轉運單號 = "7770000011";

            o2.配送資訊.指定時段 = "17-20時";
            o2.配送資訊.收貨日 = new DateTime(2016, 12, 21);
            o2.配送資訊.預定配達日 = new DateTime(2016, 12, 31);
            o2.配送資訊.發貨所 = "南港所";

            o2.收件人資訊.郵遞區號 = "12345";
            o2.收件人資訊.地址 = "台北市松江路xxx號xxx樓";
            o2.收件人資訊.姓名 = "許當麻2";
            o2.收件人資訊.電話1 = "02-25711956";
            o2.收件人資訊.電話2 = "0975543220";


            o2.寄件人資訊.地址 = "屏東縣恆春鎮海角六號";
            o2.寄件人資訊.姓名 = "潘蜥";
            o2.寄件人資訊.電話1 = "02-25711956";
            o2.寄件人資訊.電話2 = "0975543220";

            o2.備註 = "不可以自己寄給自己";
            o2.品名 = "北海道豬肉螃蟹鍋 x (2) 件";
            o2.訂單編號 = "T1213121543220";

            o2.客戶代號 = "273637101";
            o2.代收金額 = 199;
            o2.郵號條碼 = "32-331-02";
            o2.尺寸 = "150 cm";
            BlackcatShippingOrderPrintHelper.Agent.AddOrderInfo(o2);



            //第三聯的資料
            BlackcatShippingOrderPrintHelper.OrderInfo o3 = new OrderInfo();
            o3.包裹查詢號碼.轉運站代號 = "32";
            o3.包裹查詢號碼.轉運單號 = "7770000011";

            o3.配送資訊.指定時段 = "17-20時";
            o3.配送資訊.收貨日 = new DateTime(2016, 12, 21);
            o3.配送資訊.預定配達日 = new DateTime(2016, 12, 31);
            o3.配送資訊.發貨所 = "南港所";

            o3.收件人資訊.郵遞區號 = "12345";
            o3.收件人資訊.地址 = "台北市松江路zzz號zzz樓";
            o3.收件人資訊.姓名 = "許當麻3";
            o3.收件人資訊.電話1 = "02-25711956";
            o3.收件人資訊.電話2 = "0975543220";


            o3.寄件人資訊.地址 = "屏東縣恆春鎮海角六號";
            o3.寄件人資訊.姓名 = "潘蜥";
            o3.寄件人資訊.電話1 = "02-25711956";
            o3.寄件人資訊.電話2 = "0975543220";

            o3.備註 = "不可以自己寄給自己";
            o3.品名 = "北海道豬肉螃蟹鍋 x (2) 件";
            o3.訂單編號 = "T1213121543220";

            o3.客戶代號 = "273637101";
            o3.代收金額 = 199;
            o3.郵號條碼 = "32-331-02";
            o3.尺寸 = "150 cm";
            BlackcatShippingOrderPrintHelper.Agent.AddOrderInfo(o3);

            BlackcatShippingOrderPrintHelper.Agent.ExportToJPG(AppDomain.CurrentDomain.BaseDirectory + "test.jpg");

            Image1.ImageUrl = "test.jpg?" + Guid.NewGuid().ToString("N");

            //回收
            BlackcatShippingOrderPrintHelper.Agent.Dispose();



        }
    }
}