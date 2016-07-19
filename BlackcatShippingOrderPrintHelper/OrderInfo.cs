using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace BlackcatShippingOrderPrintHelper
{
    public class 包裹查詢號碼
    {
        public string 轉運站代號 { get; set; }
        public string 轉運單號 { get; set; }
    }

    public class 配送資訊
    {
        public DateTime 收貨日 { get; set; }
        public DateTime 預定配達日 { get; set; }
        public string 指定時段 { get; set; }
        public string 發貨所 { get; set; }
    }

    public class 收件人資訊
    {
        public string 郵遞區號 { get; set; }
        public string 地址 { get; set; }
        public string 姓名 { get; set; }
        public string 電話1 { get; set; }
        public string 電話2 { get; set; }
    }

    public class 寄件人資訊
    {
       // public string 郵遞區號 { get; set; }
        public string 地址 { get; set; }
        public string 姓名 { get; set; }
        public string 電話1 { get; set; }
        public string 電話2 { get; set; }
    }

    public class OrderInfo
    {
        public 包裹查詢號碼 包裹查詢號碼 { get; set; }

        public 配送資訊 配送資訊 { get; set; }


        public 收件人資訊 收件人資訊 { get; set; }

        public 寄件人資訊 寄件人資訊 { get; set; }

        public string 備註 { get; set; }

        public string 品名 { get; set; }

        public string 訂單編號 { get; set; }

        public string 客戶代號 { get; set; }

        public decimal 代收金額 { get; set; }

        public string 尺寸 { get; set; }

        /// <summary>
        /// 請含 - 符號
        /// ex: 32-331-02
        /// </summary>
        public string 郵號條碼 { get; set; }

        public OrderInfo()
        {
            包裹查詢號碼 = new 包裹查詢號碼();
            配送資訊 = new 配送資訊();
            收件人資訊 = new 收件人資訊();
            寄件人資訊 = new 寄件人資訊();
        }
    }
}
