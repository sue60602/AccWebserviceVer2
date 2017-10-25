using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace AccWebService.Model
{
    public class AccScriptJSON
    {
        public class 最外層
        {
            public string 基金代碼 { get; set; }
            public string 年度 { get; set; }
            public string 動支編號 { get; set; }
            public string 種類 { get; set; }
            public string 次別 { get; set; }
            public string 明細號 { get; set; }
            public List<傳票內容> 傳票內容 { get; set; }
        }

        public class 傳票內容
        {
            public 傳票主檔 傳票主檔 { get; set; }
            public List<傳票明細> 傳票明細 { get; set; }
            public List<傳票受款人> 傳票受款人 { get; set; }
        }

        public class 傳票主檔
        {
            public string 傳票種類 { get; set; }
            public string 製票日期 { get; set; }
            public string 主摘要 { get; set; }
            public string 交付方式 { get; set; }
        }

        public class 傳票受款人
        {
            public string 統一編號 { get; set; }
            public string 受款人名稱 { get; set; }
            public string 地址 { get; set; }
            public double 實付金額 { get; set; }
            public string 銀行代號 { get; set; }
            public string 銀行名稱 { get; set; }
            public string 銀行帳號 { get; set; }
            public string 帳戶名稱 { get; set; }
        }

        public class 傳票明細
        {
            public string 借貸別 { get; set; }
            public string 科目代號 { get; set; }
            public string 科目名稱 { get; set; }
            public string 摘要 { get; set; }
            public double 金額 { get; set; }
            public string 計畫代碼 { get; set; }
            public string 用途別代碼 { get; set; }
            public string 沖轉字號 { get; set; }
            public string 對象代碼 { get; set; }
            public string 對象說明 { get; set; }
            public string 明細號 { get; set; }
        }
    }
}