using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace AccWebService.Model
{
    public class FillVouScriptForHPA
    {
        public string 傳票年度 { get; set; }
        public string 傳票號 { get; set; }
        public string 製票日期 { get; set; }
        public string 明細號 { get; set; }
        public string 傳票明細號 { get; set; }
    }
}