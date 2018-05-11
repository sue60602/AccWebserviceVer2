using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace AccWebService.Model
{
    public class Sp_GBCVisaDetail
    {
        public string F_用途別代碼 = string.Empty;
        public string F_批號 = string.Empty;
        public string F_受款人 = string.Empty;
        public string F_受款人編號 = string.Empty;
        public string F_是否核定 = string.Empty;
        public string F_科室代碼 = string.Empty;
        public string F_計畫代碼 = string.Empty;
        public string F_原動支編號 = string.Empty;
        public string F_核定日期 = string.Empty;
        public double F_核定金額;
        public string F_動支金額 = string.Empty;
        public string F_摘要 = string.Empty;
        public string F_製票日 = string.Empty;
        public string PK_次別 = string.Empty;
        public string PK_明細號 = string.Empty;
        public string PK_動支編號 = string.Empty;
        public string PK_會計年度 = string.Empty;
        public string PK_種類 = string.Empty;
        public string 基金代碼 = string.Empty;
        public string BarCode = string.Empty;
    }
}