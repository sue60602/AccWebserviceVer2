//------------------------------------------------------------------------------
// <auto-generated>
//    這個程式碼是由範本產生。
//
//    對這個檔案進行手動變更可能導致您的應用程式產生未預期的行為。
//    如果重新產生程式碼，將會覆寫對這個檔案的手動變更。
// </auto-generated>
//------------------------------------------------------------------------------

namespace AccWebService.EF
{
    using System;
    using System.Collections.Generic;
    
    public partial class LongTermDebt
    {
        public string FundNo { get; set; }
        public string AccYear { get; set; }
        public int OrderNo { get; set; }
        public string ItemNo { get; set; }
        public string DebtName { get; set; }
        public string DebtUnit { get; set; }
        public Nullable<System.DateTime> DebtDate { get; set; }
        public Nullable<System.DateTime> ClearDate { get; set; }
        public Nullable<long> DebtAmount { get; set; }
        public Nullable<long> ReturnAmount { get; set; }
        public Nullable<long> ShortAmount { get; set; }
        public Nullable<long> InterestAmount { get; set; }
        public string Remark { get; set; }
    }
}