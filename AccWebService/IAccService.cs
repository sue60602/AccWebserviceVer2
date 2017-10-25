using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.ServiceModel;
using System.Text;

namespace AccWebService
{
    // 注意: 您可以使用 [重構] 功能表上的 [重新命名] 命令同時變更程式碼和組態檔中的介面名稱 "IAccService"。
    [ServiceContract]
    public interface IAccService
    {
        //輸入條碼,回傳傳票底稿
        [OperationContract]
        string GetVw_GBCVisaDetail(string fundNo, string acmWordNum);

        //傳票號碼回填
        [OperationContract]
        string FillVouNo(string fundNo, string acmWordNum, string vouNoJSON);

        //取年度
        [OperationContract]
        List<string> GetYear(string fundNo);

        //取動支號
        [OperationContract]
        List<string> GetAcmWordNum(string fundNo, string accYear);

        //取種類
        [OperationContract]
        List<string> GetAccKind(string fundNo, string accYear, string acmWordNum);

        //取次數
        [OperationContract]
        List<string> GetAccCount(string fundNo, string accYear, string acmWordNum, string accKind);

        //取明細號
        [OperationContract]
        List<string> GetAccDetail(string fundNo, string accYear, string acmWordNum, string accKind, string accCount);

        //依據KEY值取View
        [OperationContract]
        string GetByPrimaryKey(string fundNo, string accYear, string acmWordNum, string accKind, string accCount, string accDetail);

        //取估列List
        [OperationContract]
        List<string> GetByKind(string fundNo, string accYear, string accKind, string batch);
    }
}
