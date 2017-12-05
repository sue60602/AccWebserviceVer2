using AccWebService.EF;
using AccWebService.Model;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Configuration;
using System.Web.Services;
using static AccWebService.Model.AccScriptJSON;
using System.Security.Cryptography.X509Certificates;
using System.Net;
using System.Net.Security;

namespace AccWebService
{
    /// <summary>
    ///Acc_WebService 的摘要描述
    /// </summary>
    [WebService(Namespace = "http://tempuri.org/")]
    [WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
    [System.ComponentModel.ToolboxItem(false)]
    // 若要允許使用 ASP.NET AJAX 從指令碼呼叫此 Web 服務，請取消註解下列一行。
    // [System.Web.Script.Services.ScriptService]
    public class Acc_WebService : System.Web.Services.WebService
    {
        public static bool ValidateServerCertificate(Object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors)
        {
            return true;
        }

        [WebMethod]
        /// <summary>
        /// 依據條碼，回傳傳票明細JSON
        /// </summary>
        /// <param name="fundNo"></param>
        /// <param name="acmWordNum"></param>
        /// <returns></returns>
        public string GetVw_GBCVisaDetail(string fundNo, string acmWordNum)
        {
            ServicePointManager.ServerCertificateValidationCallback = new RemoteCertificateValidationCallback(ValidateServerCertificate);
            GBCVisaDetailAbateDetailDAO dao = new GBCVisaDetailAbateDetailDAO();
            GBCJSONRecordDAO jsonDAO = new GBCJSONRecordDAO();
            VouDetailDAO vouDetailDAO = new VouDetailDAO();
            VouMainDAO vouMainDAO = new VouMainDAO();

            Vw_GBCVisaDetail vw_GBCVisaDetail = new Vw_GBCVisaDetail();
            List<Vw_GBCVisaDetail> vwList = new List<Vw_GBCVisaDetail>();
            string JSON1 = null; //宣告回傳JSON1

            //JSON底稿定義
            //List<最外層> vouTopList = new List<最外層>(); //可能有多筆明細號,所以用集合包起來
            最外層 vouTop = new 最外層(); //宣告輸出JSON格式
            最外層 vouTop2 = null; //宣告輸出JSON2格式

            List<傳票明細> vouDtlList = new List<傳票明細>();
            List<傳票受款人> vouPayList = new List<傳票受款人>();
            List<傳票內容> vouCollectionList = new List<傳票內容>();

            //如果會開到第二種傳票時須額外再定義一次
            List<傳票明細> vouDtlList2 = new List<傳票明細>();
            List<傳票受款人> vouPayList2 = new List<傳票受款人>();
            List<傳票內容> vouCollectionList2 = new List<傳票內容>();

            傳票主檔 vouMain = new 傳票主檔();
            傳票內容 vouCollection = new 傳票內容();

            //宣告接收從預控端取得之JSON字串
            string JSONReturn = "";

            //先判斷基金代號
            if (fundNo == "010")//醫發服務參考
            {
                GBCWebService.GBCWebService ws = new GBCWebService.GBCWebService();
                JSONReturn = ws.GetVw_GBCVisaDetailJSON(acmWordNum); //呼叫預控的服務,取得此動支編號的view資料
            }
            else if (fundNo == "040")//菸害****尚未加入服務參考****
            {

            }
            else if (fundNo == "090")//家防服務參考
            {
                DVGBCWebService.GBCWebService ws = new DVGBCWebService.GBCWebService();
                JSONReturn = ws.GetVw_GBCVisaDetailJSON(acmWordNum); //呼叫預控的服務,取得此動支編號的view資料
            }
            else if (fundNo == "100")//長照
            {
                LCGBCWebService.GBCWebService ws = new LCGBCWebService.GBCWebService();
                JSONReturn = ws.GetVw_GBCVisaDetailJSON(acmWordNum); //呼叫預控的服務,取得此動支編號的view資料
            }
            else if (fundNo == "110")//生產
            {
                BAGBCWebService.GBCWebService ws = new BAGBCWebService.GBCWebService();
                JSONReturn = ws.GetVw_GBCVisaDetailJSON(acmWordNum); //呼叫預控的服務,取得此動支編號的view資料
            }
            else
            {
                return "基金代號有誤! 號碼為: " + fundNo;
            }

            try
            {
                vwList = JsonConvert.DeserializeObject<List<Vw_GBCVisaDetail>>(JSONReturn);  //反序列化JSON               
            }
            catch (Exception e)
            {
                return JSONReturn;
            }

            var accKind = (from acckind in vwList select acckind.PK_種類).First();//取種類代碼
            var accSumMoney = (from money in vwList select money.F_核定金額).Sum();//取核銷總額
            var PayVouKind = (from voukind in vwList select voukind.基金代碼).First();//取票種類(主要是用來區分付款憑單(vouKind=5)或是支出傳票(vouKind=2))

            int isPrePay = 0; //有無預付
            int isLog = 0; //有無預付

            if (PayVouKind == "090") //如果是家防基金,使用憑單
            {
                PayVouKind = "5";
            }
            else
            {
                PayVouKind = "2";
            }

            /*
             * 一共有六種狀態,分別為:
             * 1.預付    、2.核銷    、3.估列、
             * 4.估列收回、5.預撥收回、6.核銷收回
             */
            #region 預付
            if ("預付".Equals(accKind))
            {
                foreach (var vwListItem in vwList)
                {
                    vw_GBCVisaDetail.基金代碼 = vwListItem.基金代碼;
                    vw_GBCVisaDetail.PK_會計年度 = vwListItem.PK_會計年度;
                    vw_GBCVisaDetail.PK_動支編號 = vwListItem.PK_動支編號;
                    vw_GBCVisaDetail.PK_種類 = vwListItem.PK_種類;
                    vw_GBCVisaDetail.PK_次別 = vwListItem.PK_次別;
                    vw_GBCVisaDetail.PK_明細號 = vwListItem.PK_明細號;
                    vw_GBCVisaDetail.F_科室代碼 = vwListItem.F_科室代碼;
                    vw_GBCVisaDetail.F_用途別代碼 = vwListItem.F_用途別代碼;
                    vw_GBCVisaDetail.F_計畫代碼 = vwListItem.F_計畫代碼;
                    vw_GBCVisaDetail.F_動支金額 = vwListItem.F_動支金額;
                    vw_GBCVisaDetail.F_製票日 = vwListItem.F_製票日;
                    vw_GBCVisaDetail.F_是否核定 = vwListItem.F_是否核定;
                    vw_GBCVisaDetail.F_核定金額 = vwListItem.F_核定金額;
                    vw_GBCVisaDetail.F_核定日期 = vwListItem.F_核定日期;
                    vw_GBCVisaDetail.F_摘要 = vwListItem.F_摘要;
                    vw_GBCVisaDetail.F_受款人 = vwListItem.F_受款人;
                    vw_GBCVisaDetail.F_受款人編號 = vwListItem.F_受款人編號;
                    vw_GBCVisaDetail.F_原動支編號 = vwListItem.F_原動支編號;
                    vw_GBCVisaDetail.F_批號 = vwListItem.F_批號;

                    try
                    {
                        isLog = dao.FindLog(x => x.基金代碼 == vw_GBCVisaDetail.基金代碼 && x.PK_會計年度 == vw_GBCVisaDetail.PK_會計年度 && x.PK_動支編號 == vw_GBCVisaDetail.PK_動支編號 && x.PK_種類 == vw_GBCVisaDetail.PK_種類 && x.PK_次別 == vw_GBCVisaDetail.PK_次別 && x.PK_明細號 == vw_GBCVisaDetail.PK_明細號);
                        string isPass = jsonDAO.IsPass(x => x.基金代碼 == vw_GBCVisaDetail.基金代碼 && x.PFK_會計年度 == vw_GBCVisaDetail.PK_會計年度 && x.PFK_動支編號 == vw_GBCVisaDetail.PK_動支編號 && x.PFK_種類 == vw_GBCVisaDetail.PK_種類 && x.PFK_次別 == vw_GBCVisaDetail.PK_次別);

                        if ((isLog > 0) && isPass.Equals("1"))
                        {
                            return "此筆資料已轉入過,並且結案。";
                        }
                        else if (((isLog > 0) && isPass.Equals("0")))
                        {
                            dao.Update(vw_GBCVisaDetail);
                            jsonDAO.DeleteJsonRecord1(vw_GBCVisaDetail);
                        }
                        else
                        {
                            dao.Insert(vw_GBCVisaDetail);
                        }
                    }
                    catch (Exception e)
                    {
                        return e.Message;
                    }

                    傳票明細 vouDtl_D = new 傳票明細()
                    {
                        借貸別 = "借",
                        科目代號 = "1154",
                        科目名稱 = "預付費用",
                        摘要 = vw_GBCVisaDetail.F_摘要,
                        金額 = vw_GBCVisaDetail.F_核定金額,
                        計畫代碼 = vw_GBCVisaDetail.F_計畫代碼,
                        用途別代碼 = vw_GBCVisaDetail.F_用途別代碼,
                        沖轉字號 = "",
                        對象代碼 = vw_GBCVisaDetail.F_受款人編號,
                        對象說明 = vw_GBCVisaDetail.F_受款人,
                        明細號 = vw_GBCVisaDetail.PK_明細號
                    };
                    vouDtlList.Add(vouDtl_D);
                    傳票受款人 vouPay = new 傳票受款人()
                    {
                        //統一編號 = vw_GBCVisaDetail.F_受款人編號,
                        //受款人名稱 = vw_GBCVisaDetail.F_受款人,
                        //地址 = "",
                        //實付金額 = vw_GBCVisaDetail.F_核定金額,
                        //銀行代號 = "",
                        //銀行名稱 = "",
                        //銀行帳號 = "",
                        //帳戶名稱 = ""
                        統一編號 = "",
                        受款人名稱 = "",
                        地址 = "",
                        實付金額 = 0,
                        銀行代號 = "",
                        銀行名稱 = "",
                        銀行帳號 = "",
                        帳戶名稱 = ""
                    };
                    vouPayList.Add(vouPay);

                    //填傳票明細號1
                    //dao.FillVouDtl1(vw_GBCVisaDetail.基金代碼, vw_GBCVisaDetail.PK_會計年度, vw_GBCVisaDetail.PK_動支編號, vw_GBCVisaDetail.PK_種類, vw_GBCVisaDetail.PK_次別, vw_GBCVisaDetail.PK_明細號, vouDtlList.Count);
                }
                //重新處理受款人清單,如果有重複受款人名稱,則金額加總
                var vouPayGroup = from xxx in vouPayList
                                  group xxx by new { xxx.統一編號, xxx.受款人名稱, xxx.地址, xxx.銀行代號, xxx.銀行名稱, xxx.銀行帳號, xxx.帳戶名稱 } into g
                                  select new { 統一編號 = g.Key.統一編號, 受款人名稱 = g.Key.受款人名稱, 地址 = g.Key.地址, 銀行代號 = g.Key.銀行代號, 銀行名稱 = g.Key.銀行名稱, 銀行帳號 = g.Key.銀行帳號, 帳戶名稱 = g.Key.帳戶名稱, 實付金額 = g.Sum(xxx => xxx.實付金額) };
                //vouPayList = new List<傳票受款人>();
                //foreach (var vouPayGroupItem in vouPayGroup)
                //{
                //    傳票受款人 vouPay = new 傳票受款人();
                //    vouPay.統一編號 = vouPayGroupItem.統一編號;
                //    vouPay.受款人名稱 = vouPayGroupItem.受款人名稱;
                //    vouPay.地址 = vouPayGroupItem.地址;
                //    vouPay.實付金額 = vouPayGroupItem.實付金額;
                //    vouPay.銀行代號 = vouPayGroupItem.銀行代號;
                //    vouPay.銀行名稱 = vouPayGroupItem.銀行名稱;
                //    vouPay.銀行帳號 = vouPayGroupItem.銀行帳號;
                //    vouPay.帳戶名稱 = vouPayGroupItem.帳戶名稱;
                //    vouPayList.Add(vouPay);
                //}

                vouMain.傳票種類 = PayVouKind;
                vouMain.製票日期 = "";
                vouMain.主摘要 = vw_GBCVisaDetail.F_摘要;
                vouMain.交付方式 = "1";

                傳票明細 vouDtl_C = new 傳票明細()
                {
                    借貸別 = "貸",
                    科目代號 = "1112",
                    科目名稱 = "銀行存款",
                    摘要 = vw_GBCVisaDetail.F_摘要,
                    金額 = accSumMoney,
                    計畫代碼 = "",
                    用途別代碼 = "",
                    沖轉字號 = "",
                    對象代碼 = "",
                    對象說明 = ""
                };
                vouDtlList.Add(vouDtl_C);

                vouCollection.傳票主檔 = vouMain;
                vouCollection.傳票明細 = vouDtlList;
                vouCollection.傳票受款人 = vouPayList;

                vouCollectionList.Add(vouCollection);

                vouTop.基金代碼 = vw_GBCVisaDetail.基金代碼;
                vouTop.年度 = vw_GBCVisaDetail.PK_會計年度;
                vouTop.動支編號 = vw_GBCVisaDetail.PK_動支編號;
                vouTop.種類 = vw_GBCVisaDetail.PK_種類;
                vouTop.次別 = vw_GBCVisaDetail.PK_次別;
                vouTop.明細號 = vw_GBCVisaDetail.PK_明細號;
                vouTop.傳票內容 = vouCollectionList;
            }
            #endregion

            #region 核銷
            if ("核銷".Equals(accKind))
            {
                bool isCash = false;
                bool isPre = false;
                bool isEst = false;
                bool isFee = false;
                int detailCount = 0;

                foreach (var vwListItem in vwList)
                {
                    vw_GBCVisaDetail.基金代碼 = vwListItem.基金代碼;
                    vw_GBCVisaDetail.PK_會計年度 = vwListItem.PK_會計年度;
                    vw_GBCVisaDetail.PK_動支編號 = vwListItem.PK_動支編號;
                    vw_GBCVisaDetail.PK_種類 = vwListItem.PK_種類;
                    vw_GBCVisaDetail.PK_次別 = vwListItem.PK_次別;
                    vw_GBCVisaDetail.PK_明細號 = vwListItem.PK_明細號;
                    vw_GBCVisaDetail.F_科室代碼 = vwListItem.F_科室代碼;
                    vw_GBCVisaDetail.F_用途別代碼 = vwListItem.F_用途別代碼;
                    vw_GBCVisaDetail.F_計畫代碼 = vwListItem.F_計畫代碼;
                    vw_GBCVisaDetail.F_動支金額 = vwListItem.F_動支金額;
                    vw_GBCVisaDetail.F_製票日 = vwListItem.F_製票日;
                    vw_GBCVisaDetail.F_是否核定 = vwListItem.F_是否核定;
                    vw_GBCVisaDetail.F_核定金額 = vwListItem.F_核定金額;
                    vw_GBCVisaDetail.F_核定日期 = vwListItem.F_核定日期;
                    vw_GBCVisaDetail.F_摘要 = vwListItem.F_摘要;
                    vw_GBCVisaDetail.F_受款人 = vwListItem.F_受款人;
                    vw_GBCVisaDetail.F_受款人編號 = vwListItem.F_受款人編號;
                    vw_GBCVisaDetail.F_原動支編號 = vwListItem.F_原動支編號;
                    vw_GBCVisaDetail.F_批號 = vwListItem.F_批號;
                    vw_GBCVisaDetail.實支 = vwListItem.實支;
                    vw_GBCVisaDetail.預付轉正 = vwListItem.預付轉正;
                    vw_GBCVisaDetail.沖抵估列 = vwListItem.沖抵估列;
                    vw_GBCVisaDetail.費用 = vwListItem.費用;

                    //傳票號1有值,而且尚未結案時,直接回傳JSON2
                    string isVouNo1 = dao.IsVouNo1(vwListItem.基金代碼, vwListItem.PK_會計年度, vwListItem.PK_動支編號, vwListItem.PK_種類, vwListItem.PK_次別, vwListItem.PK_明細號);
                    string isPass = jsonDAO.IsPass(x => x.基金代碼 == vw_GBCVisaDetail.基金代碼 && x.PFK_會計年度 == vw_GBCVisaDetail.PK_會計年度 && x.PFK_動支編號 == vw_GBCVisaDetail.PK_動支編號 && x.PFK_種類 == vw_GBCVisaDetail.PK_種類 && x.PFK_次別 == vw_GBCVisaDetail.PK_次別);

                    if (isVouNo1 != null && isPass == "0")
                    {
                        string json2 = jsonDAO.FindJSON2(x => x.基金代碼 == vw_GBCVisaDetail.基金代碼 && x.PFK_會計年度 == vw_GBCVisaDetail.PK_會計年度 && x.PFK_動支編號 == vw_GBCVisaDetail.PK_動支編號 && x.PFK_種類 == vw_GBCVisaDetail.PK_種類 && x.PFK_次別 == vw_GBCVisaDetail.PK_次別);
                        if (json2 == "")
                        {
                            jsonDAO.UpdatePassFlg(vw_GBCVisaDetail.基金代碼, vw_GBCVisaDetail.PK_會計年度, vw_GBCVisaDetail.PK_動支編號, vw_GBCVisaDetail.PK_種類, vw_GBCVisaDetail.PK_次別);
                            return "此動支已開立完畢， 傳票號碼為： " + isVouNo1;
                        }
                        return json2;
                    }

                    if (vw_GBCVisaDetail.實支 > 0)
                    {
                        isCash = true;
                    }
                    if (vw_GBCVisaDetail.預付轉正 > 0)
                    {
                        isPre = true;
                    }
                    if (vw_GBCVisaDetail.沖抵估列 > 0)
                    {
                        isEst = true;
                    }
                    if (vw_GBCVisaDetail.費用 > 0)
                    {
                        isFee = true;
                    }

                    #region 實支(支出or付憑)，無估列應付
                    if (isPre == false && isEst == false)
                    {
                        try
                        {
                            isLog = dao.FindLog(x => x.基金代碼 == vw_GBCVisaDetail.基金代碼 && x.PK_會計年度 == vw_GBCVisaDetail.PK_會計年度 && x.PK_動支編號 == vw_GBCVisaDetail.PK_動支編號 && x.PK_種類 == vw_GBCVisaDetail.PK_種類 && x.PK_次別 == vw_GBCVisaDetail.PK_次別 && x.PK_明細號 == vw_GBCVisaDetail.PK_明細號);
                            isPass = jsonDAO.IsPass(x => x.基金代碼 == vw_GBCVisaDetail.基金代碼 && x.PFK_會計年度 == vw_GBCVisaDetail.PK_會計年度 && x.PFK_動支編號 == vw_GBCVisaDetail.PK_動支編號 && x.PFK_種類 == vw_GBCVisaDetail.PK_種類 && x.PFK_次別 == vw_GBCVisaDetail.PK_次別);
                            if ((isLog > 0) && isPass.Equals("1"))
                            {
                                return "此筆資料已轉入過,並且結案。";
                            }
                            //else if (((isLog > 0) && isPass.Equals("0")) || (isPass.Equals("0")))
                            else if (((isLog > 0) && isPass.Equals("0")))
                            {
                                dao.Update(vw_GBCVisaDetail);
                                jsonDAO.DeleteJsonRecord1(vw_GBCVisaDetail);
                            }
                            else
                            {
                                dao.Insert(vw_GBCVisaDetail);
                            }
                        }
                        catch (Exception e)
                        {
                            return e.Message;
                        }

                        傳票明細 vouDtl_D = new 傳票明細()
                        {
                            借貸別 = "借",
                            科目代號 = "5",
                            科目名稱 = "基金用途",
                            摘要 = vw_GBCVisaDetail.F_摘要,
                            金額 = vw_GBCVisaDetail.F_核定金額,
                            計畫代碼 = vw_GBCVisaDetail.F_計畫代碼,
                            用途別代碼 = vw_GBCVisaDetail.F_用途別代碼,
                            沖轉字號 = "",
                            對象代碼 = vw_GBCVisaDetail.F_受款人編號,
                            對象說明 = vw_GBCVisaDetail.F_受款人,
                            明細號 = vw_GBCVisaDetail.PK_明細號
                        };

                        //是否暫付及待結轉帳項
                        if (vw_GBCVisaDetail.F_計畫代碼 == "C1")
                        {
                            vouDtl_D.科目代號 = "1315";
                            vouDtl_D.科目名稱 = "暫付及待結轉帳事項";
                            vouDtl_D.用途別代碼 = "";
                        }

                        //是否補貼息
                        if (vw_GBCVisaDetail.F_計畫代碼.Substring(0, 2) == "01")
                        {
                            vouDtl_D.科目代號 = "2125";
                            vouDtl_D.科目名稱 = "應付費用";
                        }

                        //是否為以前年度
                        if (int.Parse(vw_GBCVisaDetail.PK_動支編號.Substring(0, 3)) < int.Parse(vw_GBCVisaDetail.PK_會計年度))
                        {
                            vouDtl_D.用途別代碼 = "91Y";
                        }
                        vouDtlList.Add(vouDtl_D);
                        傳票受款人 vouPay = new 傳票受款人()
                        {
                            //統一編號 = vw_GBCVisaDetail.F_受款人編號,
                            //受款人名稱 = vw_GBCVisaDetail.F_受款人,
                            //地址 = "",
                            //實付金額 = vw_GBCVisaDetail.F_核定金額,
                            //銀行代號 = "",
                            //銀行名稱 = "",
                            //銀行帳號 = "",
                            //帳戶名稱 = ""
                            統一編號 = "",
                            受款人名稱 = "",
                            地址 = "",
                            實付金額 = 0,
                            銀行代號 = "",
                            銀行名稱 = "",
                            銀行帳號 = "",
                            帳戶名稱 = ""
                        };
                        vouPayList.Add(vouPay);

                        detailCount += 1;

                        if (detailCount == vwList.Count())
                        {
                            vouMain.傳票種類 = PayVouKind;
                            vouMain.製票日期 = "";
                            vouMain.主摘要 = vw_GBCVisaDetail.F_摘要;
                            vouMain.交付方式 = "1";

                            傳票明細 vouDtl_C = new 傳票明細()
                            {
                                借貸別 = "貸",
                                科目代號 = "1112",
                                科目名稱 = "銀行存款",
                                摘要 = vw_GBCVisaDetail.F_摘要,
                                金額 = accSumMoney,
                                計畫代碼 = "",
                                用途別代碼 = "",
                                沖轉字號 = "",
                                對象代碼 = "",
                                對象說明 = ""
                            };

                            vouDtlList.Add(vouDtl_C);

                            vouCollection.傳票主檔 = vouMain;
                            vouCollection.傳票明細 = vouDtlList;
                            vouCollection.傳票受款人 = vouPayList;


                            vouCollectionList.Add(vouCollection);

                            vouTop.基金代碼 = vw_GBCVisaDetail.基金代碼;
                            vouTop.年度 = vw_GBCVisaDetail.PK_會計年度;
                            vouTop.動支編號 = vw_GBCVisaDetail.PK_動支編號;
                            vouTop.種類 = vw_GBCVisaDetail.PK_種類;
                            vouTop.次別 = vw_GBCVisaDetail.PK_次別;
                            vouTop.明細號 = vw_GBCVisaDetail.PK_明細號;
                            vouTop.傳票內容 = vouCollectionList;
                        }                        

                    }
                    #endregion

                    #region 實支(支出or付憑)，有估列應付
                    if (isPre == false && isEst == true)
                    {
                        try
                        {
                            isLog = dao.FindLog(x => x.基金代碼 == vw_GBCVisaDetail.基金代碼 && x.PK_會計年度 == vw_GBCVisaDetail.PK_會計年度 && x.PK_動支編號 == vw_GBCVisaDetail.PK_動支編號 && x.PK_種類 == vw_GBCVisaDetail.PK_種類 && x.PK_次別 == vw_GBCVisaDetail.PK_次別 && x.PK_明細號 == vw_GBCVisaDetail.PK_明細號);
                            isPass = jsonDAO.IsPass(x => x.基金代碼 == vw_GBCVisaDetail.基金代碼 && x.PFK_會計年度 == vw_GBCVisaDetail.PK_會計年度 && x.PFK_動支編號 == vw_GBCVisaDetail.PK_動支編號 && x.PFK_種類 == vw_GBCVisaDetail.PK_種類 && x.PFK_次別 == vw_GBCVisaDetail.PK_次別);
                            if ((isLog > 0) && isPass.Equals("1"))
                            {
                                return "此筆資料已轉入過,並且結案。";
                            }
                            //else if (((isLog > 0) && isPass.Equals("0")) || (isPass.Equals("0")))
                            else if (((isLog > 0) && isPass.Equals("0")))
                            {
                                dao.Update(vw_GBCVisaDetail);
                                jsonDAO.DeleteJsonRecord1(vw_GBCVisaDetail);
                            }
                            else
                            {
                                dao.Insert(vw_GBCVisaDetail);
                            }
                        }
                        catch (Exception e)
                        {
                            return e.Message;
                        }

                        傳票明細 vouDtl_D1 = new 傳票明細()
                        {
                            借貸別 = "借",
                            科目代號 = "2125",
                            科目名稱 = "應付費用",
                            摘要 = vw_GBCVisaDetail.F_摘要,
                            金額 = vw_GBCVisaDetail.沖抵估列,
                            計畫代碼 = vw_GBCVisaDetail.F_計畫代碼,
                            用途別代碼 = vw_GBCVisaDetail.F_用途別代碼,
                            沖轉字號 = "",
                            對象代碼 = vw_GBCVisaDetail.F_受款人編號,
                            對象說明 = vw_GBCVisaDetail.F_受款人,
                            明細號 = vw_GBCVisaDetail.PK_明細號
                        };

                        var estVouNo = dao.GetGBCVisaDetailAbateDetail(x => x.基金代碼 == vw_GBCVisaDetail.基金代碼 && x.PK_會計年度 == vw_GBCVisaDetail.PK_會計年度 && x.PK_動支編號 == vw_GBCVisaDetail.PK_動支編號 && x.PK_種類 == "估列" && x.F_受款人編號 == vw_GBCVisaDetail.F_受款人編號).OrderBy(x => x.F_製票日期1).FirstOrDefault();

                        if (estVouNo != null)
                        {
                            string estVouYear = estVouNo.F_傳票年度;
                            string estVouMainNo = estVouNo.F_傳票號1;
                            string estVouDtlNo = estVouNo.F_傳票明細號1.ToString();
                            vouDtl_D1.沖轉字號 = estVouYear + "-" + estVouMainNo + "-" + estVouDtlNo;
                        }
                        else
                        {
                            vouDtl_D1.沖轉字號 = "";
                        }

                        vouDtlList.Add(vouDtl_D1);

                        //是否開立第二列明細
                        if (vw_GBCVisaDetail.費用 > 0)
                        {                        
                            傳票明細 vouDtl_D2 = new 傳票明細()
                            {
                                借貸別 = "借",
                                科目代號 = "5",
                                科目名稱 = "基金用途",
                                摘要 = vw_GBCVisaDetail.F_摘要,
                                金額 = vw_GBCVisaDetail.費用,
                                計畫代碼 = vw_GBCVisaDetail.F_計畫代碼,
                                用途別代碼 = vw_GBCVisaDetail.F_用途別代碼,
                                沖轉字號 = "",
                                對象代碼 = vw_GBCVisaDetail.F_受款人編號,
                                對象說明 = vw_GBCVisaDetail.F_受款人,
                                明細號 = vw_GBCVisaDetail.PK_明細號
                            };

                            vouDtlList.Add(vouDtl_D2);
                        }

                        傳票受款人 vouPay = new 傳票受款人()
                        {
                            //統一編號 = vw_GBCVisaDetail.F_受款人編號,
                            //受款人名稱 = vw_GBCVisaDetail.F_受款人,
                            //地址 = "",
                            //實付金額 = vw_GBCVisaDetail.F_核定金額,
                            //銀行代號 = "",
                            //銀行名稱 = "",
                            //銀行帳號 = "",
                            //帳戶名稱 = ""
                            統一編號 = "",
                            受款人名稱 = "",
                            地址 = "",
                            實付金額 = 0,
                            銀行代號 = "",
                            銀行名稱 = "",
                            銀行帳號 = "",
                            帳戶名稱 = ""
                        };
                        vouPayList.Add(vouPay);

                        detailCount += 1;

                        if (detailCount == vwList.Count())
                        {
                            vouMain.傳票種類 = PayVouKind;
                            vouMain.製票日期 = "";
                            vouMain.主摘要 = vw_GBCVisaDetail.F_摘要;
                            vouMain.交付方式 = "1";

                            傳票明細 vouDtl_C = new 傳票明細()
                            {
                                借貸別 = "貸",
                                科目代號 = "1112",
                                科目名稱 = "銀行存款",
                                摘要 = vw_GBCVisaDetail.F_摘要,
                                金額 = accSumMoney,
                                計畫代碼 = "",
                                用途別代碼 = "",
                                沖轉字號 = "",
                                對象代碼 = "",
                                對象說明 = ""
                            };

                            vouDtlList.Add(vouDtl_C);

                            vouCollection.傳票主檔 = vouMain;
                            vouCollection.傳票明細 = vouDtlList;
                            vouCollection.傳票受款人 = vouPayList;


                            vouCollectionList.Add(vouCollection);

                            vouTop.基金代碼 = vw_GBCVisaDetail.基金代碼;
                            vouTop.年度 = vw_GBCVisaDetail.PK_會計年度;
                            vouTop.動支編號 = vw_GBCVisaDetail.PK_動支編號;
                            vouTop.種類 = vw_GBCVisaDetail.PK_種類;
                            vouTop.次別 = vw_GBCVisaDetail.PK_次別;
                            vouTop.明細號 = vw_GBCVisaDetail.PK_明細號;
                            vouTop.傳票內容 = vouCollectionList;
                        }

                    }
                    #endregion

                    #region 預付轉正
                    if (isPre)
                    {
                        #region 單張分轉情形
                        if (isCash == false)
                        {
                            try
                            {
                                isLog = dao.FindLog(x => x.基金代碼 == vw_GBCVisaDetail.基金代碼 && x.PK_會計年度 == vw_GBCVisaDetail.PK_會計年度 && x.PK_動支編號 == vw_GBCVisaDetail.PK_動支編號 && x.PK_種類 == vw_GBCVisaDetail.PK_種類 && x.PK_次別 == vw_GBCVisaDetail.PK_次別 && x.PK_明細號 == vw_GBCVisaDetail.PK_明細號);
                                isPass = jsonDAO.IsPass(x => x.基金代碼 == vw_GBCVisaDetail.基金代碼 && x.PFK_會計年度 == vw_GBCVisaDetail.PK_會計年度 && x.PFK_動支編號 == vw_GBCVisaDetail.PK_動支編號 && x.PFK_種類 == vw_GBCVisaDetail.PK_種類 && x.PFK_次別 == vw_GBCVisaDetail.PK_次別);
                                if ((isLog > 0) && isPass.Equals("1"))
                                {
                                    return "此筆資料已轉入過,並且結案。";
                                }
                                else if (((isLog > 0) && isPass.Equals("0")))
                                {
                                    dao.Update(vw_GBCVisaDetail);
                                    jsonDAO.DeleteJsonRecord1(vw_GBCVisaDetail);
                                }
                                else
                                {
                                    dao.Insert(vw_GBCVisaDetail);
                                }
                            }
                            catch (Exception e)
                            {
                                return e.Message;
                            }

                            double preAmount = 0;
                            double preAmount_tot = 0;
                            double abatAmount_tot = 0;
                            string abatePreVouNo = "";

                            var prepayList = dao.GetGBCVisaDetailAbateDetail(x => x.基金代碼 == vw_GBCVisaDetail.基金代碼 && x.PK_會計年度 == vw_GBCVisaDetail.PK_會計年度 && x.PK_動支編號 == vw_GBCVisaDetail.PK_動支編號 && x.PK_種類 =="預付" && x.F_受款人編號 == vw_GBCVisaDetail.F_受款人編號).OrderBy(x => x.F_製票日期1).ToList();
                            //預付總額
                            preAmount_tot = double.Parse((from s1 in prepayList select s1.F_核定金額).Sum().ToString());  

                            //算已轉正金額
                            foreach (var prepayListItem in prepayList)
                            {
                                string preVouYear = prepayListItem.F_傳票年度;
                                string preVouNo = prepayListItem.F_傳票號1;
                                string preVouDtlNo = prepayListItem.F_傳票明細號1.ToString();
                                var tmpPreVoDetail = vouDetailDAO.GetVouDetail(x => x.FundNo == prepayListItem.基金代碼 && x.RelatedVouNo == preVouYear + "-" + preVouNo + "-" + preVouDtlNo && x.DC == "貸" && x.SubNo == "1154").ToList();
                                
                                var tmpAbatMoney = tmpPreVoDetail == null? 0 : (from m1 in tmpPreVoDetail select m1.Amount).Sum();
                                abatAmount_tot = abatAmount_tot + double.Parse(tmpAbatMoney.ToString());
                            }

                            //判斷沖轉字號(預付明細疊加 - 已沖總額 是否大於等於本次報支金額，若true就帶本列預付傳票號)
                            foreach (var prepayListItem in prepayList)
                            {
                                preAmount = preAmount + double.Parse(prepayListItem.F_核定金額.ToString());
                                if (preAmount - abatAmount_tot >= vw_GBCVisaDetail.F_核定金額)
                                {
                                    abatePreVouNo = prepayListItem.PK_會計年度 + "-" + prepayListItem.F_傳票號1 + "-" + prepayListItem.F_傳票明細號1;
                                }
                            }

                            傳票明細 vouDtl_C = new 傳票明細()
                            {
                                借貸別 = "貸",
                                科目代號 = "1154",
                                科目名稱 = "預付費用",
                                摘要 = vw_GBCVisaDetail.F_摘要,
                                金額 = vw_GBCVisaDetail.預付轉正,
                                計畫代碼 = vw_GBCVisaDetail.F_計畫代碼,
                                用途別代碼 = vw_GBCVisaDetail.F_用途別代碼,
                                //沖轉字號 = abatePrePayVouYear.ElementAt(abateCnt) + "-" + abatePrePayVouNo.ElementAt(abateCnt) + "-" + abatePrePayVouDtlNo.ElementAt(abateCnt),
                                沖轉字號 = abatePreVouNo,
                                對象代碼 = "",
                                對象說明 = "",
                                明細號 = vw_GBCVisaDetail.PK_明細號
                            };
                            
                            傳票明細 vouDtl_D = new 傳票明細()
                            {
                                借貸別 = "借",
                                科目代號 = "5",
                                科目名稱 = "基金用途",
                                摘要 = vw_GBCVisaDetail.F_摘要,
                                金額 = vw_GBCVisaDetail.預付轉正,
                                計畫代碼 = vw_GBCVisaDetail.F_計畫代碼,
                                用途別代碼 = vw_GBCVisaDetail.F_用途別代碼,
                                沖轉字號 = "",
                                對象代碼 = vw_GBCVisaDetail.F_受款人編號,
                                對象說明 = vw_GBCVisaDetail.F_受款人,
                                明細號 = vw_GBCVisaDetail.PK_明細號
                            };

                            if (isEst)
                            {
                                vouDtl_C.金額 = vw_GBCVisaDetail.預付轉正 - vw_GBCVisaDetail.費用;

                                vouDtl_D.科目代號 = "2125";
                                vouDtl_D.科目名稱 = "應付費用";
                                vouDtl_D.金額 = vw_GBCVisaDetail.沖抵估列;
                                var estVouNo = dao.GetGBCVisaDetailAbateDetail(x => x.基金代碼 == vw_GBCVisaDetail.基金代碼 && x.PK_會計年度 == vw_GBCVisaDetail.PK_會計年度 && x.PK_動支編號 == vw_GBCVisaDetail.PK_動支編號 && x.PK_種類 == "估列" && x.F_受款人編號 == vw_GBCVisaDetail.F_受款人編號).OrderBy(x => x.F_製票日期1).FirstOrDefault();

                                if (estVouNo != null)
                                {
                                    string estVouYear = estVouNo.F_傳票年度;
                                    string estVouMainNo = estVouNo.F_傳票號1;
                                    string estVouDtlNo = estVouNo.F_傳票明細號1.ToString();
                                    vouDtl_D.沖轉字號 = estVouYear + "-" + estVouMainNo + "-" + estVouDtlNo;
                                }
                                else
                                {
                                    vouDtl_D.沖轉字號 = "";
                                }
                                
                            }
                            vouDtlList.Add(vouDtl_C);
                            vouDtlList.Add(vouDtl_D);

                            //沖抵應付時，有費用需增加預付轉實付
                            if (isFee && isEst)
                            {
                                vouDtl_C = new 傳票明細()
                                {
                                    借貸別 = "貸",
                                    科目代號 = "1154",
                                    科目名稱 = "預付費用",
                                    摘要 = vw_GBCVisaDetail.F_摘要,
                                    金額 = vw_GBCVisaDetail.費用,
                                    計畫代碼 = vw_GBCVisaDetail.F_計畫代碼,
                                    用途別代碼 = vw_GBCVisaDetail.F_用途別代碼,
                                    //沖轉字號 = abatePrePayVouYear.ElementAt(abateCnt) + "-" + abatePrePayVouNo.ElementAt(abateCnt) + "-" + abatePrePayVouDtlNo.ElementAt(abateCnt),
                                    沖轉字號 = abatePreVouNo,
                                    對象代碼 = "",
                                    對象說明 = "",
                                    明細號 = vw_GBCVisaDetail.PK_明細號
                                };
                                vouDtlList.Add(vouDtl_C);

                                vouDtl_D = new 傳票明細()
                                {
                                    借貸別 = "借",
                                    科目代號 = "5",
                                    科目名稱 = "基金用途",
                                    摘要 = vw_GBCVisaDetail.F_摘要,
                                    金額 = vw_GBCVisaDetail.費用,
                                    計畫代碼 = vw_GBCVisaDetail.F_計畫代碼,
                                    用途別代碼 = vw_GBCVisaDetail.F_用途別代碼,
                                    沖轉字號 = "",
                                    對象代碼 = vw_GBCVisaDetail.F_受款人編號,
                                    對象說明 = vw_GBCVisaDetail.F_受款人,
                                    明細號 = vw_GBCVisaDetail.PK_明細號
                                };
                                vouDtlList.Add(vouDtl_D);
                            }

                            傳票受款人 vouPay = new 傳票受款人()
                            {
                                統一編號 = "",
                                受款人名稱 = "",
                                地址 = "",
                                實付金額 = 0,
                                銀行代號 = "",
                                銀行名稱 = "",
                                銀行帳號 = "",
                                帳戶名稱 = ""
                            };
                            vouPayList.Add(vouPay);

                            detailCount += 1;

                            if (detailCount == vwList.Count())
                            {
                                vouMain.傳票種類 = "4";
                                vouMain.製票日期 = "";
                                vouMain.主摘要 = vw_GBCVisaDetail.F_摘要;
                                vouMain.交付方式 = "1";

                                vouCollection.傳票主檔 = vouMain;
                                vouCollection.傳票明細 = vouDtlList;
                                vouCollection.傳票受款人 = vouPayList;
                                vouCollectionList.Add(vouCollection);

                                vouTop.基金代碼 = vw_GBCVisaDetail.基金代碼;
                                vouTop.年度 = vw_GBCVisaDetail.PK_會計年度;
                                vouTop.動支編號 = vw_GBCVisaDetail.PK_動支編號;
                                vouTop.種類 = vw_GBCVisaDetail.PK_種類;
                                vouTop.次別 = vw_GBCVisaDetail.PK_次別;
                                vouTop.明細號 = vw_GBCVisaDetail.PK_明細號;
                                vouTop.傳票內容 = vouCollectionList;
                            }
                        }
                        #endregion

                        #region 須加開支出傳票之情形             
                        else
                        {
                            #region 轉正+實支(無應付)
                            if (isEst == false)
                            {

                                try
                                {
                                    isLog = dao.FindLog(x => x.基金代碼 == vw_GBCVisaDetail.基金代碼 && x.PK_會計年度 == vw_GBCVisaDetail.PK_會計年度 && x.PK_動支編號 == vw_GBCVisaDetail.PK_動支編號 && x.PK_種類 == vw_GBCVisaDetail.PK_種類 && x.PK_次別 == vw_GBCVisaDetail.PK_次別 && x.PK_明細號 == vw_GBCVisaDetail.PK_明細號);
                                    isPass = jsonDAO.IsPass(x => x.基金代碼 == vw_GBCVisaDetail.基金代碼 && x.PFK_會計年度 == vw_GBCVisaDetail.PK_會計年度 && x.PFK_動支編號 == vw_GBCVisaDetail.PK_動支編號 && x.PFK_種類 == vw_GBCVisaDetail.PK_種類 && x.PFK_次別 == vw_GBCVisaDetail.PK_次別);
                                    if ((isLog > 0) && isPass.Equals("1"))
                                    {
                                        return "此筆資料已轉入過,並且結案。";
                                    }
                                    else if (((isLog > 0) && isPass.Equals("0")))
                                    {
                                        dao.Update(vw_GBCVisaDetail);
                                        jsonDAO.DeleteJsonRecord1(vw_GBCVisaDetail);
                                    }
                                    else
                                    {
                                        dao.Insert(vw_GBCVisaDetail);
                                    }
                                }
                                catch (Exception e)
                                {
                                    return e.Message;
                                }

                                傳票明細 vouDtl_C = new 傳票明細()
                                {
                                    借貸別 = "貸",
                                    科目代號 = "1154",
                                    科目名稱 = "預付費用",
                                    摘要 = vw_GBCVisaDetail.F_摘要,
                                    金額 = vw_GBCVisaDetail.預付轉正,
                                    計畫代碼 = vw_GBCVisaDetail.F_計畫代碼,
                                    用途別代碼 = vw_GBCVisaDetail.F_用途別代碼,
                                    //沖轉字號 = abatePrePayVouYear.ElementAt(abateCnt) + "-" + abatePrePayVouNo.ElementAt(abateCnt) + "-" + abatePrePayVouDtlNo.ElementAt(abateCnt),
                                    沖轉字號 = "",
                                    對象代碼 = "",
                                    對象說明 = "",
                                    明細號 = vw_GBCVisaDetail.PK_明細號
                                };
                                vouDtlList.Add(vouDtl_C);
                                傳票明細 vouDtl_D = new 傳票明細()
                                {
                                    借貸別 = "借",
                                    科目代號 = "5",
                                    科目名稱 = "基金用途",
                                    摘要 = vw_GBCVisaDetail.F_摘要,
                                    金額 = vw_GBCVisaDetail.預付轉正,
                                    計畫代碼 = vw_GBCVisaDetail.F_計畫代碼,
                                    用途別代碼 = vw_GBCVisaDetail.F_用途別代碼,
                                    //沖轉字號 = abateEstimateVouYear.ElementAt(abateCnt) + "-" + abateEstimateVouNo.ElementAt(abateCnt) + "-" + abateEstimateVouDtlNo.ElementAt(abateCnt),
                                    沖轉字號 = "",
                                    對象代碼 = vw_GBCVisaDetail.F_受款人編號,
                                    對象說明 = vw_GBCVisaDetail.F_受款人,
                                    明細號 = vw_GBCVisaDetail.PK_明細號
                                };
                                vouDtlList.Add(vouDtl_D);
                                傳票受款人 vouPay = new 傳票受款人()
                                {
                                    //統一編號 = vw_GBCVisaDetail.F_受款人編號,
                                    //受款人名稱 = vw_GBCVisaDetail.F_受款人,
                                    //地址 = "",
                                    //實付金額 = prePayBalance,
                                    //銀行代號 = "",
                                    //銀行名稱 = "",
                                    //銀行帳號 = "",
                                    //帳戶名稱 = ""
                                    統一編號 = "",
                                    受款人名稱 = "",
                                    地址 = "",
                                    實付金額 = 0,
                                    銀行代號 = "",
                                    銀行名稱 = "",
                                    銀行帳號 = "",
                                    帳戶名稱 = ""
                                };
                                vouPayList.Add(vouPay);

                                vouMain.傳票種類 = "4";
                                vouMain.製票日期 = "";
                                vouMain.主摘要 = vw_GBCVisaDetail.F_摘要;
                                vouMain.交付方式 = "1";

                                vouCollection.傳票主檔 = vouMain;
                                vouCollection.傳票明細 = vouDtlList;
                                vouCollection.傳票受款人 = vouPayList;
                                vouCollectionList.Add(vouCollection);

                                vouTop.基金代碼 = vw_GBCVisaDetail.基金代碼;
                                vouTop.年度 = vw_GBCVisaDetail.PK_會計年度;
                                vouTop.動支編號 = vw_GBCVisaDetail.PK_動支編號;
                                vouTop.種類 = vw_GBCVisaDetail.PK_種類;
                                vouTop.次別 = vw_GBCVisaDetail.PK_次別;
                                vouTop.明細號 = vw_GBCVisaDetail.PK_明細號;
                                vouTop.傳票內容 = vouCollectionList;

                                //------支出傳票------
                                傳票明細 vouDtl_D2 = new 傳票明細()
                                {
                                    借貸別 = "借",
                                    科目代號 = "5",
                                    科目名稱 = "基金用途",
                                    摘要 = vw_GBCVisaDetail.F_摘要,
                                    金額 = vw_GBCVisaDetail.實支,
                                    計畫代碼 = vw_GBCVisaDetail.F_計畫代碼,
                                    用途別代碼 = vw_GBCVisaDetail.F_用途別代碼,
                                    沖轉字號 = "",
                                    對象代碼 = vw_GBCVisaDetail.F_受款人編號,
                                    對象說明 = vw_GBCVisaDetail.F_受款人,
                                    明細號 = vw_GBCVisaDetail.PK_明細號
                                };
                                vouDtlList2.Add(vouDtl_D2);
                                傳票受款人 vouPay2 = new 傳票受款人()
                                {
                                    //統一編號 = vw_GBCVisaDetail.F_受款人編號,
                                    //受款人名稱 = vw_GBCVisaDetail.F_受款人,
                                    //地址 = "",
                                    //實付金額 = vw_GBCVisaDetail.F_核定金額 - prePayBalance,
                                    //銀行代號 = "",
                                    //銀行名稱 = "",
                                    //銀行帳號 = "",
                                    //帳戶名稱 = ""
                                    統一編號 = "",
                                    受款人名稱 = "",
                                    地址 = "",
                                    實付金額 = 0,
                                    銀行代號 = "",
                                    銀行名稱 = "",
                                    銀行帳號 = "",
                                    帳戶名稱 = ""

                                };
                                vouPayList2.Add(vouPay2);

                                傳票主檔 vouMain2 = new 傳票主檔()
                                {
                                    傳票種類 = PayVouKind,
                                    製票日期 = "",
                                    主摘要 = vw_GBCVisaDetail.F_摘要,
                                    交付方式 = "1"
                                };
                                傳票明細 vouDtl_C2 = new 傳票明細()
                                {
                                    借貸別 = "貸",
                                    科目代號 = "1112",
                                    科目名稱 = "銀行存款",
                                    摘要 = vw_GBCVisaDetail.F_摘要,
                                    金額 = vw_GBCVisaDetail.實支,
                                    計畫代碼 = "",
                                    用途別代碼 = "",
                                    沖轉字號 = "",
                                    對象代碼 = "",
                                    對象說明 = "",
                                    明細號 = vw_GBCVisaDetail.PK_明細號
                                };
                                vouDtlList2.Add(vouDtl_C2);
                                傳票內容 vouCollection2 = new 傳票內容()
                                {
                                    傳票主檔 = vouMain2,
                                    傳票明細 = vouDtlList2,
                                    傳票受款人 = vouPayList2
                                };
                                vouCollectionList2.Add(vouCollection2);

                                vouTop2 = new 最外層()
                                {
                                    基金代碼 = vw_GBCVisaDetail.基金代碼,
                                    年度 = vw_GBCVisaDetail.PK_會計年度,
                                    動支編號 = vw_GBCVisaDetail.PK_動支編號,
                                    種類 = vw_GBCVisaDetail.PK_種類,
                                    次別 = vw_GBCVisaDetail.PK_次別,
                                    明細號 = vw_GBCVisaDetail.PK_明細號,
                                    傳票內容 = vouCollectionList2
                                };                                
                            }
                            #endregion

                            #region 轉正+實支(有應付)
                            else
                            {
                                try
                                {
                                    isLog = dao.FindLog(x => x.基金代碼 == vw_GBCVisaDetail.基金代碼 && x.PK_會計年度 == vw_GBCVisaDetail.PK_會計年度 && x.PK_動支編號 == vw_GBCVisaDetail.PK_動支編號 && x.PK_種類 == vw_GBCVisaDetail.PK_種類 && x.PK_次別 == vw_GBCVisaDetail.PK_次別 && x.PK_明細號 == vw_GBCVisaDetail.PK_明細號);
                                    isPass = jsonDAO.IsPass(x => x.基金代碼 == vw_GBCVisaDetail.基金代碼 && x.PFK_會計年度 == vw_GBCVisaDetail.PK_會計年度 && x.PFK_動支編號 == vw_GBCVisaDetail.PK_動支編號 && x.PFK_種類 == vw_GBCVisaDetail.PK_種類 && x.PFK_次別 == vw_GBCVisaDetail.PK_次別);
                                    if ((isLog > 0) && isPass.Equals("1"))
                                    {
                                        return "此筆資料已轉入過,並且結案。";
                                    }
                                    else if (((isLog > 0) && isPass.Equals("0")))
                                    {
                                        dao.Update(vw_GBCVisaDetail);
                                        jsonDAO.DeleteJsonRecord1(vw_GBCVisaDetail);
                                    }
                                    else
                                    {
                                        dao.Insert(vw_GBCVisaDetail);
                                    }
                                }
                                catch (Exception e)
                                {
                                    return e.Message;
                                }

                                傳票明細 vouDtl_C = new 傳票明細()
                                {
                                    借貸別 = "貸",
                                    科目代號 = "1154",
                                    科目名稱 = "預付費用",
                                    摘要 = vw_GBCVisaDetail.F_摘要,
                                    金額 = vw_GBCVisaDetail.預付轉正,
                                    計畫代碼 = vw_GBCVisaDetail.F_計畫代碼,
                                    用途別代碼 = vw_GBCVisaDetail.F_用途別代碼,
                                    //沖轉字號 = abatePrePayVouYear.ElementAt(abateCnt) + "-" + abatePrePayVouNo.ElementAt(abateCnt) + "-" + abatePrePayVouDtlNo.ElementAt(abateCnt),
                                    沖轉字號 = "",
                                    對象代碼 = "",
                                    對象說明 = "",
                                    明細號 = vw_GBCVisaDetail.PK_明細號
                                };
                                vouDtlList.Add(vouDtl_C);
                                傳票明細 vouDtl_D = new 傳票明細()
                                {
                                    借貸別 = "借",
                                    科目代號 = "2125",
                                    科目名稱 = "應付費用",
                                    摘要 = vw_GBCVisaDetail.F_摘要,
                                    金額 = vw_GBCVisaDetail.預付轉正,
                                    計畫代碼 = vw_GBCVisaDetail.F_計畫代碼,
                                    用途別代碼 = vw_GBCVisaDetail.F_用途別代碼,
                                    //沖轉字號 = abateEstimateVouYear.ElementAt(abateCnt) + "-" + abateEstimateVouNo.ElementAt(abateCnt) + "-" + abateEstimateVouDtlNo.ElementAt(abateCnt),
                                    沖轉字號 = "",
                                    對象代碼 = vw_GBCVisaDetail.F_受款人編號,
                                    對象說明 = vw_GBCVisaDetail.F_受款人,
                                    明細號 = vw_GBCVisaDetail.PK_明細號
                                };
                                vouDtlList.Add(vouDtl_D);

                                傳票受款人 vouPay = new 傳票受款人()
                                {
                                    //統一編號 = vw_GBCVisaDetail.F_受款人編號,
                                    //受款人名稱 = vw_GBCVisaDetail.F_受款人,
                                    //地址 = "",
                                    //實付金額 = prePayBalance,
                                    //銀行代號 = "",
                                    //銀行名稱 = "",
                                    //銀行帳號 = "",
                                    //帳戶名稱 = ""
                                    統一編號 = "",
                                    受款人名稱 = "",
                                    地址 = "",
                                    實付金額 = 0,
                                    銀行代號 = "",
                                    銀行名稱 = "",
                                    銀行帳號 = "",
                                    帳戶名稱 = ""
                                };
                                vouPayList.Add(vouPay);

                                vouMain.傳票種類 = "4";
                                vouMain.製票日期 = "";
                                vouMain.主摘要 = vw_GBCVisaDetail.F_摘要;
                                vouMain.交付方式 = "1";

                                vouCollection.傳票主檔 = vouMain;
                                vouCollection.傳票明細 = vouDtlList;
                                vouCollection.傳票受款人 = vouPayList;
                                vouCollectionList.Add(vouCollection);

                                vouTop.基金代碼 = vw_GBCVisaDetail.基金代碼;
                                vouTop.年度 = vw_GBCVisaDetail.PK_會計年度;
                                vouTop.動支編號 = vw_GBCVisaDetail.PK_動支編號;
                                vouTop.種類 = vw_GBCVisaDetail.PK_種類;
                                vouTop.次別 = vw_GBCVisaDetail.PK_次別;
                                vouTop.明細號 = vw_GBCVisaDetail.PK_明細號;
                                vouTop.傳票內容 = vouCollectionList;

                                //------支出傳票------
                                傳票明細 vouDtl_D2 = new 傳票明細()
                                {
                                    借貸別 = "借",
                                    科目代號 = "2125",
                                    科目名稱 = "應付費用",
                                    摘要 = vw_GBCVisaDetail.F_摘要,
                                    金額 = vw_GBCVisaDetail.實支,
                                    計畫代碼 = vw_GBCVisaDetail.F_計畫代碼,
                                    用途別代碼 = vw_GBCVisaDetail.F_用途別代碼,
                                    沖轉字號 = "",
                                    對象代碼 = vw_GBCVisaDetail.F_受款人編號,
                                    對象說明 = vw_GBCVisaDetail.F_受款人,
                                    明細號 = vw_GBCVisaDetail.PK_明細號
                                };
                                vouDtlList2.Add(vouDtl_D2);
                                傳票受款人 vouPay2 = new 傳票受款人()
                                {
                                    //統一編號 = vw_GBCVisaDetail.F_受款人編號,
                                    //受款人名稱 = vw_GBCVisaDetail.F_受款人,
                                    //地址 = "",
                                    //實付金額 = vw_GBCVisaDetail.F_核定金額 - prePayBalance,
                                    //銀行代號 = "",
                                    //銀行名稱 = "",
                                    //銀行帳號 = "",
                                    //帳戶名稱 = ""
                                    統一編號 = "",
                                    受款人名稱 = "",
                                    地址 = "",
                                    實付金額 = 0,
                                    銀行代號 = "",
                                    銀行名稱 = "",
                                    銀行帳號 = "",
                                    帳戶名稱 = ""

                                };
                                vouPayList2.Add(vouPay2);

                                傳票主檔 vouMain2 = new 傳票主檔()
                                {
                                    傳票種類 = PayVouKind,
                                    製票日期 = "",
                                    主摘要 = vw_GBCVisaDetail.F_摘要,
                                    交付方式 = "1"
                                };
                                傳票明細 vouDtl_C2 = new 傳票明細()
                                {
                                    借貸別 = "貸",
                                    科目代號 = "1112",
                                    科目名稱 = "銀行存款",
                                    摘要 = vw_GBCVisaDetail.F_摘要,
                                    金額 = vw_GBCVisaDetail.實支,
                                    計畫代碼 = "",
                                    用途別代碼 = "",
                                    沖轉字號 = "",
                                    對象代碼 = "",
                                    對象說明 = "",
                                    明細號 = vw_GBCVisaDetail.PK_明細號
                                };
                                vouDtlList2.Add(vouDtl_C2);
                                傳票內容 vouCollection2 = new 傳票內容()
                                {
                                    傳票主檔 = vouMain2,
                                    傳票明細 = vouDtlList2,
                                    傳票受款人 = vouPayList2
                                };
                                vouCollectionList2.Add(vouCollection2);

                                vouTop2 = new 最外層()
                                {
                                    基金代碼 = vw_GBCVisaDetail.基金代碼,
                                    年度 = vw_GBCVisaDetail.PK_會計年度,
                                    動支編號 = vw_GBCVisaDetail.PK_動支編號,
                                    種類 = vw_GBCVisaDetail.PK_種類,
                                    次別 = vw_GBCVisaDetail.PK_次別,
                                    明細號 = vw_GBCVisaDetail.PK_明細號,
                                    傳票內容 = vouCollectionList2
                                };                               
                            }
                            #endregion

                        }
                        #endregion
                    }
                    #endregion

                }

            }
            #endregion

            #region 估列
            if ("估列".Equals(accKind))
            {
                foreach (var vwListItem in vwList)
                {
                    vw_GBCVisaDetail.基金代碼 = vwListItem.基金代碼;
                    vw_GBCVisaDetail.PK_會計年度 = vwListItem.PK_會計年度;
                    vw_GBCVisaDetail.PK_動支編號 = vwListItem.PK_動支編號;
                    vw_GBCVisaDetail.PK_種類 = vwListItem.PK_種類;
                    vw_GBCVisaDetail.PK_次別 = vwListItem.PK_次別;
                    vw_GBCVisaDetail.PK_明細號 = vwListItem.PK_明細號;
                    vw_GBCVisaDetail.F_科室代碼 = vwListItem.F_科室代碼;
                    vw_GBCVisaDetail.F_用途別代碼 = vwListItem.F_用途別代碼;
                    vw_GBCVisaDetail.F_計畫代碼 = vwListItem.F_計畫代碼;
                    vw_GBCVisaDetail.F_動支金額 = vwListItem.F_動支金額;
                    vw_GBCVisaDetail.F_製票日 = vwListItem.F_製票日;
                    vw_GBCVisaDetail.F_是否核定 = vwListItem.F_是否核定;
                    vw_GBCVisaDetail.F_核定金額 = vwListItem.F_核定金額;
                    vw_GBCVisaDetail.F_核定日期 = vwListItem.F_核定日期;
                    vw_GBCVisaDetail.F_摘要 = vwListItem.F_摘要;
                    vw_GBCVisaDetail.F_受款人 = vwListItem.F_受款人;
                    vw_GBCVisaDetail.F_受款人編號 = vwListItem.F_受款人編號;
                    vw_GBCVisaDetail.F_原動支編號 = vwListItem.F_原動支編號;
                    vw_GBCVisaDetail.F_批號 = vwListItem.F_批號;

                    try
                    {
                        isLog = dao.FindLog(x => x.基金代碼 == vw_GBCVisaDetail.基金代碼 && x.PK_會計年度 == vw_GBCVisaDetail.PK_會計年度 && x.PK_動支編號 == vw_GBCVisaDetail.PK_動支編號 && x.PK_種類 == vw_GBCVisaDetail.PK_種類 && x.PK_次別 == vw_GBCVisaDetail.PK_次別 && x.PK_明細號 == vw_GBCVisaDetail.PK_明細號);
                        string isPass = jsonDAO.IsPass(x => x.基金代碼 == vw_GBCVisaDetail.基金代碼 && x.PFK_會計年度 == vw_GBCVisaDetail.PK_會計年度 && x.PFK_動支編號 == vw_GBCVisaDetail.PK_動支編號 && x.PFK_種類 == vw_GBCVisaDetail.PK_種類 && x.PFK_次別 == vw_GBCVisaDetail.PK_次別);

                        if ((isLog > 0) && isPass.Equals("1"))
                        {
                            return "此筆資料已轉入過,並且結案。";
                        }
                        else if (((isLog > 0) && isPass.Equals("0")))
                        {
                            dao.Update(vw_GBCVisaDetail);
                            jsonDAO.DeleteJsonRecord1(vw_GBCVisaDetail);
                        }
                        else
                        {
                            dao.Insert(vw_GBCVisaDetail);
                        }
                    }
                    catch (Exception e)
                    {
                        return e.Message;
                    }

                    傳票明細 vouDtl_D = new 傳票明細()
                    {
                        借貸別 = "借",
                        科目代號 = "5",
                        科目名稱 = "基金用途",
                        摘要 = vw_GBCVisaDetail.F_摘要 + "-" + vw_GBCVisaDetail.F_受款人,
                        金額 = vw_GBCVisaDetail.F_核定金額,
                        計畫代碼 = vw_GBCVisaDetail.F_計畫代碼,
                        用途別代碼 = vw_GBCVisaDetail.F_用途別代碼,
                        沖轉字號 = "",
                        對象代碼 = vw_GBCVisaDetail.F_受款人編號,
                        對象說明 = vw_GBCVisaDetail.F_受款人,

                    };
                    vouDtlList.Add(vouDtl_D);

                    傳票明細 vouDtl_C = new 傳票明細()
                    {
                        借貸別 = "貸",
                        科目代號 = "2125",
                        科目名稱 = "應付費用",
                        摘要 = vw_GBCVisaDetail.F_摘要 + "-" + vw_GBCVisaDetail.F_受款人,
                        金額 = vw_GBCVisaDetail.F_核定金額,
                        計畫代碼 = vw_GBCVisaDetail.F_計畫代碼,
                        用途別代碼 = vw_GBCVisaDetail.F_用途別代碼,
                        沖轉字號 = "",
                        對象代碼 = vw_GBCVisaDetail.F_受款人編號,
                        對象說明 = vw_GBCVisaDetail.F_受款人,
                        明細號 = vw_GBCVisaDetail.PK_明細號
                    };
                    vouDtlList.Add(vouDtl_C);
                    傳票受款人 vouPay = new 傳票受款人()
                    {
                        //統一編號 = vw_GBCVisaDetail.F_受款人編號,
                        //受款人名稱 = vw_GBCVisaDetail.F_受款人,
                        //地址 = "",
                        //實付金額 = vw_GBCVisaDetail.F_核定金額,
                        //銀行代號 = "",
                        //銀行名稱 = "",
                        //銀行帳號 = "",
                        //帳戶名稱 = ""
                        統一編號 = "",
                        受款人名稱 = "",
                        地址 = "",
                        實付金額 = 0,
                        銀行代號 = "",
                        銀行名稱 = "",
                        銀行帳號 = "",
                        帳戶名稱 = ""
                    };
                    vouPayList.Add(vouPay);

                    //填傳票明細號1
                    //dao.FillVouDtl1(vw_GBCVisaDetail.基金代碼, vw_GBCVisaDetail.PK_會計年度, vw_GBCVisaDetail.PK_動支編號, vw_GBCVisaDetail.PK_種類, vw_GBCVisaDetail.PK_次別, vw_GBCVisaDetail.PK_明細號, vouDtlList.Count);
                }
                //重新處理受款人清單,如果有重複受款人名稱,則金額加總
                //var vouPayGroup = from xxx in vouPayList
                //                  group xxx by new { xxx.統一編號, xxx.受款人名稱, xxx.地址, xxx.銀行代號, xxx.銀行名稱, xxx.銀行帳號, xxx.帳戶名稱 } into g
                //                  select new { 統一編號 = g.Key.統一編號, 受款人名稱 = g.Key.受款人名稱, 地址 = g.Key.地址, 銀行代號 = g.Key.銀行代號, 銀行名稱 = g.Key.銀行名稱, 銀行帳號 = g.Key.銀行帳號, 帳戶名稱 = g.Key.帳戶名稱, 實付金額 = g.Sum(xxx => xxx.實付金額) };
                //vouPayList = new List<傳票受款人>();
                //foreach (var vouPayGroupItem in vouPayGroup)
                //{
                //    傳票受款人 vouPay = new 傳票受款人();
                //    vouPay.統一編號 = vouPayGroupItem.統一編號;
                //    vouPay.受款人名稱 = vouPayGroupItem.受款人名稱;
                //    vouPay.地址 = vouPayGroupItem.地址;
                //    vouPay.實付金額 = vouPayGroupItem.實付金額;
                //    vouPay.銀行代號 = vouPayGroupItem.銀行代號;
                //    vouPay.銀行名稱 = vouPayGroupItem.銀行名稱;
                //    vouPay.銀行帳號 = vouPayGroupItem.銀行帳號;
                //    vouPay.帳戶名稱 = vouPayGroupItem.帳戶名稱;
                //    vouPayList.Add(vouPay);
                //}

                vouMain.傳票種類 = "4";
                vouMain.製票日期 = "";
                vouMain.主摘要 = vw_GBCVisaDetail.F_摘要;
                vouMain.交付方式 = "1";

                vouCollection.傳票主檔 = vouMain;
                vouCollection.傳票明細 = vouDtlList;
                vouCollection.傳票受款人 = vouPayList;


                vouCollectionList.Add(vouCollection);

                vouTop.基金代碼 = vw_GBCVisaDetail.基金代碼;
                vouTop.年度 = vw_GBCVisaDetail.PK_會計年度;
                vouTop.動支編號 = vw_GBCVisaDetail.PK_動支編號;
                vouTop.種類 = vw_GBCVisaDetail.PK_種類;
                vouTop.次別 = vw_GBCVisaDetail.PK_次別;
                vouTop.明細號 = vw_GBCVisaDetail.PK_明細號;
                vouTop.傳票內容 = vouCollectionList;

            }
            #endregion

            #region 估列收回
            if ("估列收回".Equals(accKind))
            {
                int estimateMoney = 0;
                int estimateMoneyAbate = 0;
                int estimateBalance = 0;

                foreach (var vwListItem in vwList)
                {
                    vw_GBCVisaDetail.基金代碼 = vwListItem.基金代碼;
                    vw_GBCVisaDetail.PK_會計年度 = vwListItem.PK_會計年度;
                    vw_GBCVisaDetail.PK_動支編號 = vwListItem.PK_動支編號;
                    vw_GBCVisaDetail.PK_種類 = vwListItem.PK_種類;
                    vw_GBCVisaDetail.PK_次別 = vwListItem.PK_次別;
                    vw_GBCVisaDetail.PK_明細號 = vwListItem.PK_明細號;
                    vw_GBCVisaDetail.F_科室代碼 = vwListItem.F_科室代碼;
                    vw_GBCVisaDetail.F_用途別代碼 = vwListItem.F_用途別代碼;
                    vw_GBCVisaDetail.F_計畫代碼 = vwListItem.F_計畫代碼;
                    vw_GBCVisaDetail.F_動支金額 = vwListItem.F_動支金額;
                    vw_GBCVisaDetail.F_製票日 = vwListItem.F_製票日;
                    vw_GBCVisaDetail.F_是否核定 = vwListItem.F_是否核定;
                    vw_GBCVisaDetail.F_核定金額 = vwListItem.F_核定金額;
                    vw_GBCVisaDetail.F_核定日期 = vwListItem.F_核定日期;
                    vw_GBCVisaDetail.F_摘要 = vwListItem.F_摘要;
                    vw_GBCVisaDetail.F_受款人 = vwListItem.F_受款人;
                    vw_GBCVisaDetail.F_受款人編號 = vwListItem.F_受款人編號;
                    vw_GBCVisaDetail.F_原動支編號 = vwListItem.F_原動支編號;
                    vw_GBCVisaDetail.F_批號 = vwListItem.F_批號;
                    try
                    {
                        isLog = dao.FindLog(x => x.基金代碼 == vw_GBCVisaDetail.基金代碼 && x.PK_會計年度 == vw_GBCVisaDetail.PK_會計年度 && x.PK_動支編號 == vw_GBCVisaDetail.PK_動支編號 && x.PK_種類 == vw_GBCVisaDetail.PK_種類 && x.PK_次別 == vw_GBCVisaDetail.PK_次別 && x.PK_明細號 == vw_GBCVisaDetail.PK_明細號);
                        string isPass = jsonDAO.IsPass(x => x.基金代碼 == vw_GBCVisaDetail.基金代碼 && x.PFK_會計年度 == vw_GBCVisaDetail.PK_會計年度 && x.PFK_動支編號 == vw_GBCVisaDetail.PK_動支編號 && x.PFK_種類 == vw_GBCVisaDetail.PK_種類 && x.PFK_次別 == vw_GBCVisaDetail.PK_次別);
                        if ((isLog > 0) && isPass.Equals("1"))
                        {
                            return "此筆資料已轉入過,並且結案。";
                        }
                        //else if (((isLog > 0) && isPass.Equals("0")) || (isPass.Equals("0")))
                        else if (((isLog > 0) && isPass.Equals("0")))
                        {
                            dao.Update(vw_GBCVisaDetail);
                            jsonDAO.DeleteJsonRecord1(vw_GBCVisaDetail);
                        }
                        else
                        {
                            dao.Insert(vw_GBCVisaDetail);
                        }
                    }
                    catch (Exception e)
                    {
                        return e.Message;
                    }

                    //計算估列沖銷餘額
                    var estimateNouNoList = dao.GetGBCVisaDetailAbateDetail(x => x.基金代碼 == vw_GBCVisaDetail.基金代碼 && x.PK_會計年度 == vw_GBCVisaDetail.PK_會計年度 && x.PK_動支編號 == vw_GBCVisaDetail.PK_動支編號 && x.PK_種類 == "估列").ToList();

                    //找應付沖轉字號
                    var abateEstimateVouYear = (from estvou in estimateNouNoList select estvou.F_傳票年度).ToList();
                    var abateEstimateVouNo = (from estvou in estimateNouNoList select estvou.F_傳票號1).ToList();
                    var abateEstimateVouDtlNo = (from estvou in estimateNouNoList select estvou.F_傳票明細號1).ToList();

                    int abateCnt = 0;

                    傳票明細 vouDtl_D = new 傳票明細()
                    {
                        借貸別 = "借",
                        科目代號 = "2125",
                        科目名稱 = "應付費用",
                        摘要 = vw_GBCVisaDetail.F_摘要,
                        金額 = vw_GBCVisaDetail.F_核定金額,
                        計畫代碼 = vw_GBCVisaDetail.F_計畫代碼,
                        用途別代碼 = vw_GBCVisaDetail.F_用途別代碼,
                        沖轉字號 = abateEstimateVouYear + "-" + abateEstimateVouNo.ElementAt(abateCnt) + "-" + abateEstimateVouDtlNo.ElementAt(abateCnt),
                        對象代碼 = vw_GBCVisaDetail.F_受款人編號,
                        對象說明 = vw_GBCVisaDetail.F_受款人,
                        明細號 = vw_GBCVisaDetail.PK_明細號
                    };
                    vouDtlList.Add(vouDtl_D);

                    傳票明細 vouDtl_C = new 傳票明細()
                    {
                        借貸別 = "貸",
                        科目代號 = "5",
                        科目名稱 = "基金用途",
                        摘要 = vw_GBCVisaDetail.F_摘要,
                        金額 = vw_GBCVisaDetail.F_核定金額,
                        計畫代碼 = vw_GBCVisaDetail.F_計畫代碼,
                        用途別代碼 = vw_GBCVisaDetail.F_用途別代碼,
                        沖轉字號 = "",
                        對象代碼 = vw_GBCVisaDetail.F_受款人編號,
                        對象說明 = vw_GBCVisaDetail.F_受款人
                    };

                    //是否為以前年度
                    if (int.Parse(vw_GBCVisaDetail.PK_動支編號.Substring(0, 3)) < int.Parse(vw_GBCVisaDetail.PK_會計年度))
                    {
                        vouDtl_C.科目代號 = "4YY";
                        vouDtl_C.科目名稱 = "雜項收入";
                        vouDtl_C.計畫代碼 = "";
                        vouDtl_C.用途別代碼 = "";
                    }
                    vouDtlList.Add(vouDtl_C);

                    傳票受款人 vouPay = new 傳票受款人()
                    {
                        統一編號 = vw_GBCVisaDetail.F_受款人編號,
                        受款人名稱 = vw_GBCVisaDetail.F_受款人,
                        地址 = "",
                        實付金額 = vw_GBCVisaDetail.F_核定金額,
                        銀行代號 = "",
                        銀行名稱 = "",
                        銀行帳號 = "",
                        帳戶名稱 = ""
                    };
                    vouPayList.Add(vouPay);

                    abateCnt++;

                    //填傳票明細號1
                    //dao.FillVouDtl1(vw_GBCVisaDetail.基金代碼, vw_GBCVisaDetail.PK_會計年度, vw_GBCVisaDetail.PK_動支編號, vw_GBCVisaDetail.PK_種類, vw_GBCVisaDetail.PK_次別, vw_GBCVisaDetail.PK_明細號, vouDtlList.Count);
                }
                //重新處理受款人清單,如果有重複受款人名稱,則金額加總
                //var vouPayGroup = from xxx in vouPayList
                //                  group xxx by new { xxx.統一編號, xxx.受款人名稱, xxx.地址, xxx.銀行代號, xxx.銀行名稱, xxx.銀行帳號, xxx.帳戶名稱 } into g
                //                  select new { 統一編號 = g.Key.統一編號, 受款人名稱 = g.Key.受款人名稱, 地址 = g.Key.地址, 銀行代號 = g.Key.銀行代號, 銀行名稱 = g.Key.銀行名稱, 銀行帳號 = g.Key.銀行帳號, 帳戶名稱 = g.Key.帳戶名稱, 實付金額 = g.Sum(xxx => xxx.實付金額) };
                //vouPayList = new List<傳票受款人>();
                //foreach (var vouPayGroupItem in vouPayGroup)
                //{
                //    傳票受款人 vouPay = new 傳票受款人();
                //    vouPay.統一編號 = vouPayGroupItem.統一編號;
                //    vouPay.受款人名稱 = vouPayGroupItem.受款人名稱;
                //    vouPay.地址 = vouPayGroupItem.地址;
                //    vouPay.實付金額 = vouPayGroupItem.實付金額;
                //    vouPay.銀行代號 = vouPayGroupItem.銀行代號;
                //    vouPay.銀行名稱 = vouPayGroupItem.銀行名稱;
                //    vouPay.銀行帳號 = vouPayGroupItem.銀行帳號;
                //    vouPay.帳戶名稱 = vouPayGroupItem.帳戶名稱;
                //    vouPayList.Add(vouPay);
                //}

                vouMain.傳票種類 = "4";
                vouMain.製票日期 = "";
                vouMain.主摘要 = vw_GBCVisaDetail.F_摘要;
                vouMain.交付方式 = "1";


                vouCollection.傳票主檔 = vouMain;
                vouCollection.傳票明細 = vouDtlList;
                vouCollection.傳票受款人 = vouPayList;

                vouCollectionList.Add(vouCollection);

                vouTop.基金代碼 = vw_GBCVisaDetail.基金代碼;
                vouTop.年度 = vw_GBCVisaDetail.PK_會計年度;
                vouTop.動支編號 = vw_GBCVisaDetail.PK_動支編號;
                vouTop.種類 = vw_GBCVisaDetail.PK_種類;
                vouTop.次別 = vw_GBCVisaDetail.PK_次別;
                vouTop.明細號 = vw_GBCVisaDetail.PK_明細號;
                vouTop.傳票內容 = vouCollectionList;
            }
            #endregion

            #region 預撥收回
            if ("預撥收回".Equals(accKind))
            {
                int prePayMoney = 0;
                int prePayMoneyAbate = 0;
                int prePayBalance = 0;

                //貸方要沖銷預付
                foreach (var vwListItem in vwList)
                {
                    vw_GBCVisaDetail.基金代碼 = vwListItem.基金代碼;
                    vw_GBCVisaDetail.PK_會計年度 = vwListItem.PK_會計年度;
                    vw_GBCVisaDetail.PK_動支編號 = vwListItem.PK_動支編號;
                    vw_GBCVisaDetail.PK_種類 = vwListItem.PK_種類;
                    vw_GBCVisaDetail.PK_次別 = vwListItem.PK_次別;
                    vw_GBCVisaDetail.PK_明細號 = vwListItem.PK_明細號;
                    vw_GBCVisaDetail.F_科室代碼 = vwListItem.F_科室代碼;
                    vw_GBCVisaDetail.F_用途別代碼 = vwListItem.F_用途別代碼;
                    vw_GBCVisaDetail.F_計畫代碼 = vwListItem.F_計畫代碼;
                    vw_GBCVisaDetail.F_動支金額 = vwListItem.F_動支金額;
                    vw_GBCVisaDetail.F_製票日 = vwListItem.F_製票日;
                    vw_GBCVisaDetail.F_是否核定 = vwListItem.F_是否核定;
                    vw_GBCVisaDetail.F_核定金額 = vwListItem.F_核定金額;
                    vw_GBCVisaDetail.F_核定日期 = vwListItem.F_核定日期;
                    vw_GBCVisaDetail.F_摘要 = vwListItem.F_摘要;
                    vw_GBCVisaDetail.F_受款人 = vwListItem.F_受款人;
                    vw_GBCVisaDetail.F_受款人編號 = vwListItem.F_受款人編號;
                    vw_GBCVisaDetail.F_原動支編號 = vwListItem.F_原動支編號;
                    vw_GBCVisaDetail.F_批號 = vwListItem.F_批號;
                    try
                    {
                        isLog = dao.FindLog(x => x.基金代碼 == vw_GBCVisaDetail.基金代碼 && x.PK_會計年度 == vw_GBCVisaDetail.PK_會計年度 && x.PK_動支編號 == vw_GBCVisaDetail.PK_動支編號 && x.PK_種類 == vw_GBCVisaDetail.PK_種類 && x.PK_次別 == vw_GBCVisaDetail.PK_次別 && x.PK_明細號 == vw_GBCVisaDetail.PK_明細號);
                        string isPass = jsonDAO.IsPass(x => x.基金代碼 == vw_GBCVisaDetail.基金代碼 && x.PFK_會計年度 == vw_GBCVisaDetail.PK_會計年度 && x.PFK_動支編號 == vw_GBCVisaDetail.PK_動支編號 && x.PFK_種類 == vw_GBCVisaDetail.PK_種類 && x.PFK_次別 == vw_GBCVisaDetail.PK_次別);
                        if ((isLog > 0) && isPass.Equals("1"))
                        {
                            return "此筆資料已轉入過,並且結案。";
                        }
                        //else if (((isLog > 0) && isPass.Equals("0")) || (isPass.Equals("0")))
                        else if (((isLog > 0) && isPass.Equals("0")))
                        {
                            dao.Update(vw_GBCVisaDetail);
                            jsonDAO.DeleteJsonRecord1(vw_GBCVisaDetail);
                        }
                        else
                        {
                            dao.Insert(vw_GBCVisaDetail);
                        }
                    }
                    catch (Exception e)
                    {
                        return e.Message;
                    }

                    //計算預付沖銷餘額
                    var prePayNouNoList = dao.GetGBCVisaDetailAbateDetail(x => x.基金代碼 == vw_GBCVisaDetail.基金代碼 && x.PK_動支編號 == vw_GBCVisaDetail.PK_動支編號 && x.PK_種類 == "預付" && x.F_受款人編號 == vw_GBCVisaDetail.F_受款人編號).ToList();

                    //foreach (var prePayVouNo in prePayNouNoList)
                    //{
                    //    prePayMoney = prePayMoney + dao.PrePayMoney(vwListItem.基金代碼, vwListItem.PK_會計年度, prePayVouNo.傳票號, prePayVouNo.傳票明細號);
                    //    prePayMoneyAbate = prePayMoneyAbate + dao.PrePayMoneyAbate(vwListItem.基金代碼, vwListItem.PK_會計年度, prePayVouNo.傳票號, prePayVouNo.傳票明細號);
                    //}

                    ////預付沖銷餘額 = 已預付 - 已轉正
                    //prePayBalance = prePayMoney - prePayMoneyAbate;

                    //if (prePayBalance <= 0)
                    //{
                    //    return "預付沖銷餘額不足!";
                    //}

                    //找預付沖轉字號
                    var abatePrePayVouYear = (from prevou in prePayNouNoList select prevou.F_傳票年度).ToList();
                    var abatePrePayVouNo = (from prevou in prePayNouNoList select prevou.F_傳票號1).ToList();
                    var abatePrePayVouDtlNo = (from prevou in prePayNouNoList select prevou.F_傳票明細號1).ToList();

                    int abateCnt = 0;

                    傳票明細 vouDtl_C = new 傳票明細()
                    {
                        借貸別 = "貸",
                        科目代號 = "1154",
                        科目名稱 = "預付費用",
                        摘要 = vw_GBCVisaDetail.F_摘要,
                        金額 = vw_GBCVisaDetail.F_核定金額,
                        計畫代碼 = vw_GBCVisaDetail.F_計畫代碼,
                        用途別代碼 = vw_GBCVisaDetail.F_用途別代碼,
                        沖轉字號 = abatePrePayVouYear.ElementAt(abateCnt) + "-" + abatePrePayVouNo.ElementAt(abateCnt) + "-" + abatePrePayVouDtlNo.ElementAt(abateCnt), //沖轉支出傳票 from prePayNouNoList
                        對象代碼 = vw_GBCVisaDetail.F_受款人編號,
                        對象說明 = vw_GBCVisaDetail.F_受款人,
                        明細號 = vw_GBCVisaDetail.PK_明細號
                    };
                    vouDtlList.Add(vouDtl_C);

                    傳票受款人 vouPay = new 傳票受款人()
                    {
                        //統一編號 = vw_GBCVisaDetail.F_受款人編號,
                        //受款人名稱 = vw_GBCVisaDetail.F_受款人,
                        //地址 = "",
                        //實付金額 = vw_GBCVisaDetail.F_核定金額,
                        //銀行代號 = "",
                        //銀行名稱 = "",
                        //銀行帳號 = "",
                        //帳戶名稱 = ""
                        統一編號 = "",
                        受款人名稱 = "",
                        地址 = "",
                        實付金額 = 0,
                        銀行代號 = "",
                        銀行名稱 = "",
                        銀行帳號 = "",
                        帳戶名稱 = ""
                    };
                    vouPayList.Add(vouPay);

                    abateCnt++;

                    //填傳票明細號1
                    //dao.FillVouDtl1(vw_GBCVisaDetail.基金代碼, vw_GBCVisaDetail.PK_會計年度, vw_GBCVisaDetail.PK_動支編號, vw_GBCVisaDetail.PK_種類, vw_GBCVisaDetail.PK_次別, vw_GBCVisaDetail.PK_明細號, vouDtlList.Count);
                }
                //重新處理受款人清單,如果有重複受款人名稱,則金額加總
                var vouPayGroup = from xxx in vouPayList
                                  group xxx by new { xxx.統一編號, xxx.受款人名稱, xxx.地址, xxx.銀行代號, xxx.銀行名稱, xxx.銀行帳號, xxx.帳戶名稱 } into g
                                  select new { 統一編號 = g.Key.統一編號, 受款人名稱 = g.Key.受款人名稱, 地址 = g.Key.地址, 銀行代號 = g.Key.銀行代號, 銀行名稱 = g.Key.銀行名稱, 銀行帳號 = g.Key.銀行帳號, 帳戶名稱 = g.Key.帳戶名稱, 實付金額 = g.Sum(xxx => xxx.實付金額) };
                //vouPayList = new List<傳票受款人>();
                //foreach (var vouPayGroupItem in vouPayGroup)
                //{
                //    傳票受款人 vouPay = new 傳票受款人();
                //    vouPay.統一編號 = vouPayGroupItem.統一編號;
                //    vouPay.受款人名稱 = vouPayGroupItem.受款人名稱;
                //    vouPay.地址 = vouPayGroupItem.地址;
                //    vouPay.實付金額 = vouPayGroupItem.實付金額;
                //    vouPay.銀行代號 = vouPayGroupItem.銀行代號;
                //    vouPay.銀行名稱 = vouPayGroupItem.銀行名稱;
                //    vouPay.銀行帳號 = vouPayGroupItem.銀行帳號;
                //    vouPay.帳戶名稱 = vouPayGroupItem.帳戶名稱;
                //    vouPayList.Add(vouPay);
                //}

                vouMain.傳票種類 = "1";
                vouMain.製票日期 = "";
                vouMain.主摘要 = vw_GBCVisaDetail.F_摘要;
                vouMain.交付方式 = "1";

                傳票明細 vouDtl_D = new 傳票明細()
                {
                    借貸別 = "借",
                    科目代號 = "1112",
                    科目名稱 = "銀行存款",
                    摘要 = vw_GBCVisaDetail.F_摘要,
                    金額 = accSumMoney,
                    計畫代碼 = "",
                    用途別代碼 = "",
                    沖轉字號 = "",
                    對象代碼 = vw_GBCVisaDetail.F_受款人編號,
                    對象說明 = vw_GBCVisaDetail.F_受款人
                };
                vouDtlList.Add(vouDtl_D);

                vouCollection.傳票主檔 = vouMain;
                vouCollection.傳票明細 = vouDtlList;
                vouCollection.傳票受款人 = vouPayList;
                vouCollectionList.Add(vouCollection);

                vouTop.基金代碼 = vw_GBCVisaDetail.基金代碼;
                vouTop.年度 = vw_GBCVisaDetail.PK_會計年度;
                vouTop.動支編號 = vw_GBCVisaDetail.PK_動支編號;
                vouTop.種類 = vw_GBCVisaDetail.PK_種類;
                vouTop.次別 = vw_GBCVisaDetail.PK_次別;
                vouTop.明細號 = vw_GBCVisaDetail.PK_明細號;
                vouTop.傳票內容 = vouCollectionList;
            }
            #endregion

            #region 核銷收回
            if ("核銷收回".Equals(accKind))
            {
                foreach (var item in vwList)
                {
                    vw_GBCVisaDetail.基金代碼 = item.基金代碼;
                    vw_GBCVisaDetail.PK_會計年度 = item.PK_會計年度;
                    vw_GBCVisaDetail.PK_動支編號 = item.PK_動支編號;
                    vw_GBCVisaDetail.PK_種類 = item.PK_種類;
                    vw_GBCVisaDetail.PK_次別 = item.PK_次別;
                    vw_GBCVisaDetail.PK_明細號 = item.PK_明細號;
                    vw_GBCVisaDetail.F_科室代碼 = item.F_科室代碼;
                    vw_GBCVisaDetail.F_用途別代碼 = item.F_用途別代碼;
                    vw_GBCVisaDetail.F_計畫代碼 = item.F_計畫代碼;
                    vw_GBCVisaDetail.F_動支金額 = item.F_動支金額;
                    vw_GBCVisaDetail.F_製票日 = item.F_製票日;
                    vw_GBCVisaDetail.F_是否核定 = item.F_是否核定;
                    vw_GBCVisaDetail.F_核定金額 = item.F_核定金額;
                    vw_GBCVisaDetail.F_核定日期 = item.F_核定日期;
                    vw_GBCVisaDetail.F_摘要 = item.F_摘要;
                    vw_GBCVisaDetail.F_受款人 = item.F_受款人;
                    vw_GBCVisaDetail.F_受款人編號 = item.F_受款人編號;
                    vw_GBCVisaDetail.F_原動支編號 = item.F_原動支編號;
                    vw_GBCVisaDetail.F_批號 = item.F_批號;
                    try
                    {
                        isLog = dao.FindLog(x => x.基金代碼 == vw_GBCVisaDetail.基金代碼 && x.PK_會計年度 == vw_GBCVisaDetail.PK_會計年度 && x.PK_動支編號 == vw_GBCVisaDetail.PK_動支編號 && x.PK_種類 == vw_GBCVisaDetail.PK_種類 && x.PK_次別 == vw_GBCVisaDetail.PK_次別 && x.PK_明細號 == vw_GBCVisaDetail.PK_明細號);
                        string isPass = jsonDAO.IsPass(x => x.基金代碼 == vw_GBCVisaDetail.基金代碼 && x.PFK_會計年度 == vw_GBCVisaDetail.PK_會計年度 && x.PFK_動支編號 == vw_GBCVisaDetail.PK_動支編號 && x.PFK_種類 == vw_GBCVisaDetail.PK_種類 && x.PFK_次別 == vw_GBCVisaDetail.PK_次別);
                        if ((isLog > 0) && isPass.Equals("1"))
                        {
                            return "此筆資料已轉入過,並且結案。";
                        }
                        //else if (((isLog > 0) && isPass.Equals("0")) || (isPass.Equals("0")))
                        else if (((isLog > 0) && isPass.Equals("0")))
                        {
                            dao.Update(vw_GBCVisaDetail);
                            jsonDAO.DeleteJsonRecord1(vw_GBCVisaDetail);
                        }
                        else
                        {
                            dao.Insert(vw_GBCVisaDetail);
                        }
                    }
                    catch (Exception e)
                    {
                        return e.Message;
                    }

                    //計算估列沖銷餘額
                    var payVouNoList = dao.GetGBCVisaDetailAbateDetail(x => x.基金代碼 == vw_GBCVisaDetail.基金代碼 && x.PK_會計年度 == vw_GBCVisaDetail.PK_會計年度 && x.PK_動支編號 == vw_GBCVisaDetail.PK_動支編號 && x.PK_種類 == "估列").ToList();

                    //找應付沖轉字號
                    var payVouNo = from payvou in payVouNoList select payvou.F_傳票號1;
                    var payVouDtlNo = from payvou in payVouNoList select payvou.F_傳票明細號1;

                    int abateCnt = 0;

                    傳票明細 vouDtl_C = new 傳票明細()
                    {
                        借貸別 = "貸",
                        科目代號 = "5",
                        科目名稱 = "基金用途",
                        摘要 = vw_GBCVisaDetail.F_摘要,
                        金額 = vw_GBCVisaDetail.F_核定金額,
                        計畫代碼 = vw_GBCVisaDetail.F_計畫代碼,
                        用途別代碼 = vw_GBCVisaDetail.F_用途別代碼,
                        沖轉字號 = payVouNo.ElementAt(abateCnt) + "-" + payVouDtlNo.ElementAt(abateCnt),
                        對象代碼 = vw_GBCVisaDetail.F_受款人編號,
                        對象說明 = vw_GBCVisaDetail.F_受款人,
                        明細號 = vw_GBCVisaDetail.PK_明細號
                    };

                    //是否為以前年度
                    if (int.Parse(vw_GBCVisaDetail.PK_動支編號.Substring(0, 3)) < int.Parse(vw_GBCVisaDetail.PK_會計年度))
                    {

                        vouDtl_C.科目代號 = "4YY";
                        vouDtl_C.科目名稱 = "雜項收入";
                        vouDtl_C.計畫代碼 = "";
                        vouDtl_C.用途別代碼 = "";
                        vouDtl_C.沖轉字號 = ""; //不用沖
                    }
                    vouDtlList.Add(vouDtl_C);

                    傳票受款人 vouPay = new 傳票受款人()
                    {
                        //統一編號 = vw_GBCVisaDetail.F_受款人編號,
                        //受款人名稱 = vw_GBCVisaDetail.F_受款人,
                        //地址 = "",
                        //實付金額 = vw_GBCVisaDetail.F_核定金額,
                        //銀行代號 = "",
                        //銀行名稱 = "",
                        //銀行帳號 = "",
                        //帳戶名稱 = ""
                        統一編號 = "",
                        受款人名稱 = "",
                        地址 = "",
                        實付金額 = 0,
                        銀行代號 = "",
                        銀行名稱 = "",
                        銀行帳號 = "",
                        帳戶名稱 = ""
                    };
                    vouPayList.Add(vouPay);

                    abateCnt++;

                    //填傳票明細號1
                    //dao.FillVouDtl1(vw_GBCVisaDetail.基金代碼, vw_GBCVisaDetail.PK_會計年度, vw_GBCVisaDetail.PK_動支編號, vw_GBCVisaDetail.PK_種類, vw_GBCVisaDetail.PK_次別, vw_GBCVisaDetail.PK_明細號, vouDtlList.Count);
                }
                //重新處理受款人清單,如果有重複受款人名稱,則金額加總
                var vouPayGroup = from xxx in vouPayList
                                  group xxx by new { xxx.統一編號, xxx.受款人名稱, xxx.地址, xxx.銀行代號, xxx.銀行名稱, xxx.銀行帳號, xxx.帳戶名稱 } into g
                                  select new { 統一編號 = g.Key.統一編號, 受款人名稱 = g.Key.受款人名稱, 地址 = g.Key.地址, 銀行代號 = g.Key.銀行代號, 銀行名稱 = g.Key.銀行名稱, 銀行帳號 = g.Key.銀行帳號, 帳戶名稱 = g.Key.帳戶名稱, 實付金額 = g.Sum(xxx => xxx.實付金額) };
                //vouPayList = new List<傳票受款人>();
                //foreach (var vouPayGroupItem in vouPayGroup)
                //{
                //    傳票受款人 vouPay = new 傳票受款人();
                //    vouPay.統一編號 = vouPayGroupItem.統一編號;
                //    vouPay.受款人名稱 = vouPayGroupItem.受款人名稱;
                //    vouPay.地址 = vouPayGroupItem.地址;
                //    vouPay.實付金額 = vouPayGroupItem.實付金額;
                //    vouPay.銀行代號 = vouPayGroupItem.銀行代號;
                //    vouPay.銀行名稱 = vouPayGroupItem.銀行名稱;
                //    vouPay.銀行帳號 = vouPayGroupItem.銀行帳號;
                //    vouPay.帳戶名稱 = vouPayGroupItem.帳戶名稱;
                //    vouPayList.Add(vouPay);
                //}

                vouMain.傳票種類 = PayVouKind;
                vouMain.製票日期 = "";
                vouMain.主摘要 = vw_GBCVisaDetail.F_摘要;
                vouMain.交付方式 = "1";

                傳票明細 vouDtl_D = new 傳票明細()
                {
                    借貸別 = "借",
                    科目代號 = "1112",
                    科目名稱 = "銀行存款",
                    摘要 = vw_GBCVisaDetail.F_摘要,
                    金額 = accSumMoney,
                    計畫代碼 = vw_GBCVisaDetail.F_計畫代碼,
                    用途別代碼 = vw_GBCVisaDetail.F_用途別代碼,
                    沖轉字號 = "",
                    對象代碼 = vw_GBCVisaDetail.F_受款人編號,
                    對象說明 = vw_GBCVisaDetail.F_受款人
                };
                vouDtlList.Add(vouDtl_D);

                vouCollection.傳票主檔 = vouMain;
                vouCollection.傳票明細 = vouDtlList;
                vouCollection.傳票受款人 = vouPayList;
                vouCollectionList.Add(vouCollection);

                vouTop.基金代碼 = vw_GBCVisaDetail.基金代碼;
                vouTop.年度 = vw_GBCVisaDetail.PK_會計年度;
                vouTop.動支編號 = vw_GBCVisaDetail.PK_動支編號;
                vouTop.種類 = vw_GBCVisaDetail.PK_種類;
                vouTop.次別 = vw_GBCVisaDetail.PK_次別;
                vouTop.明細號 = vw_GBCVisaDetail.PK_明細號;
                vouTop.傳票內容 = vouCollectionList;
            }
            #endregion

            //紀錄第一張傳票底稿
            try
            {
                jsonDAO.InsertJsonRecord1(vw_GBCVisaDetail, JsonConvert.SerializeObject(vouTop));
            }
            catch (Exception e)
            {
                return e.Message;
            }
            //回傳第一張傳票底稿
            JSON1 = jsonDAO.FindJSON1(x => x.基金代碼 == vw_GBCVisaDetail.基金代碼 && x.PFK_會計年度 == vw_GBCVisaDetail.PK_會計年度 && x.PFK_動支編號 == vw_GBCVisaDetail.PK_動支編號 && x.PFK_種類 == vw_GBCVisaDetail.PK_種類 && x.PFK_次別 == vw_GBCVisaDetail.PK_次別);

            //若有開立第二張，則紀錄第二張傳票底稿
            if (vouTop2 != null)
            {                
                try
                {
                    jsonDAO.InsertJsonRecord2(vw_GBCVisaDetail, JsonConvert.SerializeObject(vouTop2));
                }
                catch (Exception e)
                {
                    return e.Message;
                }
            }

            //return JsonConvert.SerializeObject(JSON1);
            return JSON1;
        }

        [WebMethod]
        /// <summary>
        /// 回填傳票號
        /// </summary>
        /// <param name="fundNo"></param>
        /// <param name="acmWordNum"></param>
        /// <param name="vouNoJSON"></param>
        /// <returns></returns>
        public string FillVouNo(string fundNo, string acmWordNum, string vouNoJSON)
        {
            ServicePointManager.ServerCertificateValidationCallback = new RemoteCertificateValidationCallback(ValidateServerCertificate);
            GBCVisaDetailAbateDetailDAO dao = new GBCVisaDetailAbateDetailDAO();
            GBCJSONRecordDAO jsonDAO = new GBCJSONRecordDAO();
            FillVouScript fillVouScript = new FillVouScript();
            //GBCJSONRecordVO gbcJSONRecordVO = new GBCJSONRecordVO();
            GBCVisaDetailAbateDetail gbcVisaDetailAbateDetail = new GBCVisaDetailAbateDetail();

            string isVouNo1 = "";
            string isVouNo2 = "";
            string isJSON2 = "";
            string isPass = "";
            int count = 0;


            vouNoJSON = vouNoJSON.Replace(@"\r\n", ""); //清除\r\n              
            vouNoJSON = vouNoJSON.Replace(@"\", "");    //清除\        
            vouNoJSON = vouNoJSON.Replace(@"""[", "["); //將 "[  改為 [
            vouNoJSON = vouNoJSON.Replace(@"]""", "]"); //將 ["  改為 ]

            //-------寫入Log------------------
            //jsonDAO.InsertJsonLog(fundNo, acmWordNum, vouNoJSON);
            //--------------------------------

            try
            {
                fillVouScript = JsonConvert.DeserializeObject<FillVouScript>(vouNoJSON);  //反序列化JSON
            }
            catch (Exception e)
            {
                return e.StackTrace;
            }

            string[] strs = acmWordNum.Split('-'); //以"-"區分種類及次號
            string acmWordNumOut = strs[0]; //動支編號(8碼)
            string acmKind = null; //種類
            switch (strs[1])
            {
                case "1":
                    acmKind = "預付";
                    break;
                case "2":
                    acmKind = "核銷";
                    break;
                case "3":
                    acmKind = "估列";
                    break;
                case "4":
                    acmKind = "估列收回";
                    break;
                case "5":
                    acmKind = "預撥收回";
                    break;
                case "6":
                    acmKind = "核銷收回";
                    break;
                default:
                    acmKind = "無";
                    break;
            }
            string acmNo = strs[2]; //次別           

            if (int.Parse(acmWordNumOut.Substring(0, 3)) < DateTime.Now.Year - 1911)
            {
                int isOrigNum = dao.GetGBCVisaDetailAbateDetail(x => x.PK_會計年度 == fillVouScript.傳票年度 && x.F_原動支編號 == acmWordNumOut).Count();
                if (isOrigNum > 0)
                {
                    acmWordNumOut = (dao.GetGBCVisaDetailAbateDetail(x => x.基金代碼 == fundNo && x.F_傳票年度 == fillVouScript.傳票年度 && x.F_原動支編號 == acmWordNumOut))
                                .FirstOrDefault().PK_動支編號;
                }
                
            }

            //isPass = jsonDAO.IsPass(fundNo, acmWordNumOut.Substring(0, 3), acmWordNumOut, acmKind, acmNo);
            //isJSON2 = jsonDAO.FindJSON2(fundNo, acmWordNumOut.Substring(0, 3), acmWordNumOut, acmKind, acmNo);

            isPass = jsonDAO.IsPass(x => x.基金代碼 == fundNo && x.PFK_會計年度 == fillVouScript.傳票年度 && x.PFK_動支編號 == acmWordNumOut && x.PFK_種類 == acmKind && x.PFK_次別 == acmNo);
            isJSON2 = jsonDAO.FindJSON2(x => x.基金代碼 == fundNo && x.PFK_會計年度 == fillVouScript.傳票年度 && x.PFK_動支編號 == acmWordNumOut && x.PFK_種類 == acmKind && x.PFK_次別 == acmNo);

            gbcVisaDetailAbateDetail.基金代碼 = fundNo;
            gbcVisaDetailAbateDetail.PK_會計年度 = fillVouScript.傳票年度;
            gbcVisaDetailAbateDetail.PK_動支編號 = acmWordNumOut;
            gbcVisaDetailAbateDetail.PK_種類 = acmKind;
            gbcVisaDetailAbateDetail.PK_次別 = acmNo;
            gbcVisaDetailAbateDetail.F_傳票年度 = fillVouScript.傳票年度;
            gbcVisaDetailAbateDetail.F_傳票號1 = fillVouScript.傳票號;
            gbcVisaDetailAbateDetail.F_製票日期1 = fillVouScript.製票日期;

            foreach (var 傳票明細Item in fillVouScript.傳票明細)
            {
                isVouNo1 = dao.IsVouNo1(fundNo, fillVouScript.傳票年度, acmWordNumOut, acmKind, acmNo, 傳票明細Item.明細號);
                gbcVisaDetailAbateDetail.PK_明細號 = 傳票明細Item.明細號;
                gbcVisaDetailAbateDetail.F_傳票明細號1 = int.Parse(傳票明細Item.傳票明細號);

                if ((isVouNo1 == null) && (isPass == "0")) //傳票1未回填 AND 未結案 --回填至傳票1
                {
                    dao.UpdateVouNo1(gbcVisaDetailAbateDetail);
                    count++;
                    if ((isJSON2.Trim().Length == 0) && (count == fillVouScript.傳票明細.Count) && (傳票明細Item.明細號.Trim().Length > 1))
                    {
                        jsonDAO.UpdatePassFlg(fundNo, fillVouScript.傳票年度, acmWordNumOut, acmKind, acmNo);
                    }
                }
                else if ((isVouNo1 != null) && (isPass == "0"))//傳票1已回填 AND 未結案 --回填至傳票2
                {
                    dao.UpdateVouNo2(gbcVisaDetailAbateDetail);
                    isVouNo2 = dao.IsVouNo2(fundNo, fillVouScript.傳票年度, acmWordNumOut, acmKind, acmNo, 傳票明細Item.明細號);
                    if (isVouNo2 != null)
                    {
                        jsonDAO.UpdatePassFlg(fundNo, fillVouScript.傳票年度, acmWordNumOut, acmKind, acmNo);
                    }                    
                }
                else
                {
                    return acmWordNumOut + "-" + acmKind + "-" + acmNo + "...回填失敗!  請確認是否已回填完畢。";
                }

                #region 傳票號回寫至預控系統
                //傳票號回寫至預控系統
                //由Web.Config來開關是否回填至預控系統
                string isFillToGBC = WebConfigurationManager.AppSettings["isFillToGBC"];

                if ((isFillToGBC.Trim()).Equals("1"))
                {
                    //判斷基金代號,回填至對應的預控系統(GBC)
                    if (gbcVisaDetailAbateDetail.基金代碼 == "010")//醫發服務參考
                    {
                        GBCWebService.GBCWebService ws = new GBCWebService.GBCWebService();
                        ws.FillVouNo(gbcVisaDetailAbateDetail.PK_會計年度, gbcVisaDetailAbateDetail.PK_動支編號, gbcVisaDetailAbateDetail.PK_種類, gbcVisaDetailAbateDetail.PK_次別, gbcVisaDetailAbateDetail.PK_明細號, gbcVisaDetailAbateDetail.F_傳票號1, gbcVisaDetailAbateDetail.F_製票日期1, gbcVisaDetailAbateDetail.F_傳票號1, gbcVisaDetailAbateDetail.F_製票日期1);
                    }
                    else if (gbcVisaDetailAbateDetail.基金代碼 == "040")//菸害****尚未加入服務參考****
                    {

                    }
                    else if (gbcVisaDetailAbateDetail.基金代碼 == "090")//家防服務參考
                    {
                        DVGBCWebService.GBCWebService ws = new DVGBCWebService.GBCWebService();
                        ws.FillVouNo(gbcVisaDetailAbateDetail.PK_會計年度, gbcVisaDetailAbateDetail.PK_動支編號, gbcVisaDetailAbateDetail.PK_種類, gbcVisaDetailAbateDetail.PK_次別, gbcVisaDetailAbateDetail.PK_明細號, gbcVisaDetailAbateDetail.F_傳票號1, gbcVisaDetailAbateDetail.F_製票日期1, gbcVisaDetailAbateDetail.F_傳票號1, gbcVisaDetailAbateDetail.F_製票日期1);
                    }
                    else if (gbcVisaDetailAbateDetail.基金代碼 == "100")//長照
                    {
                        LCGBCWebService.GBCWebService ws = new LCGBCWebService.GBCWebService();
                        ws.FillVouNo(gbcVisaDetailAbateDetail.PK_會計年度, gbcVisaDetailAbateDetail.PK_動支編號, gbcVisaDetailAbateDetail.PK_種類, gbcVisaDetailAbateDetail.PK_次別, gbcVisaDetailAbateDetail.PK_明細號, gbcVisaDetailAbateDetail.F_傳票號1, gbcVisaDetailAbateDetail.F_製票日期1, gbcVisaDetailAbateDetail.F_傳票號1, gbcVisaDetailAbateDetail.F_製票日期1);
                    }
                    else if (gbcVisaDetailAbateDetail.基金代碼 == "110")//生產
                    {
                        BAGBCWebService.GBCWebService ws = new BAGBCWebService.GBCWebService();
                        ws.FillVouNo(gbcVisaDetailAbateDetail.PK_會計年度, gbcVisaDetailAbateDetail.PK_動支編號, gbcVisaDetailAbateDetail.PK_種類, gbcVisaDetailAbateDetail.PK_次別, gbcVisaDetailAbateDetail.PK_明細號, gbcVisaDetailAbateDetail.F_傳票號1, gbcVisaDetailAbateDetail.F_製票日期1, gbcVisaDetailAbateDetail.F_傳票號1, gbcVisaDetailAbateDetail.F_製票日期1);
                    }
                }
                #endregion
                
            }

            return "回填完畢";
        }

        #region 手動搜尋功能
        //==================手動搜尋功能===================
        [WebMethod]
        /// <summary>
        /// 找年度
        /// </summary>
        /// <param name="fundNo"></param>
        /// <returns></returns>
        public List<string> GetYear(string fundNo)
        {
            ServicePointManager.ServerCertificateValidationCallback = new RemoteCertificateValidationCallback(ValidateServerCertificate);
            //先判斷基金代號
            if (fundNo == "010")//醫發服務參考
            {
                GBCWebService.GBCWebService ws = new GBCWebService.GBCWebService();
                List<string> yearList = new List<string>(ws.GetYear());

                return yearList;
            }
            else if (fundNo == "040")//菸害****尚未加入服務參考****
            {
                return null;
            }
            else if (fundNo == "090")//家防服務參考
            {
                DVGBCWebService.GBCWebService ws = new DVGBCWebService.GBCWebService();
                List<string> yearList = new List<string>(ws.GetYear());

                return yearList;
            }
            else if (fundNo == "100")//長照****尚未加入服務參考****
            {
                LCGBCWebService.GBCWebService ws = new LCGBCWebService.GBCWebService();
                List<string> yearList = new List<string>(ws.GetYear());

                return yearList;
            }
            else if (fundNo == "110")//生產****尚未加入服務參考****
            {
                BAGBCWebService.GBCWebService ws = new BAGBCWebService.GBCWebService();
                List<string> yearList = new List<string>(ws.GetYear());

                return yearList;
            }
            else
            {
                return null;
            }
        }

        [WebMethod]
        /// <summary>
        /// 找動支編號
        /// </summary>
        /// <param name="fundNo"></param>
        /// <param name="accYear"></param>
        /// <returns></returns>
        public List<string> GetAcmWordNum(string fundNo, string accYear)
        {
            ServicePointManager.ServerCertificateValidationCallback = new RemoteCertificateValidationCallback(ValidateServerCertificate);
            //先判斷基金代號
            if (fundNo == "010")//醫發服務參考
            {
                GBCWebService.GBCWebService ws = new GBCWebService.GBCWebService();
                List<string> acmNoList = new List<string>(
                    ws.GetAcmWordNum(accYear));

                return acmNoList;
            }
            else if (fundNo == "040")//菸害****尚未加入服務參考****
            {
                return null;
            }
            else if (fundNo == "090")//家防服務參考
            {
                DVGBCWebService.GBCWebService ws = new DVGBCWebService.GBCWebService();
                List<string> acmNoList = new List<string>(
                    ws.GetAcmWordNum(accYear));

                return acmNoList;
            }
            else if (fundNo == "100")//長照****尚未加入服務參考****
            {
                LCGBCWebService.GBCWebService ws = new LCGBCWebService.GBCWebService();
                List<string> acmNoList = new List<string>(
                    ws.GetAcmWordNum(accYear));

                return acmNoList;
            }
            else if (fundNo == "110")//生產****尚未加入服務參考****
            {
                BAGBCWebService.GBCWebService ws = new BAGBCWebService.GBCWebService();
                List<string> acmNoList = new List<string>(
                    ws.GetAcmWordNum(accYear));

                return acmNoList;
            }
            else
            {
                return null;
            }
        }

        [WebMethod]
        /// <summary>
        /// 找種類
        /// </summary>
        /// <param name="fundNo"></param>
        /// <param name="accYear"></param>
        /// <param name="acmWordNum"></param>
        /// <returns></returns>
        public List<string> GetAccKind(string fundNo, string accYear, string acmWordNum)
        {
            //先判斷基金代號
            if (fundNo == "010")//醫發服務參考
            {
                GBCWebService.GBCWebService ws = new GBCWebService.GBCWebService();
                List<string> accKindList = new List<string>(
                    ws.GetAccKind(accYear, acmWordNum));
                return accKindList;
            }
            else if (fundNo == "040")//菸害****尚未加入服務參考****
            {
                return null;
            }
            else if (fundNo == "090")//家防服務參考
            {
                DVGBCWebService.GBCWebService ws = new DVGBCWebService.GBCWebService();
                List<string> accKindList = new List<string>(
                    ws.GetAccKind(accYear, acmWordNum));
                return accKindList;
            }
            else if (fundNo == "100")//長照****尚未加入服務參考****
            {
                LCGBCWebService.GBCWebService ws = new LCGBCWebService.GBCWebService();
                List<string> accKindList = new List<string>(
                    ws.GetAccKind(accYear, acmWordNum));
                return accKindList;
            }
            else if (fundNo == "110")//生產****尚未加入服務參考****
            {
                BAGBCWebService.GBCWebService ws = new BAGBCWebService.GBCWebService();
                List<string> accKindList = new List<string>(
                    ws.GetAccKind(accYear, acmWordNum));
                return accKindList;
            }
            else
            {
                return null;
            }
        }

        [WebMethod]
        /// <summary>
        /// 找次數
        /// </summary>
        /// <param name="fundNo"></param>
        /// <param name="accYear"></param>
        /// <param name="acmWordNum"></param>
        /// <param name="accKind"></param>
        /// <returns></returns>
        public List<string> GetAccCount(string fundNo, string accYear, string acmWordNum, string accKind)
        {
            ServicePointManager.ServerCertificateValidationCallback = new RemoteCertificateValidationCallback(ValidateServerCertificate);
            //先判斷基金代號
            if (fundNo == "010")//醫發服務參考
            {
                GBCWebService.GBCWebService ws = new GBCWebService.GBCWebService();
                List<string> accDetailList = new List<string>(
                    ws.GetAccCount(accYear, acmWordNum, accKind));
                return accDetailList;
            }
            else if (fundNo == "040")//菸害****尚未加入服務參考****
            {
                return null;
            }
            else if (fundNo == "090")//家防服務參考
            {
                DVGBCWebService.GBCWebService ws = new DVGBCWebService.GBCWebService();
                List<string> accDetailList = new List<string>(
                    ws.GetAccCount(accYear, acmWordNum, accKind));
                return accDetailList;
            }
            else if (fundNo == "100")//長照****尚未加入服務參考****
            {
                LCGBCWebService.GBCWebService ws = new LCGBCWebService.GBCWebService();
                List<string> accDetailList = new List<string>(
                    ws.GetAccCount(accYear, acmWordNum, accKind));
                return accDetailList;
            }
            else if (fundNo == "110")//生產****尚未加入服務參考****
            {
                BAGBCWebService.GBCWebService ws = new BAGBCWebService.GBCWebService();
                List<string> accDetailList = new List<string>(
                    ws.GetAccCount(accYear, acmWordNum, accKind));
                return accDetailList;
            }
            else
            {
                return null;
            }
        }

        [WebMethod]
        /// <summary>
        /// 找動支明細號
        /// </summary>
        /// <param name="fundNo"></param>
        /// <param name="accYear"></param>
        /// <param name="acmWordNum"></param>
        /// <param name="accKind"></param>
        /// <param name="accCount"></param>
        /// <returns></returns>
        public List<string> GetAccDetail(string fundNo, string accYear, string acmWordNum, string accKind, string accCount)
        {
            ServicePointManager.ServerCertificateValidationCallback = new RemoteCertificateValidationCallback(ValidateServerCertificate);
            //先判斷基金代號
            if (fundNo == "010")//醫發服務參考
            {
                GBCWebService.GBCWebService ws = new GBCWebService.GBCWebService();
                List<string> accDetailList = new List<string>(
                    ws.GetAccDetail(accYear, acmWordNum, accKind, accCount));
                return accDetailList;
            }
            else if (fundNo == "040")//菸害****尚未加入服務參考****
            {
                return null;
            }
            else if (fundNo == "090")//家防服務參考
            {
                DVGBCWebService.GBCWebService ws = new DVGBCWebService.GBCWebService();
                List<string> accDetailList = new List<string>(
                    ws.GetAccDetail(accYear, acmWordNum, accKind, accCount));
                return accDetailList;
            }
            else if (fundNo == "100")//長照****尚未加入服務參考****
            {
                LCGBCWebService.GBCWebService ws = new LCGBCWebService.GBCWebService();
                List<string> accDetailList = new List<string>(
                    ws.GetAccDetail(accYear, acmWordNum, accKind, accCount));
                return accDetailList;
            }
            else if (fundNo == "110")//生產****尚未加入服務參考****
            {
                BAGBCWebService.GBCWebService ws = new BAGBCWebService.GBCWebService();
                List<string> accDetailList = new List<string>(
                    ws.GetAccDetail(accYear, acmWordNum, accKind, accCount));
                return accDetailList;
            }
            else
            {
                return null;
            }
        }

        [WebMethod]
        /// <summary>
        /// 根據所有key搜尋
        /// </summary>
        /// <param name="fundNo"></param>
        /// <param name="accYear"></param>
        /// <param name="acmWordNum"></param>
        /// <param name="accKind"></param>
        /// <param name="accCount"></param>
        /// <param name="accDetail"></param>
        /// <returns></returns>
        public string GetByPrimaryKey(string fundNo, string accYear, string acmWordNum, string accKind, string accCount, string accDetail)
        {
            ServicePointManager.ServerCertificateValidationCallback = new RemoteCertificateValidationCallback(ValidateServerCertificate);
            //先判斷基金代號
            if (fundNo == "010")//醫發服務參考
            {
                GBCWebService.GBCWebService ws = new GBCWebService.GBCWebService();
                string getGBCVisaDetail = ws.GetByPrimaryKey(accYear, acmWordNum, accKind, accCount, accDetail);
                return getGBCVisaDetail;
            }
            else if (fundNo == "040")//菸害****尚未加入服務參考****
            {
                return null;
            }
            else if (fundNo == "090")//家防服務參考
            {
                DVGBCWebService.GBCWebService ws = new DVGBCWebService.GBCWebService();
                string getGBCVisaDetail = ws.GetByPrimaryKey(accYear, acmWordNum, accKind, accCount, accDetail);
                return getGBCVisaDetail;
            }
            else if (fundNo == "100")//長照****尚未加入服務參考****
            {
                LCGBCWebService.GBCWebService ws = new LCGBCWebService.GBCWebService();
                string getGBCVisaDetail = ws.GetByPrimaryKey(accYear, acmWordNum, accKind, accCount, accDetail);
                return getGBCVisaDetail;
            }
            else if (fundNo == "110")//生產****尚未加入服務參考****
            {
                BAGBCWebService.GBCWebService ws = new BAGBCWebService.GBCWebService();
                string getGBCVisaDetail = ws.GetByPrimaryKey(accYear, acmWordNum, accKind, accCount, accDetail);
                return getGBCVisaDetail;
            }
            else
            {
                return null;
            }
        }

        [WebMethod]
        /// <summary>
        /// 找整批估列
        /// </summary>
        /// <param name="fundNo"></param>
        /// <param name="accYear"></param>
        /// <param name="accKind"></param>
        /// <param name="batch"></param>
        /// <returns></returns>
        public List<string> GetByKind(string fundNo, string accYear, string accKind, string batch)
        {
            ServicePointManager.ServerCertificateValidationCallback = new RemoteCertificateValidationCallback(ValidateServerCertificate);
            switch (batch)
            {
                case "6月":
                    batch = "1";
                    break;
                case "12月":
                    batch = "2";
                    break;
                default:
                    batch = "";
                    break;
            }
            //先判斷基金代號
            if (fundNo == "010")//醫發服務參考
            {
                GBCWebService.GBCWebService ws = new GBCWebService.GBCWebService();
                List<string> accDetailList = new List<string>(ws.GetByKind(accYear, accKind, batch));
                List<string> resultList = new List<string>();

                foreach (var accDetailListItem in accDetailList)
                {
                    resultList.Add(GetVw_GBCVisaDetail(fundNo, accDetailListItem));
                }

                return resultList;
            }
            else if (fundNo == "040")//菸害****尚未加入服務參考****
            {
                return null;
            }
            else if (fundNo == "090")//家防服務參考
            {
                DVGBCWebService.GBCWebService ws = new DVGBCWebService.GBCWebService();
                List<string> accDetailList = new List<string>(
                            ws.GetByKind(accYear, accKind, batch));

                return accDetailList;
            }
            else if (fundNo == "100")//長照****尚未加入服務參考****
            {
                LCGBCWebService.GBCWebService ws = new LCGBCWebService.GBCWebService();
                List<string> accDetailList = new List<string>(
                            ws.GetByKind(accYear, accKind, batch));

                return accDetailList;
            }
            else if (fundNo == "110")//生產****尚未加入服務參考****
            {
                BAGBCWebService.GBCWebService ws = new BAGBCWebService.GBCWebService();
                List<string> accDetailList = new List<string>(
                            ws.GetByKind(accYear, accKind, batch));

                return accDetailList;
            }
            else
            {
                return null;
            }
        }
        #endregion
    }
}
