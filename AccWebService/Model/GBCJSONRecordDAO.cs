using AccWebService.GBCWebService;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using AccWebService.EF;
using System.Linq.Expressions;


namespace AccWebService.Model
{
    public class GBCJSONRecordDAO
    {
        NPSFEntities db = new NPSFEntities();

        public IQueryable<GBCJSONRecord> GetGBCJSONRecord(Expression<Func<GBCJSONRecord, bool>> condition)
        {
            var result = from s1 in db.GBCJSONRecord select s1;

            return result.Where(condition);
        }

        public string IsPass(Expression<Func<GBCJSONRecord, bool>> condition)
        {
            var result = (from res in db.GBCJSONRecord select res).Where(condition);
            var isPass = (from res in result select res.是否結案).FirstOrDefault();

            if (isPass == null)
            {
                isPass = "0";
            }

            return isPass;
        }

        public void DeleteJsonRecord1(Vw_GBCVisaDetail vw_GBCVisaDetail)
        {
            var result = (from s1 in db.GBCJSONRecord select s1)
            .Where(x => x.基金代碼 == vw_GBCVisaDetail.基金代碼 && x.PFK_會計年度 == vw_GBCVisaDetail.PK_會計年度 && x.PFK_動支編號 == vw_GBCVisaDetail.PK_動支編號 && x.PFK_種類 == vw_GBCVisaDetail.PK_種類 && x.PFK_次別 == vw_GBCVisaDetail.PK_次別)
            .ToList();

            foreach (var resultItem in result)
            {
                db.GBCJSONRecord.Remove(resultItem);
            }
            
            //db.Entry(result).State = System.Data.EntityState.Deleted;
            db.SaveChanges();
        }

        public void InsertJsonRecord1(Vw_GBCVisaDetail vw_GBCVisaDetail, string vouJoson)
        {
            GBCJSONRecord gbcJSONRecord = new GBCJSONRecord();
            gbcJSONRecord.基金代碼 = vw_GBCVisaDetail.基金代碼;            
            gbcJSONRecord.PFK_會計年度 = vw_GBCVisaDetail.PK_會計年度;
            gbcJSONRecord.PFK_動支編號 = vw_GBCVisaDetail.PK_動支編號;
            gbcJSONRecord.PFK_種類 = vw_GBCVisaDetail.PK_種類;
            gbcJSONRecord.PFK_次別 = vw_GBCVisaDetail.PK_次別;
            gbcJSONRecord.傳票JSON1 = vouJoson;

            db.GBCJSONRecord.Add(gbcJSONRecord);
            db.SaveChanges();
        }

        public void InsertJsonRecord2(Vw_GBCVisaDetail vw_GBCVisaDetail, string vouJoson)
        {
            var gbcJSONRecord = (from s1 in db.GBCJSONRecord select s1)
                .Where(x => x.基金代碼 == vw_GBCVisaDetail.基金代碼 && x.PFK_會計年度 == vw_GBCVisaDetail.PK_會計年度 && x.PFK_動支編號 == vw_GBCVisaDetail.PK_動支編號 && x.PFK_種類 == vw_GBCVisaDetail.PK_種類 && x.PFK_次別 == vw_GBCVisaDetail.PK_次別)
                .FirstOrDefault();

            gbcJSONRecord.基金代碼 = vw_GBCVisaDetail.基金代碼;
            gbcJSONRecord.PFK_會計年度 = vw_GBCVisaDetail.PK_會計年度;
            gbcJSONRecord.PFK_動支編號 = vw_GBCVisaDetail.PK_動支編號;
            gbcJSONRecord.PFK_種類 = vw_GBCVisaDetail.PK_種類;
            gbcJSONRecord.PFK_次別 = vw_GBCVisaDetail.PK_次別;
            gbcJSONRecord.傳票JSON2 = vouJoson;

            db.SaveChanges();
        }

        public string FindJSON1(Expression<Func<GBCJSONRecord, bool>> condition)
        {
            var result = (from s1 in db.GBCJSONRecord select s1).Where(condition);
            var json1 = (from s1 in result select s1.傳票JSON1).FirstOrDefault();

            if (json1 == null)
            {
                json1 = "";
            }

            return json1;
        }

        public string FindJSON2(Expression<Func<GBCJSONRecord, bool>> condition)
        {
            var result = (from s1 in db.GBCJSONRecord select s1).Where(condition);
            var json2 = (from s1 in result select s1.傳票JSON2).FirstOrDefault();

            if (json2 == null)
            {
                json2 = "";
            }

            return json2;
        }

        public void InsertJsonLog(string fundNo, string acmWordNum, string vouNoJSON)
        {
            GBCJSONRecordLog gbcJSONRecord = new GBCJSONRecordLog();
            gbcJSONRecord.基金代碼 = fundNo;
            gbcJSONRecord.條碼 = acmWordNum;
            gbcJSONRecord.JSON紀錄 = vouNoJSON;
            gbcJSONRecord.接收時間 = DateTime.Now;

            db.GBCJSONRecordLog.Add(gbcJSONRecord);
            db.SaveChanges();
        }

        /// <summary>
        /// 標上結案旗標
        /// </summary>
        /// <param name="基金代碼"></param>
        /// <param name="會計年度"></param>
        /// <param name="動支編號"></param>
        /// <param name="種類"></param>
        /// <param name="次別"></param>
        public void UpdatePassFlg(string 基金代碼, string 會計年度, string 動支編號, string 種類, string 次別)
        {
            var getOne = (from s1 in db.GBCJSONRecord select s1)
                .Where(x => x.基金代碼 == 基金代碼 && x.PFK_會計年度 == 會計年度 && x.PFK_動支編號 == 動支編號 && x.PFK_種類 == 種類 && x.PFK_次別 == 次別)
                .FirstOrDefault();

            getOne.是否結案 = "1";

            db.SaveChanges();
        }
    }
}