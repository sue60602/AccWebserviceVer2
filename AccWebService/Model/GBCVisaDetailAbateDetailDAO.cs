using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using AccWebService.EF;
using System.Linq.Expressions;

namespace AccWebService.Model
{
    public class GBCVisaDetailAbateDetailDAO
    {
        NPSFEntities db = new NPSFEntities();

        public int FindLog(Expression<Func<GBCVisaDetailAbateDetail, bool>> condition)
        {
            var result = from s1 in db.GBCVisaDetailAbateDetail select s1;
            
            return result.Where(condition).Count();
        }

        public void Update(Vw_GBCVisaDetail vw_GBCVisaDetail)
        {
            var result = (from s1 in db.GBCVisaDetailAbateDetail select s1)
                .Where(x => x.基金代碼 == vw_GBCVisaDetail.基金代碼 && x.PK_會計年度 == vw_GBCVisaDetail.PK_會計年度 && x.PK_動支編號 == vw_GBCVisaDetail.PK_動支編號 && x.PK_種類 == vw_GBCVisaDetail.PK_種類 && x.PK_次別 == vw_GBCVisaDetail.PK_次別 && x.PK_明細號 == vw_GBCVisaDetail.PK_明細號)
                .First();

            result.基金代碼 = vw_GBCVisaDetail.基金代碼;
            result.PK_會計年度 = vw_GBCVisaDetail.PK_會計年度;
            result.PK_動支編號 = vw_GBCVisaDetail.PK_動支編號;
            result.PK_種類 = vw_GBCVisaDetail.PK_種類;
            result.PK_次別 = vw_GBCVisaDetail.PK_次別;
            result.PK_明細號 = vw_GBCVisaDetail.PK_明細號;
            result.F_核定金額 = Convert.ToDecimal(vw_GBCVisaDetail.F_核定金額);
            result.F_受款人編號 = vw_GBCVisaDetail.F_受款人編號;
            result.F_受款人 = vw_GBCVisaDetail.F_受款人;
            result.F_原動支編號 = vw_GBCVisaDetail.F_原動支編號;

            db.SaveChanges();
            
        }

        public void Update(Vw_GBCVisaDetailForHPA vw_GBCVisaDetail)
        {
            var result = (from s1 in db.GBCVisaDetailAbateDetail select s1)
                .Where(x => x.基金代碼 == vw_GBCVisaDetail.基金代碼 && x.PK_會計年度 == vw_GBCVisaDetail.PK_會計年度 && x.PK_動支編號 == vw_GBCVisaDetail.PK_動支編號 && x.PK_種類 == vw_GBCVisaDetail.PK_種類 && x.PK_次別 == vw_GBCVisaDetail.PK_次別 && x.PK_明細號 == vw_GBCVisaDetail.PK_明細號)
                .First();

            result.基金代碼 = vw_GBCVisaDetail.基金代碼;
            result.PK_會計年度 = vw_GBCVisaDetail.PK_會計年度;
            result.PK_動支編號 = vw_GBCVisaDetail.PK_動支編號;
            result.PK_種類 = vw_GBCVisaDetail.PK_種類;
            result.PK_次別 = vw_GBCVisaDetail.PK_次別;
            result.PK_明細號 = vw_GBCVisaDetail.PK_明細號;
            result.F_核定金額 = Convert.ToDecimal(vw_GBCVisaDetail.F_核定金額);
            result.F_受款人編號 = vw_GBCVisaDetail.F_受款人編號;
            result.F_受款人 = vw_GBCVisaDetail.F_受款人;
            result.F_原動支編號 = vw_GBCVisaDetail.F_原動支編號;

            db.SaveChanges();

        }

        public void Insert(Vw_GBCVisaDetail vw_GBCVisaDetail)
        {
            GBCVisaDetailAbateDetail gbcVisaDetailAbateDetail = new GBCVisaDetailAbateDetail();
            gbcVisaDetailAbateDetail.基金代碼 = vw_GBCVisaDetail.基金代碼;
            gbcVisaDetailAbateDetail.PK_會計年度 = vw_GBCVisaDetail.PK_會計年度;
            gbcVisaDetailAbateDetail.PK_動支編號 = vw_GBCVisaDetail.PK_動支編號;
            gbcVisaDetailAbateDetail.PK_種類 = vw_GBCVisaDetail.PK_種類;
            gbcVisaDetailAbateDetail.PK_次別 = vw_GBCVisaDetail.PK_次別;
            gbcVisaDetailAbateDetail.PK_明細號 = vw_GBCVisaDetail.PK_明細號;
            gbcVisaDetailAbateDetail.F_核定金額 = Convert.ToDecimal(vw_GBCVisaDetail.F_核定金額);
            gbcVisaDetailAbateDetail.F_受款人編號 = vw_GBCVisaDetail.F_受款人編號;
            gbcVisaDetailAbateDetail.F_受款人 = vw_GBCVisaDetail.F_受款人;
            gbcVisaDetailAbateDetail.F_傳票年度 = "";
            gbcVisaDetailAbateDetail.F_原動支編號 = vw_GBCVisaDetail.F_原動支編號;

            db.GBCVisaDetailAbateDetail.Add(gbcVisaDetailAbateDetail);
            db.SaveChanges();
        }

        public void Insert(Vw_GBCVisaDetailForHPA vw_GBCVisaDetail)
        {
            GBCVisaDetailAbateDetail gbcVisaDetailAbateDetail = new GBCVisaDetailAbateDetail();
            gbcVisaDetailAbateDetail.基金代碼 = vw_GBCVisaDetail.基金代碼;
            gbcVisaDetailAbateDetail.PK_會計年度 = vw_GBCVisaDetail.PK_會計年度;
            gbcVisaDetailAbateDetail.PK_動支編號 = vw_GBCVisaDetail.PK_動支編號;
            gbcVisaDetailAbateDetail.PK_種類 = vw_GBCVisaDetail.PK_種類;
            gbcVisaDetailAbateDetail.PK_次別 = vw_GBCVisaDetail.PK_次別;
            gbcVisaDetailAbateDetail.PK_明細號 = vw_GBCVisaDetail.PK_明細號;
            gbcVisaDetailAbateDetail.F_核定金額 = Convert.ToDecimal(vw_GBCVisaDetail.F_核定金額);
            gbcVisaDetailAbateDetail.F_受款人編號 = vw_GBCVisaDetail.F_受款人編號;
            gbcVisaDetailAbateDetail.F_受款人 = vw_GBCVisaDetail.F_受款人;
            gbcVisaDetailAbateDetail.F_傳票年度 = "";
            gbcVisaDetailAbateDetail.F_原動支編號 = vw_GBCVisaDetail.F_原動支編號;

            db.GBCVisaDetailAbateDetail.Add(gbcVisaDetailAbateDetail);
            db.SaveChanges();
        }

        public IQueryable<GBCVisaDetailAbateDetail> GetGBCVisaDetailAbateDetail(Expression<Func<GBCVisaDetailAbateDetail, bool>> condition)
        {
            var result = from s1 in db.GBCVisaDetailAbateDetail select s1;

            return result.Where(condition);
        }

        public string IsVouNo1(string 基金代碼, string 會計年度, string 動支編號, string 種類, string 次別, string 明細號)
        {
            var getAll = GetGBCVisaDetailAbateDetail(x => x.基金代碼 == 基金代碼 && x.PK_會計年度 == 會計年度 && x.PK_動支編號 == 動支編號 && x.PK_種類 == 種類 && x.PK_次別 == 次別 && x.PK_明細號 == 明細號);
            var vouNo1 = (from s1 in getAll select s1.F_傳票號1).FirstOrDefault();

            return vouNo1;
        }

        public string IsVouNo2(string 基金代碼, string 會計年度, string 動支編號, string 種類, string 次別, string 明細號)
        {
            var getAll = GetGBCVisaDetailAbateDetail(x => x.基金代碼 == 基金代碼 && x.PK_會計年度 == 會計年度 && x.PK_動支編號 == 動支編號 && x.PK_種類 == 種類 && x.PK_次別 == 次別 && x.PK_明細號 == 明細號);
            var vouNo1 = (from s1 in getAll select s1.F_傳票號2).FirstOrDefault();

            return vouNo1;
        }

        public void UpdateVouNo1(GBCVisaDetailAbateDetail gbcVisaDetailAbateDetail)
        {
            var getOne = (from s1 in db.GBCVisaDetailAbateDetail select s1)
                .Where(x => x.基金代碼 == gbcVisaDetailAbateDetail.基金代碼 && x.PK_會計年度 == gbcVisaDetailAbateDetail.PK_會計年度 && x.PK_動支編號 == gbcVisaDetailAbateDetail.PK_動支編號 && x.PK_種類 == gbcVisaDetailAbateDetail.PK_種類 && x.PK_次別 == gbcVisaDetailAbateDetail.PK_次別 && x.PK_明細號 == gbcVisaDetailAbateDetail.PK_明細號)
                .FirstOrDefault();

            //先移除
            db.GBCVisaDetailAbateDetail.Remove(getOne);
            db.SaveChanges();

            //以插入方式更新
            getOne.F_傳票年度 = gbcVisaDetailAbateDetail.F_傳票年度;
            getOne.F_傳票號1 = gbcVisaDetailAbateDetail.F_傳票號1;
            getOne.F_傳票明細號1 = gbcVisaDetailAbateDetail.F_傳票明細號1;
            getOne.F_製票日期1 = gbcVisaDetailAbateDetail.F_製票日期1;

            db.GBCVisaDetailAbateDetail.Add(getOne);
            db.SaveChanges();


        }

        public void UpdateVouNo2(GBCVisaDetailAbateDetail gbcVisaDetailAbateDetail)
        {
            var getOne = (from s1 in db.GBCVisaDetailAbateDetail select s1)
                .Where(x => x.基金代碼 == gbcVisaDetailAbateDetail.基金代碼 && x.PK_會計年度 == gbcVisaDetailAbateDetail.PK_會計年度 && x.PK_動支編號 == gbcVisaDetailAbateDetail.PK_動支編號 && x.PK_種類 == gbcVisaDetailAbateDetail.PK_種類 && x.PK_次別 == gbcVisaDetailAbateDetail.PK_次別 && x.PK_明細號 == gbcVisaDetailAbateDetail.PK_明細號)
                .FirstOrDefault();
            //先移除
            db.GBCVisaDetailAbateDetail.Remove(getOne);
            db.SaveChanges();

            //以插入方式更新
            getOne.F_傳票年度 = gbcVisaDetailAbateDetail.F_傳票年度;
            getOne.F_傳票號2 = gbcVisaDetailAbateDetail.F_傳票號1;
            getOne.F_傳票明細號2 = gbcVisaDetailAbateDetail.F_傳票明細號1;
            getOne.F_製票日期2 = gbcVisaDetailAbateDetail.F_製票日期1;

            db.GBCVisaDetailAbateDetail.Add(getOne);
            db.SaveChanges();
        }

        public void FillVouNoForEstimate(string fundNo, string AccYear, string batch, string VouNo)
        {
            
            var getGBCVisaDetailAbateDetail = from s1 in db.GBCVisaDetailAbateDetail
                                              .Where(x => x.基金代碼 == fundNo && x.PK_會計年度 == AccYear && x.PK_種類=="估列" && x.PK_次別 == batch).ToList()
                                              select s1;

            foreach (var getGBCVisaDetailAbateDetailItem in getGBCVisaDetailAbateDetail)
            {
                db.GBCVisaDetailAbateDetail.Remove(getGBCVisaDetailAbateDetailItem);
                db.SaveChanges();

                getGBCVisaDetailAbateDetailItem.F_傳票年度 = AccYear;
                getGBCVisaDetailAbateDetailItem.F_傳票號1 = VouNo;
                db.GBCVisaDetailAbateDetail.Add(getGBCVisaDetailAbateDetailItem);

                db.SaveChanges();
            }
        }
    }
}