using AccWebService.EF;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Web;

namespace AccWebService.Model
{
    public class VouDetailDAO
    {
        NPSFEntities db = new NPSFEntities();
        public IQueryable<VouDetail> GetVouDetail(Expression<Func<VouDetail, bool>> condition)
        {
            var result = from s1 in db.VouDetail select s1;

            return result.Where(condition);
        }

    }
}