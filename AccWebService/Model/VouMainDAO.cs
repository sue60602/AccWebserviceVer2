using AccWebService.EF;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Web;

namespace AccWebService.Model
{
    public class VouMainDAO
    {
        NPSFEntities db = new NPSFEntities();
        public IQueryable<VouMain> GetVouMain(Expression<Func<VouMain, bool>> condition)
        {
            var result = from s1 in db.VouMain select s1;

            return result.Where(condition);
        }
    }
}