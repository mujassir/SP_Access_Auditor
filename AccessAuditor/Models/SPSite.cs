using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AccessAuditor.Models
{
    public class SPSite
    {

        public SPSite()
        {
            Sites = new List<SPSite>();
            Lists = new List<SPList>();
            Permissions = new List<SPPermission>();
        }
        public string Title { get; set; }

        public IEnumerable<SPSite> Sites { get; set; }

        public IEnumerable<SPList> Lists { get; set; }

        public IEnumerable<SPPermission> Permissions { get; set; }


    }
}
