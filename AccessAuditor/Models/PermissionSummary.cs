using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AccessAuditor.Models
{
   public class PermissionSummary
    {
        public PermissionSummary() {
            ListCollection = new List<SPList>();
            SiteCollection = new List<SPSite>();
        }
        public string Member { get; set; }
        public int Sites { get; set; }
        public int Lists { get; set; }
        public int FullControl { get; set; }
        public int Edit { get; set; }
        public int Read { get; set; }
        public int LimitedAccess { get; set; }

        public List<SPList> ListCollection { get; set; }
        public List<SPSite> SiteCollection { get; set; }
    }
}
