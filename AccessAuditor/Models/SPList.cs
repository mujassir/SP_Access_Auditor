using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AccessAuditor.Models
{
    public class SPList
    {
        public string Title { get; set; }

        public IEnumerable<SPPermission> Permissions { get; set; }


    }
}
