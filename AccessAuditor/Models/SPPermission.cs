using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AccessAuditor.Models
{
    public class SPPermission
    {
        public SPPermission()
        {
            //RoleBindings = new List<string>();
        }
        public string Member { get; set; }

        public string Permissions { get; set; }
        //public IEnumerable<string> RoleBindings { get; set; }

    }
}
