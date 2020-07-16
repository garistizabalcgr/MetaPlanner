using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MetaPlanner.Model
{
    class MetaPlannerUser
    {
        public string UserId { get; set; }
        public string UserPrincipalName { get; set; }
        public string Mail { get; set; }
        public string DisplayName { get; set; }
        public string Department { get; set; }
        public string JobTitle { get; set; }
        public string ManagerMail { get; set; }
        public string ManagerUPN { get; set; }

    }
}
