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
        public string Email { get; set; }
        public string UserPrincipalName { get; set; }
        public string DisplayName { get; set; }
        public string ParentId { get; set; }
        public string ParentPrincipalName { get; set; }
    }
}
