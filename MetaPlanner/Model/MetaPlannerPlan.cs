using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MetaPlanner.Model
{
    public class MetaPlannerPlan
    {
        public string PlanId { get; set; }
        public string PlanName { get; set; }
        public string CreatedBy { get; set; }
        public string CreatedDate { get; set; }
        public string Owner { get; set; }
        public string Url { get; set; }
    }
}
