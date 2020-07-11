using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MetaPlanner.Model
{
    public class Plan
    {
        public string PlanId { get; set; }
        public string PlanName { get; set; }
       // public string MailAddress { get; set; }
        //public string Description { get; set; }
       // public bool IsPublic { get; set; }
       // public string GroupName { get; set; }
        //public string GroupId { get; set; }
        public string CreatedBy { get; set; }
        public string CreatedDate { get; set; }
        public string Owner { get; set; }
       // public string SourceAppName { get; set; }
        //public string SourceAppId { get; set; }
        //public int TotalTasks { get; set; }
        //public int TotalBuckets { get; set; }
    }
}
