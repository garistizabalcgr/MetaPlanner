using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MetaPlanner.Model
{
    class MetaPlannerBucket
    {
        public string PlanId { get; set; }
        public string BucketId { get; set; }
        public string BucketName { get; set; }
        public string OrderHint { get; set; }
    }
}
