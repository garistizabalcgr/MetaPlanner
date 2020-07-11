using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MetaPlanner.Model
{
    class Bucket
    {
        public string BucketId { get; set; }
        public string BucketName { get; set; }
        public string OrderHint { get; set; }

        public string PlanId { get; set; }
    }
}
