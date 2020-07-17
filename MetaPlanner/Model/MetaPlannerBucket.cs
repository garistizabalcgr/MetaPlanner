using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MetaPlanner.Model
{
    class MetaPlannerBucket
    {
        public MetaPlannerBucket()
        {
        }

        public MetaPlannerBucket(IDictionary<string, object> fields)
        {
            fields.TryGetValue("Title", out bucketId);
            fields.TryGetValue("BucketName", out bucketName);
            fields.TryGetValue("PlanId", out planId);
            fields.TryGetValue("OrderHint", out orderHint);
        }

        private object planId;
        public string PlanId
        {
            get { return planId.ToString(); }
            set { planId = value; }
        }

        private object bucketId;
        public string BucketId
        {
            get { return bucketId.ToString(); }
            set { bucketId = value; }
        }

        private object bucketName;
        public string BucketName
        {
            get { return bucketName.ToString(); }
            set { bucketName = value; }
        }

        private object orderHint;
        public string OrderHint
        {
            get { return orderHint.ToString(); }
            set { orderHint = value; }
        }
    }
}
