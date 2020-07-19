using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MetaPlanner.Model
{
    class MetaPlannerAssignment
    {


        public MetaPlannerAssignment()
        {
        }

        public MetaPlannerAssignment(IDictionary<string, object> fields)
        {
            fields.TryGetValue("TaskId", out taskId);
            fields.TryGetValue("UserId", out userId);
        }

        private object taskId;
        public string TaskId
        {
            get { return taskId.ToString(); }
            set { taskId = value; }
        }

        private object userId;
        public string UserId
        {
            get { return userId.ToString(); }
            set { userId = value; }
        }

    }
}
